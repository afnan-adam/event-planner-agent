require("dotenv").config();
const express = require("express");
const path = require("path");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const Anthropic = require("@anthropic-ai/sdk");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
} = require("docx");

// Fail fast if the API key is missing — avoids cryptic errors on first request.
if (!process.env.ANTHROPIC_API_KEY) {
  console.error("ERROR: ANTHROPIC_API_KEY is not set. Add it to your .env file and restart.");
  process.exit(1);
}

const app = express();
const PORT = process.env.PORT || 3000;

// ── Security headers ─────────────────────────────────────────────────────────
// helmet sets X-Content-Type-Options, X-Frame-Options, HSTS, etc. in one call.
// contentSecurityPolicy is relaxed slightly to allow inline scripts used by the
// single-page frontend (no external scripts are loaded, so the risk is low).
app.use(helmet({ contentSecurityPolicy: false }));

// ── CORS — same-origin only ───────────────────────────────────────────────────
// The frontend is served by this same Express process, so there is no legitimate
// cross-origin caller. Reject any preflight or credentialed cross-origin request.
app.use((req, res, next) => {
  const origin = req.get("Origin");
  if (origin) {
    const host = req.get("Host");
    const requestOrigin = new URL(origin);
    if (requestOrigin.host !== host) {
      return res.status(403).json({ error: "Cross-origin requests are not allowed." });
    }
  }
  next();
});

// ── Rate limiting ─────────────────────────────────────────────────────────────
// 20 requests per 15 minutes per IP is generous for a demo but prevents runaway
// spend if the API key is exposed or the app is crawled.
const apiLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: "Too many requests — please wait a few minutes and try again." },
});

app.use(express.json({ limit: "5mb" }));
app.use(express.static(path.join(__dirname, "public")));

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── Input validation helpers ──────────────────────────────────────────────────
const MAX_CONTEXT = 20_000;
const MAX_STR = 200;
const MAX_DOC_LEN = 30_000;
const MAX_DOCS = 3;

function isNonEmptyString(v) { return typeof v === "string" && v.trim().length > 0; }
function isBoundedString(v, max) { return typeof v === "string" && v.length <= max; }

function validateGenerate(body) {
  const { context, leadTime } = body;
  if (!isNonEmptyString(context)) return "context must be a non-empty string.";
  if (context.length > MAX_CONTEXT) return `context must be under ${MAX_CONTEXT} characters.`;
  if (leadTime !== undefined && !Number.isFinite(Number(leadTime)))
    return "leadTime must be a number.";
  return null;
}

function validateExportWord(body) {
  const { docs, eventName, org, date } = body;
  if (!Array.isArray(docs) || docs.length === 0) return "docs must be a non-empty array.";
  if (docs.length > MAX_DOCS) return `docs may contain at most ${MAX_DOCS} items.`;
  for (const d of docs) {
    if (typeof d !== "string") return "each doc must be a string.";
    if (d.length > MAX_DOC_LEN) return `each doc must be under ${MAX_DOC_LEN} characters.`;
  }
  if (eventName !== undefined && !isBoundedString(eventName, MAX_STR))
    return `eventName must be a string under ${MAX_STR} characters.`;
  if (org !== undefined && !isBoundedString(org, MAX_STR))
    return `org must be a string under ${MAX_STR} characters.`;
  if (date !== undefined && !isBoundedString(date, MAX_STR))
    return `date must be a string under ${MAX_STR} characters.`;
  return null;
}

// ── Generate documents ────────────────────────────────────────────────────────
// Calls Claude in parallel for three documents: volunteer brief, staffing plan,
// and day-of checklist. All three are returned in a single JSON response.
app.post("/api/generate", apiLimiter, async (req, res) => {
  const validationError = validateGenerate(req.body);
  if (validationError) return res.status(400).json({ error: validationError });

  const { context } = req.body;
  const leadTime = Number(req.body.leadTime) || 60;

  try {
    const [brief, staffing, checklist] = await Promise.all([
      anthropic.messages.create({
        model: "claude-sonnet-4-6",
        max_tokens: 2048,
        messages: [{
          role: "user",
          content: `You are an expert event operations coordinator. Using the event details below, produce a comprehensive Volunteer HQ Brief. Include these emoji-headed sections:\n\n📌 Room Assignments\n🧑‍💻 Technical Support Assignments (by room, using volunteer names and roles provided)\n📺 A/V & Tech Setup (setup owner, equipment, hybrid link if applicable)\n🗓 Run-of-Show Timeline (build a timed schedule using the agenda segments and start time)\n🏢 Overflow Room Plan (include contingency if overflow is not yet activated)\n📸 Recording & Photography (opt-out seating note, photo team duties, recommended shot list)\n🏷️ Check-In & Attendance (method, name tag details, tracker name and snapshot time)\n🪧 Signage & Materials (who places what, handouts)\n🔧 Room Setup & Arrival Plan (staggered arrival times per role, upon-arrival checklist)\n🤝 Special Guests (specific handling notes per guest)\n📱 On-Site Contact\n📣 Post-Event Actions (follow-up email, sharing, debrief)\n\nUse all names provided. Be specific and actionable. Output ONLY the document.\n\nEVENT DATA:\n${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),

      anthropic.messages.create({
        model: "claude-sonnet-4-6",
        max_tokens: 1500,
        messages: [{
          role: "user",
          content: `You are an expert event operations coordinator. Using the event details below, produce a clear Staffing Plan:\n\n1. Roster table — Name | Role | Room | Key Responsibilities\n2. Room-by-room breakdown with specific duties\n3. Escalation chain\n4. Staggered arrival timeline per person\n5. Contingency notes (overflow, A/V issues, late arrivals)\n\nOutput ONLY the document.\n\nEVENT DATA:\n${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),

      anthropic.messages.create({
        model: "claude-sonnet-4-6",
        max_tokens: 1500,
        messages: [{
          role: "user",
          content: `You are an expert event operations coordinator. Using the event details below, produce a Day-of Logistics Checklist organized by time block:\n\n- Pre-arrival (T-${leadTime}min)\n- Setup (T-30min)\n- Final sweep (T-10min)\n- Doors open\n- One section per agenda segment (by name)\n- Wrap-up\n- Post-event close-out\n\nEach item as [ ]. Include owner names where known. Output ONLY the checklist.\n\nEVENT DATA:\n${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),
    ]);

    res.json({ brief, staffing, checklist });
  } catch (e) {
    // Log the full error server-side; return only a generic message to the client
    // so internal details (model names, stack traces) are never exposed.
    console.error("[/api/generate]", e);
    res.status(500).json({ error: "Document generation failed. Please try again." });
  }
});

// ── Export to Word ────────────────────────────────────────────────────────────
// Converts the three Claude-generated markdown strings into a styled .docx file
// using the `docx` library, then streams it directly to the client.
app.post("/api/export-word", apiLimiter, async (req, res) => {
  const validationError = validateExportWord(req.body);
  if (validationError) return res.status(400).json({ error: validationError });

  const { docs, eventName, org, date } = req.body;

  const ACCENT = "1F4E79";   // dark navy
  const ACCENT2 = "2E74B5";  // mid blue
  const LIGHT = "D6E4F0";    // light blue tint

  const titles = ["Volunteer HQ Brief", "Staffing Plan", "Logistics Checklist"];

  function stripInline(text) {
    return text.replace(/\*\*([^*]+)\*\*/g, "$1").replace(/\*([^*]+)\*/g, "$1").replace(/`([^`]+)`/g, "$1");
  }

  function parseInline(text, baseSize = 20, baseColor = "2d2d2d") {
    const cleaned = text.replace(/`([^`]+)`/g, "$1").replace(/\*(?!\*)([^*]+)\*(?!\*)/g, "$1");
    const runs = [];
    cleaned.split(/(\*\*[^*]+\*\*)/).forEach(part => {
      const m = part.match(/^\*\*([^*]+)\*\*$/);
      if (m) runs.push(new TextRun({ text: m[1], bold: true, size: baseSize, color: baseColor }));
      else if (part) runs.push(new TextRun({ text: part, size: baseSize, color: baseColor }));
    });
    return runs.length ? runs : [new TextRun({ text: "", size: baseSize })];
  }

  function buildTable(lines) {
    const rows = lines.filter(l => !/^\|[-| :]+\|$/.test(l.trim()));
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: rows.map((line, ri) => {
        const cols = line.trim().split("|").filter((_, i, a) => i > 0 && i < a.length - 1).map(c => c.trim());
        const isHeader = ri === 0;
        return new TableRow({
          tableHeader: isHeader,
          children: cols.map(col => new TableCell({
            shading: isHeader ? { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" }
                              : (ri % 2 === 0 ? { fill: "F5F9FC", type: ShadingType.CLEAR, color: "auto" } : undefined),
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({
              children: [new TextRun({
                text: stripInline(col),
                bold: isHeader,
                color: isHeader ? "FFFFFF" : "2d2d2d",
                size: 18,
              })],
            })],
          })),
        });
      }),
    });
  }

  const sections = [];

  const titleChildren = [
    new Paragraph({
      children: [new TextRun({ text: eventName || "Event Brief", bold: true, size: 52, color: "FFFFFF" })],
      alignment: AlignmentType.CENTER,
      shading: { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" },
      spacing: { before: 480, after: 120 },
    }),
    new Paragraph({
      children: [new TextRun({ text: org || "", size: 26, color: "FFFFFF" })],
      alignment: AlignmentType.CENTER,
      shading: { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" },
      spacing: { after: 120 },
    }),
    new Paragraph({
      children: [new TextRun({ text: date || "", size: 22, color: "FFFFFF" })],
      alignment: AlignmentType.CENTER,
      shading: { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" },
      spacing: { after: 480 },
    }),
  ];
  sections.push({ children: titleChildren });

  docs.forEach((doc, di) => {
    const children = [];

    children.push(new Paragraph({
      children: [new TextRun({ text: titles[di].toUpperCase(), bold: true, size: 32, color: "FFFFFF" })],
      shading: { fill: ACCENT2, type: ShadingType.CLEAR, color: "auto" },
      spacing: { before: 0, after: 240 },
    }));

    const lines = doc.split("\n");
    let i = 0;
    let inCodeBlock = false;

    while (i < lines.length) {
      const line = lines[i];
      const t = line.trim();

      if (/^```/.test(t)) { inCodeBlock = !inCodeBlock; i++; continue; }
      if (inCodeBlock) {
        children.push(new Paragraph({
          children: [new TextRun({ text: line, size: 18, color: "444444", font: "Courier New" })],
          spacing: { before: 20, after: 20 },
          indent: { left: 360 },
        }));
        i++; continue;
      }

      if (/^\|/.test(t)) {
        const tableLines = [];
        while (i < lines.length && /^\|/.test(lines[i].trim())) { tableLines.push(lines[i]); i++; }
        if (tableLines.filter(l => !/^\|[-| :]+\|$/.test(l.trim())).length > 1) {
          children.push(buildTable(tableLines));
          children.push(new Paragraph({ children: [new TextRun({ text: "" })], spacing: { after: 120 } }));
        }
        continue;
      }

      if (/^\|[-| :]+\|$/.test(t) || /^---+$/.test(t)) { i++; continue; }

      if (/^# /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^# /, ""), bold: true, size: 28, color: ACCENT })],
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 320, after: 120 },
          border: { bottom: { color: LIGHT, size: 6, space: 1, style: BorderStyle.SINGLE } },
        }));
        i++; continue;
      }

      if (/^## /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^## /, ""), bold: true, size: 24, color: ACCENT })],
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 240, after: 100 },
        }));
        i++; continue;
      }

      if (/^### /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^### /, ""), bold: true, size: 22, color: ACCENT2 })],
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 180, after: 80 },
        }));
        i++; continue;
      }

      if (/^> /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: stripInline(t.replace(/^> /, "")), italics: true, size: 18, color: "555555" })],
          spacing: { before: 60, after: 60 },
          indent: { left: 360 },
          shading: { fill: "F5F9FC", type: ShadingType.CLEAR, color: "auto" },
        }));
        i++; continue;
      }

      if (/^[-*] /.test(t)) {
        children.push(new Paragraph({
          children: parseInline(t.replace(/^[-*] /, ""), 20, "2d2d2d"),
          bullet: { level: 0 },
          spacing: { before: 40, after: 40 },
        }));
        i++; continue;
      }

      if (/^[📌🧑📺🗓🏢📸🏷🪧🔧🤝📱📣]/.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t, bold: true, size: 24, color: ACCENT })],
          spacing: { before: 240, after: 120 },
          border: { bottom: { color: LIGHT, size: 4, space: 1, style: BorderStyle.SINGLE } },
        }));
        i++; continue;
      }

      if (!t) {
        children.push(new Paragraph({ children: [new TextRun({ text: "" })], spacing: { after: 80 } }));
        i++; continue;
      }

      children.push(new Paragraph({
        children: parseInline(t, 20, "2d2d2d"),
        spacing: { after: 80 },
      }));
      i++;
    }

    sections.push({ children });
  });

  try {
    const document = new Document({ sections });
    const buffer = await Packer.toBuffer(document);
    const filename = `${(eventName || "event").replace(/\s+/g, "_")}_brief.docx`;

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (e) {
    console.error("[/api/export-word]", e);
    res.status(500).json({ error: "Word export failed. Please try again." });
  }
});

app.listen(PORT, () => {
  console.log(`\n✅ Event Planning Agent running at http://localhost:${PORT}\n`);
});
