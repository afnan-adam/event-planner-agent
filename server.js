require("dotenv").config();
const express = require("express");
const path = require("path");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const Anthropic = require("@anthropic-ai/sdk");
const {
  Document, Packer, Paragraph, TextRun, ShadingType, BorderStyle,
  Table, TableRow, TableCell, WidthType,
} = require("docx");
const JSZip = require("jszip");

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
// Three documents in parallel: volunteer brief, staffing plan, logistics checklist.
// All three are previewed in tabs and exported as separate .docx files in a zip.
app.post("/api/generate", apiLimiter, async (req, res) => {
  const validationError = validateGenerate(req.body);
  if (validationError) return res.status(400).json({ error: validationError });

  const { context } = req.body;
  const leadTime = Number(req.body.leadTime) || 60;

  try {
    const [brief, staffing, checklist] = await Promise.all([
      // 1. Volunteer-facing brief — short, scannable, one page.
      anthropic.messages.create({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 800,
        messages: [{
          role: "user",
          content: `You are an event operations coordinator writing a one-page brief for day-of volunteers. Be concise — use short phrases, not full sentences. No tables, no equipment checklists, no escalation diagrams.

Produce exactly this structure and nothing else:

EVENT SUMMARY
[Event name] | [Date] | [Start time – End time] | [Primary room] | [Expected attendees] attendees
On-site contact: [Name] [Phone]
[One line only if hybrid: Virtual link: ...]

VOLUNTEER ROLES
[For each volunteer: Name (Role, Room) — arrive by [arrival time]. One-line responsibility.]

RUN OF SHOW
[For each agenda segment: Time — Segment name (duration)
  - Name: one short action phrase
  - Name: one short action phrase
  Only list volunteers who have something specific to do in that segment. Max 2-3 bullets per segment.]

KEY NOTES
[3 bullets max. Only genuinely important operational facts — e.g. who to call if A/V fails, opt-out names for photography, overflow room trigger.]

Output ONLY the document. No markdown headers with #. No extra sections.

EVENT DATA:
${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),

      // 2. Staffing plan — coordinator reference with roster and room duties.
      anthropic.messages.create({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 1000,
        messages: [{
          role: "user",
          content: `You are an expert event operations coordinator. Using the event details below, produce a concise Staffing Plan with these sections:

ROSTER
[Name | Role | Room | Arrival time | Key responsibility — one line each]

ROOM DUTIES
[For each room: room name, then bullet per volunteer with their specific duties]

ESCALATION
[Who to call first for: A/V issues, medical, overflow activation, late speaker]

CONTINGENCY
[3-4 bullets covering the most likely issues and the response]

Be specific. Use names from the data. No full-sentence essays. Output ONLY the document.

EVENT DATA:
${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),

      // 3. Day-of logistics checklist — time-blocked with owner names.
      anthropic.messages.create({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 800,
        messages: [{
          role: "user",
          content: `You are an expert event operations coordinator. Using the event details below, produce a Day-of Logistics Checklist organised by time block. Use [ ] for each item. Include the responsible person's name. Be brief — one line per item.

Sections:
T-${leadTime}min (Pre-arrival)
T-30min (Setup)
T-10min (Final sweep)
Doors open
[One section per agenda segment, named]
Wrap-up
Post-event close-out

Output ONLY the checklist. No prose.

EVENT DATA:
${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),
    ]);

    res.json({ brief, staffing, checklist });
  } catch (e) {
    console.error("[/api/generate]", e);
    res.status(500).json({ error: "Document generation failed. Please try again." });
  }
});

// ── Export to Word (zip of three .docx files) ────────────────────────────────
// Builds one .docx per document and streams them as a single .zip download.
app.post("/api/export-word", apiLimiter, async (req, res) => {
  const validationError = validateExportWord(req.body);
  if (validationError) return res.status(400).json({ error: validationError });

  const { docs, eventName, org, date } = req.body;

  const ACCENT = "1F4E79";  // dark navy
  const NAME   = "2E74B5";  // mid blue — volunteer names
  const LIGHT  = "D6E4F0";  // light blue rule lines

  const docTitles = ["Volunteer Brief", "Staffing Plan", "Logistics Checklist"];
  const fileNames = ["1_Volunteer_Brief", "2_Staffing_Plan", "3_Logistics_Checklist"];

  // Extract volunteer names from brief text (Name: pattern at bullet start)
  function extractNames(text) {
    const names = new Set();
    for (const line of text.split("\n")) {
      const m = line.match(/^[\s•\-*]*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?):/);
      if (m) names.add(m[1].trim());
    }
    return names;
  }

  // Split text into TextRuns, highlighting any known volunteer name in blue+bold
  function parseWithNames(text, names, baseSize = 18) {
    if (!names.size) return [new TextRun({ text, size: baseSize, color: "2d2d2d" })];
    const escaped = [...names].map(n => n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|");
    return text.split(new RegExp(`(${escaped})`, "g")).filter(p => p).map(p =>
      names.has(p)
        ? new TextRun({ text: p, bold: true, size: baseSize, color: NAME })
        : new TextRun({ text: p, size: baseSize, color: "2d2d2d" })
    );
  }

  // Convert consecutive markdown table lines into a styled Word table
  function buildTable(lines) {
    const dataRows = lines.filter(l => !/^\|[-| :]+\|$/.test(l.trim()));
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: dataRows.map((line, ri) => {
        const cols = line.trim().split("|").filter((_, i, a) => i > 0 && i < a.length - 1).map(c => c.trim());
        const isHeader = ri === 0;
        return new TableRow({
          tableHeader: isHeader,
          children: cols.map(col => new TableCell({
            shading: isHeader
              ? { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" }
              : (ri % 2 === 0 ? { fill: "F5F9FC", type: ShadingType.CLEAR, color: "auto" } : undefined),
            margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: [new Paragraph({
              children: [new TextRun({
                text: col.replace(/\*\*([^*]+)\*\*/g, "$1"),
                bold: isHeader,
                color: isHeader ? "FFFFFF" : "2d2d2d",
                size: 17,
              })],
            })],
          })),
        });
      }),
    });
  }

  // Build one .docx buffer from a document string
  function buildDocx(text, title, names) {
    const children = [];

    // Title banner
    children.push(new Paragraph({
      children: [
        new TextRun({ text: eventName || "Event Brief", bold: true, size: 24, color: "FFFFFF" }),
        new TextRun({ text: `   •   ${[org, date].filter(Boolean).join("  ·  ")}`, size: 20, color: "D6E4F0" }),
      ],
      shading: { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" },
      spacing: { before: 0, after: 80 },
    }));

    // Document label (e.g. "VOLUNTEER BRIEF")
    children.push(new Paragraph({
      children: [new TextRun({ text: title.toUpperCase(), bold: true, size: 20, color: "FFFFFF" })],
      shading: { fill: NAME, type: ShadingType.CLEAR, color: "auto" },
      spacing: { before: 0, after: 120 },
    }));

    const lines = text.split("\n");
    let i = 0;
    while (i < lines.length) {
      const t = lines[i].trim();

      // Collect consecutive table lines and render as a real Word table
      if (/^\|/.test(t)) {
        const tableLines = [];
        while (i < lines.length && /^\|/.test(lines[i].trim())) { tableLines.push(lines[i]); i++; }
        if (tableLines.filter(l => !/^\|[-| :]+\|$/.test(l.trim())).length > 1) {
          children.push(buildTable(tableLines));
          children.push(new Paragraph({ children: [new TextRun({ text: "" })], spacing: { after: 80 } }));
        }
        continue;
      }

      if (!t) { i++; continue; }

      // ALL-CAPS section headers (EVENT SUMMARY, RUN OF SHOW, ROSTER, etc.)
      if (/^[A-Z][A-Z\s]{3,}$/.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t, bold: true, size: 19, color: ACCENT })],
          spacing: { before: 140, after: 50 },
          border: { bottom: { color: LIGHT, size: 4, space: 1, style: BorderStyle.SINGLE } },
        }));
        i++; continue;
      }

      // Markdown headings
      if (/^## /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^## /, ""), bold: true, size: 19, color: ACCENT })],
          spacing: { before: 120, after: 40 },
          border: { bottom: { color: LIGHT, size: 3, space: 1, style: BorderStyle.SINGLE } },
        }));
        i++; continue;
      }
      if (/^### /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^### /, ""), bold: true, size: 18, color: NAME })],
          spacing: { before: 80, after: 30 },
        }));
        i++; continue;
      }

      // Time-block lines (e.g. "3:00 PM — Welcome")
      if (/^\d{1,2}:\d{2}/.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t, bold: true, size: 18, color: ACCENT })],
          spacing: { before: 100, after: 20 },
        }));
        i++; continue;
      }

      // Checkbox items [ ]
      if (/^\[[ x]\]/.test(t)) {
        children.push(new Paragraph({
          children: parseWithNames(t, names, 18),
          spacing: { before: 14, after: 14 },
          indent: { left: 240 },
        }));
        i++; continue;
      }

      // Bullet lines
      if (/^[•\-*]\s/.test(t)) {
        children.push(new Paragraph({
          children: parseWithNames(t.replace(/^[•\-*]\s+/, ""), names, 18),
          bullet: { level: 0 },
          spacing: { before: 14, after: 14 },
        }));
        i++; continue;
      }

      // Default paragraph
      children.push(new Paragraph({
        children: parseWithNames(t, names, 18),
        spacing: { after: 36 },
      }));
      i++;
    }

    return new Document({
      sections: [{
        properties: { page: { margin: { top: 720, bottom: 720, left: 720, right: 720 } } },
        children,
      }],
    });
  }

  try {
    // Extract names once from the brief so all three docs share the same highlight set
    const names = extractNames(docs[0]);

    // Build all three .docx buffers in parallel
    const buffers = await Promise.all(
      docs.map((text, i) => Packer.toBuffer(buildDocx(text, docTitles[i], names)))
    );

    // Pack all three .docx files into a zip and send it
    const folderName = (eventName || "event").replace(/\s+/g, "_");
    const zipFilename = `${folderName}_docs.zip`;

    const zip = new JSZip();
    const folder = zip.folder(folderName);
    buffers.forEach((buf, i) => folder.file(`${fileNames[i]}.docx`, buf));

    const zipBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${zipFilename}"`);
    res.send(zipBuffer);
  } catch (e) {
    console.error("[/api/export-word]", e);
    res.status(500).json({ error: "Word export failed. Please try again." });
  }
});

app.listen(PORT, () => {
  console.log(`\n✅ Event Planning Agent running at http://localhost:${PORT}\n`);
});
