require("dotenv").config();
const express = require("express");
const path = require("path");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const Anthropic = require("@anthropic-ai/sdk");
const { Document, Packer, Paragraph, TextRun, ShadingType, BorderStyle } = require("docx");

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
// "brief" is the volunteer-facing one-pager exported to Word.
// "ops" is a fuller coordinator reference shown in a preview tab only.
app.post("/api/generate", apiLimiter, async (req, res) => {
  const validationError = validateGenerate(req.body);
  if (validationError) return res.status(400).json({ error: validationError });

  const { context } = req.body;
  const leadTime = Number(req.body.leadTime) || 60;

  try {
    const [brief, ops] = await Promise.all([
      // Volunteer-facing brief — short, scannable, one page target.
      anthropic.messages.create({
        model: "claude-sonnet-4-6",
        max_tokens: 900,
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
  • Name: one short action phrase
  • Name: one short action phrase
  Only list volunteers who have something specific to do in that segment. Max 2–3 bullets per segment.]

KEY NOTES
[3 bullets max. Only genuinely important operational facts that don't fit above — e.g. who to call if A/V fails, opt-out names for photography, overflow room trigger. Nothing that repeats the timeline.]

Output ONLY the document. No markdown headers with #. No extra sections.

EVENT DATA:
${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),

      // Coordinator ops reference — fuller detail, preview tab only, not exported.
      anthropic.messages.create({
        model: "claude-sonnet-4-6",
        max_tokens: 1200,
        messages: [{
          role: "user",
          content: `You are an expert event operations coordinator. Using the event details below, produce a concise coordinator ops reference covering: staffing roster, room-by-room duties, escalation contact, arrival timeline, and key contingency notes. Use bullet points. Be specific but brief.\n\nEVENT DATA:\n${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),
    ]);

    res.json({ brief, staffing: ops, checklist: `Pre-arrival (T-${leadTime}min) checklist is included in the ops reference tab.` });
  } catch (e) {
    console.error("[/api/generate]", e);
    res.status(500).json({ error: "Document generation failed. Please try again." });
  }
});

// ── Export to Word ────────────────────────────────────────────────────────────
// Exports only docs[0] (the volunteer brief). Volunteer names are highlighted
// in accent blue so each person can visually scan to their own lines.
app.post("/api/export-word", apiLimiter, async (req, res) => {
  const validationError = validateExportWord(req.body);
  if (validationError) return res.status(400).json({ error: validationError });

  const { docs, eventName, org, date } = req.body;

  const ACCENT = "1F4E79";   // dark navy — headings, section labels
  const NAME  = "2E74B5";   // mid blue — volunteer names
  const LIGHT = "D6E4F0";   // light blue — rule lines

  // Build a set of volunteer first+last names from the eventName field isn't
  // available here, so we extract them from the doc text itself: any token
  // that appears before a colon at the start of a bullet is treated as a name.
  function extractNames(text) {
    const names = new Set();
    for (const line of text.split("\n")) {
      // Match "Name:" at start of a bullet (any bullet char) or at start of a line
      const m = line.match(/^[\s•\-*]*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?):/);
      if (m) names.add(m[1].trim());
    }
    return names;
  }

  // Split a line into TextRun segments, bolding+coloring any known volunteer name.
  function parseWithNames(text, names, baseSize = 18) {
    if (!names.size) return [new TextRun({ text, size: baseSize, color: "2d2d2d" })];
    const escaped = [...names].map(n => n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|");
    const parts = text.split(new RegExp(`(${escaped})`, "g"));
    return parts.filter(p => p).map(p =>
      names.has(p)
        ? new TextRun({ text: p, bold: true, size: baseSize, color: NAME })
        : new TextRun({ text: p, size: baseSize, color: "2d2d2d" })
    );
  }

  const doc = docs[0];
  const names = extractNames(doc);
  const children = [];

  // Single-line title banner
  children.push(new Paragraph({
    children: [
      new TextRun({ text: eventName || "Event Brief", bold: true, size: 24, color: "FFFFFF" }),
      new TextRun({ text: `   •   ${[org, date].filter(Boolean).join("  ·  ")}`, size: 20, color: "D6E4F0" }),
    ],
    shading: { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" },
    spacing: { before: 0, after: 140 },
  }));

  for (const line of doc.split("\n")) {
    const t = line.trim();
    if (!t) continue;

    // ALL-CAPS section headers (e.g. "EVENT SUMMARY", "RUN OF SHOW")
    if (/^[A-Z][A-Z\s]{3,}$/.test(t)) {
      children.push(new Paragraph({
        children: [new TextRun({ text: t, bold: true, size: 20, color: "FFFFFF" })],
        shading: { fill: ACCENT, type: ShadingType.CLEAR, color: "auto" },
        spacing: { before: 120, after: 60 },
      }));
      continue;
    }

    // Markdown headings
    if (/^### /.test(t)) {
      children.push(new Paragraph({
        children: [new TextRun({ text: t.replace(/^### /, ""), bold: true, size: 18, color: ACCENT })],
        spacing: { before: 80, after: 30 },
      }));
      continue;
    }
    if (/^## /.test(t)) {
      children.push(new Paragraph({
        children: [new TextRun({ text: t.replace(/^## /, ""), bold: true, size: 20, color: ACCENT })],
        spacing: { before: 100, after: 40 },
        border: { bottom: { color: LIGHT, size: 3, space: 1, style: BorderStyle.SINGLE } },
      }));
      continue;
    }
    if (/^# /.test(t)) {
      children.push(new Paragraph({
        children: [new TextRun({ text: t.replace(/^# /, ""), bold: true, size: 22, color: ACCENT })],
        spacing: { before: 120, after: 50 },
        border: { bottom: { color: LIGHT, size: 4, space: 1, style: BorderStyle.SINGLE } },
      }));
      continue;
    }

    // Time-block lines (e.g. "3:00 PM — Welcome & Intro (10 min)")
    if (/^\d{1,2}:\d{2}/.test(t)) {
      children.push(new Paragraph({
        children: [new TextRun({ text: t, bold: true, size: 18, color: ACCENT })],
        spacing: { before: 100, after: 20 },
      }));
      continue;
    }

    // Bullet lines — apply name highlighting
    if (/^[•\-*]\s/.test(t)) {
      const content = t.replace(/^[•\-*]\s+/, "");
      children.push(new Paragraph({
        children: parseWithNames(content, names, 18),
        bullet: { level: 0 },
        spacing: { before: 16, after: 16 },
      }));
      continue;
    }

    // Default paragraph — apply name highlighting
    children.push(new Paragraph({
      children: parseWithNames(t, names, 18),
      spacing: { after: 36 },
    }));
  }

  try {
    const document = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 720, bottom: 720, left: 720, right: 720 },
          },
        },
        children,
      }],
    });
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
