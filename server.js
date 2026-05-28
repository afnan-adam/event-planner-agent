require("dotenv").config();
const express = require("express");
const path = require("path");
const Anthropic = require("@anthropic-ai/sdk");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  PageBreak, UnderlineType,
} = require("docx");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: "5mb" }));
app.use(express.static(path.join(__dirname, "public")));

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── Generate documents ───────────────────────────────────────────────────────
app.post("/api/generate", async (req, res) => {
  const { context, leadTime } = req.body;
  if (!context) return res.status(400).json({ error: "Missing event context" });

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
          content: `You are an expert event operations coordinator. Using the event details below, produce a Day-of Logistics Checklist organized by time block:\n\n- Pre-arrival (T-${leadTime || 60}min)\n- Setup (T-30min)\n- Final sweep (T-10min)\n- Doors open\n- One section per agenda segment (by name)\n- Wrap-up\n- Post-event close-out\n\nEach item as [ ]. Include owner names where known. Output ONLY the checklist.\n\nEVENT DATA:\n${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),
    ]);

    res.json({ brief, staffing, checklist });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// ── Export to Word ───────────────────────────────────────────────────────────
app.post("/api/export-word", async (req, res) => {
  const { docs, eventName, org, date } = req.body;
  if (!docs?.length) return res.status(400).json({ error: "No documents to export" });

  const ACCENT = "1F4E79";   // dark navy
  const ACCENT2 = "2E74B5";  // mid blue
  const LIGHT = "D6E4F0";    // light blue tint

  const titles = ["Volunteer HQ Brief", "Staffing Plan", "Logistics Checklist"];

  // Strip markdown symbols from plain text
  function stripInline(text) {
    return text.replace(/\*\*([^*]+)\*\*/g, "$1").replace(/\*([^*]+)\*/g, "$1").replace(/`([^`]+)`/g, "$1");
  }

  // Parse inline **bold** into TextRun array
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

  // Build a styled Word table from markdown table lines
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

  // Title page section
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

  // One section per document
  docs.forEach((doc, di) => {
    const children = [];

    // Section title banner
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

      // Code block toggle
      if (/^```/.test(t)) { inCodeBlock = !inCodeBlock; i++; continue; }
      if (inCodeBlock) {
        children.push(new Paragraph({
          children: [new TextRun({ text: line, size: 18, color: "444444", font: "Courier New" })],
          spacing: { before: 20, after: 20 },
          indent: { left: 360 },
        }));
        i++; continue;
      }

      // Collect consecutive table lines and render as a real Word table
      if (/^\|/.test(t)) {
        const tableLines = [];
        while (i < lines.length && /^\|/.test(lines[i].trim())) {
          tableLines.push(lines[i]);
          i++;
        }
        if (tableLines.filter(l => !/^\|[-| :]+\|$/.test(l.trim())).length > 1) {
          children.push(buildTable(tableLines));
          children.push(new Paragraph({ children: [new TextRun({ text: "" })], spacing: { after: 120 } }));
        }
        continue;
      }

      // Skip separator lines
      if (/^\|[-| :]+\|$/.test(t) || /^---+$/.test(t)) { i++; continue; }

      // H1
      if (/^# /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^# /, ""), bold: true, size: 28, color: ACCENT })],
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 320, after: 120 },
          border: { bottom: { color: LIGHT, size: 6, space: 1, style: BorderStyle.SINGLE } },
        }));
        i++; continue;
      }

      // H2
      if (/^## /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^## /, ""), bold: true, size: 24, color: ACCENT })],
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 240, after: 100 },
        }));
        i++; continue;
      }

      // H3
      if (/^### /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^### /, ""), bold: true, size: 22, color: ACCENT2 })],
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 180, after: 80 },
        }));
        i++; continue;
      }

      // Blockquote
      if (/^> /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: stripInline(t.replace(/^> /, "")), italics: true, size: 18, color: "555555" })],
          spacing: { before: 60, after: 60 },
          indent: { left: 360 },
          shading: { fill: "F5F9FC", type: ShadingType.CLEAR, color: "auto" },
        }));
        i++; continue;
      }

      // Bullet
      if (/^[-*] /.test(t)) {
        children.push(new Paragraph({
          children: parseInline(t.replace(/^[-*] /, ""), 20, "2d2d2d"),
          bullet: { level: 0 },
          spacing: { before: 40, after: 40 },
        }));
        i++; continue;
      }

      // Emoji section heading
      if (/^[📌🧑📺🗓🏢📸🏷🪧🔧🤝📱📣]/.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t, bold: true, size: 24, color: ACCENT })],
          spacing: { before: 240, after: 120 },
          border: { bottom: { color: LIGHT, size: 4, space: 1, style: BorderStyle.SINGLE } },
        }));
        i++; continue;
      }

      // Empty line
      if (!t) {
        children.push(new Paragraph({ children: [new TextRun({ text: "" })], spacing: { after: 80 } }));
        i++; continue;
      }

      // Default paragraph
      children.push(new Paragraph({
        children: parseInline(t, 20, "2d2d2d"),
        spacing: { after: 80 },
      }));
      i++;
    }

    sections.push({ children });
  });

  const document = new Document({ sections });
  const buffer = await Packer.toBuffer(document);
  const filename = `${(eventName || "event").replace(/\s+/g, "_")}_brief.docx`;

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(buffer);
});


app.listen(PORT, () => {
  console.log(`\n✅ Event Planning Agent running at http://localhost:${PORT}\n`);
});
