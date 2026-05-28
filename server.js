require("dotenv").config();
const express = require("express");
const path = require("path");
const Anthropic = require("@anthropic-ai/sdk");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
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

  const titles = ["Volunteer HQ Brief", "Staffing Plan", "Logistics Checklist"];
  const children = [];

  children.push(new Paragraph({
    children: [new TextRun({ text: eventName || "Event Brief", bold: true, size: 40 })],
    alignment: AlignmentType.CENTER,
  }));
  if (org) children.push(new Paragraph({
    children: [new TextRun({ text: org, size: 24, color: "555555" })],
    alignment: AlignmentType.CENTER,
  }));
  if (date) children.push(new Paragraph({
    children: [new TextRun({ text: date, size: 22, color: "888888" })],
    alignment: AlignmentType.CENTER,
  }));
  children.push(new Paragraph({ text: "" }));
  children.push(new Paragraph({ text: "" }));

  // Parse inline bold (**text**) into TextRun array
  function parseInline(text, baseSize = 20, baseColor = "333333") {
    const runs = [];
    const parts = text.split(/(\*\*[^*]+\*\*)/g);
    parts.forEach(part => {
      const boldMatch = part.match(/^\*\*([^*]+)\*\*$/);
      if (boldMatch) {
        runs.push(new TextRun({ text: boldMatch[1], bold: true, size: baseSize, color: baseColor }));
      } else if (part) {
        runs.push(new TextRun({ text: part, size: baseSize, color: baseColor }));
      }
    });
    return runs.length ? runs : [new TextRun({ text: "", size: baseSize })];
  }

  docs.forEach((doc, di) => {
    children.push(new Paragraph({
      children: [new TextRun({ text: titles[di], bold: true, size: 28, color: "111111" })],
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }));

    let inCodeBlock = false;
    doc.split("\n").forEach(line => {
      const t = line.trim();

      // Toggle code block — render contents as monospace indented text
      if (/^```/.test(t)) { inCodeBlock = !inCodeBlock; return; }
      if (inCodeBlock) {
        children.push(new Paragraph({
          children: [new TextRun({ text: line, size: 18, color: "444444", font: "Courier New" })],
          spacing: { before: 20, after: 20 },
          indent: { left: 360 },
        }));
        return;
      }

      // Skip markdown table separator rows
      if (/^\|[-| :]+\|$/.test(t)) return;

      // H1: # heading
      if (/^# /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^# /, ""), bold: true, size: 28, color: "111111" })],
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 320, after: 160 },
        }));
        return;
      }

      // H2: ## heading
      if (/^## /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^## /, ""), bold: true, size: 24, color: "111111" })],
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 240, after: 120 },
        }));
        return;
      }

      // H3: ### heading
      if (/^### /.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t.replace(/^### /, ""), bold: true, size: 22, color: "222222" })],
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 180, after: 80 },
        }));
        return;
      }

      // Horizontal rule ---
      if (/^---+$/.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: "─".repeat(60), size: 16, color: "cccccc" })],
          spacing: { before: 80, after: 80 },
        }));
        return;
      }

      // Table row | col | col |
      if (/^\|/.test(t)) {
        const cols = t.split("|").filter((_, i, a) => i > 0 && i < a.length - 1).map(c => c.trim());
        const rowText = cols.join("   |   ");
        children.push(new Paragraph({
          children: parseInline(rowText, 18, "222222"),
          spacing: { before: 20, after: 20 },
          indent: { left: 120 },
        }));
        return;
      }

      // Blockquote > text
      if (/^> /.test(t)) {
        const inner = t.replace(/^> /, "").replace(/\*\*([^*]+)\*\*/g, "$1");
        children.push(new Paragraph({
          children: [new TextRun({ text: inner, italics: true, size: 18, color: "555555" })],
          spacing: { before: 60, after: 60 },
          indent: { left: 360 },
        }));
        return;
      }

      // Bullet: - item or * item
      if (/^[-*] /.test(t)) {
        children.push(new Paragraph({
          children: parseInline(t.replace(/^[-*] /, ""), 20, "333333"),
          bullet: { level: 0 },
          spacing: { before: 40, after: 40 },
        }));
        return;
      }

      // Emoji section heading (📌 🧑 etc.)
      if (/^[📌🧑📺🗓🏢📸🏷🪧🔧🤝📱📣]/.test(t)) {
        children.push(new Paragraph({
          children: [new TextRun({ text: t, bold: true, size: 22, color: "111111" })],
          spacing: { before: 200, after: 100 },
        }));
        return;
      }

      // Empty line
      if (!t) {
        children.push(new Paragraph({ children: [new TextRun({ text: "" })], spacing: { after: 60 } }));
        return;
      }

      // Default paragraph — strip inline code backticks, then parse bold
      const cleaned = t.replace(/`([^`]+)`/g, "$1").replace(/\*([^*]+)\*/g, "$1");
      children.push(new Paragraph({
        children: parseInline(cleaned, 20, "333333"),
        spacing: { after: 60 },
      }));
    });

    children.push(new Paragraph({ text: "" }));
  });

  const document = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(document);
  const filename = `${(eventName || "event").replace(/\s+/g, "_")}_brief.docx`;

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(buffer);
});


app.listen(PORT, () => {
  console.log(`\n✅ Event Planning Agent running at http://localhost:${PORT}\n`);
});
