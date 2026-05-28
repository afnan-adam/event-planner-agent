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

  docs.forEach((doc, di) => {
    children.push(new Paragraph({
      children: [new TextRun({ text: titles[di], bold: true, size: 28 })],
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 200 },
    }));

    doc.split("\n").forEach(line => {
      const isHeading = /^[📌🧑📺🗓🏢📸🏷🪧🔧🤝📱📣]/.test(line.trim());
      children.push(new Paragraph({
        children: [new TextRun({
          text: line,
          bold: isHeading,
          size: isHeading ? 24 : 20,
          color: isHeading ? "111111" : "333333",
        })],
        spacing: { after: isHeading ? 140 : 40 },
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

// ── Google Drive export ──────────────────────────────────────────────────────
app.post("/api/export-drive", async (req, res) => {
  const { docs, eventName } = req.body;
  if (!docs?.length) return res.status(400).json({ error: "No documents" });

  const allContent = docs.map((d, i) =>
    `${"─".repeat(38)}\n${["VOLUNTEER HQ BRIEF", "STAFFING PLAN", "LOGISTICS CHECKLIST"][i]}\n${"─".repeat(38)}\n\n${d}`
  ).join("\n\n\n");

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": process.env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "anthropic-beta": "mcp-client-1.0",
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 1024,
        messages: [{
          role: "user",
          content: `Create a new Google Doc titled "${eventName || "Event"} – Volunteer Planning Brief" with the following content. After creating it, share the document URL in your response.\n\n${allContent}`,
        }],
        mcp_servers: [{
          type: "url",
          url: "https://drivemcp.googleapis.com/mcp/v1",
          name: "google-drive-mcp",
        }],
      }),
    });

    const data = await response.json();
    const text = data.content?.filter(b => b.type === "text").map(b => b.text).join(" ") || "";
    const toolResults = data.content?.filter(b => b.type === "mcp_tool_result")
      .map(b => Array.isArray(b.content) ? b.content.map(c => c.text || "").join(" ") : b.content || "")
      .join(" ") || "";

    const urlMatch = (text + " " + toolResults).match(/https:\/\/docs\.google\.com\/[^\s)"'<>]+/);
    res.json({ url: urlMatch?.[0] || null, message: text });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.listen(PORT, () => {
  console.log(`\n✅ Event Planning Agent running at http://localhost:${PORT}\n`);
});
