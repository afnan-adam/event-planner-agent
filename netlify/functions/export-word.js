const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
} = require("docx");
const JSZip = require("jszip");

const MAX_DOC_LENGTH = 30000;
const TITLES = ["Volunteer Brief", "Staffing Plan", "Logistics Checklist"];
const FILES = ["1_Volunteer_Brief.docx", "2_Staffing_Plan.docx", "3_Logistics_Checklist.docx"];

function json(statusCode, payload) {
  return {
    statusCode,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
    body: JSON.stringify(payload),
  };
}

function safeFilename(value) {
  return String(value || "event")
    .normalize("NFKD")
    .replace(/[^a-zA-Z0-9_-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 60) || "event";
}

function makeParagraph(line) {
  const text = line.trim();
  if (!text) return null;

  if (/^#{1,3}\s/.test(text)) {
    const level = (text.match(/^#+/) || ["#"])[0].length;
    const headingText = text.replace(/^#{1,3}\s*/, "");
    return new Paragraph({
      text: headingText,
      heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
      spacing: { before: 160, after: 70 },
      border: level === 1 ? {
        bottom: { color: "D6E4F0", size: 4, space: 1, style: BorderStyle.SINGLE },
      } : undefined,
    });
  }

  if (/^[A-Z][A-Z\s&/()'-]{3,}$/.test(text)) {
    return new Paragraph({
      children: [new TextRun({ text, bold: true, color: "1F4E79", size: 21 })],
      spacing: { before: 160, after: 60 },
      border: { bottom: { color: "D6E4F0", size: 4, space: 1, style: BorderStyle.SINGLE } },
    });
  }

  if (/^\[[ x]\]/i.test(text)) {
    return new Paragraph({
      children: [new TextRun({ text, size: 19, color: "2D2D2D" })],
      indent: { left: 240 },
      spacing: { before: 30, after: 30 },
    });
  }

  if (/^[•*-]\s/.test(text)) {
    return new Paragraph({
      text: text.replace(/^[•*-]\s+/, ""),
      bullet: { level: 0 },
      spacing: { before: 25, after: 25 },
    });
  }

  if (/^\d{1,2}:\d{2}/.test(text)) {
    return new Paragraph({
      children: [new TextRun({ text, bold: true, color: "1F4E79", size: 19 })],
      spacing: { before: 90, after: 30 },
    });
  }

  return new Paragraph({
    children: [new TextRun({ text, size: 19, color: "2D2D2D" })],
    spacing: { after: 55 },
  });
}

function buildDocx(content, title, metadata) {
  const children = [
    new Paragraph({
      children: [new TextRun({
        text: metadata.eventName || "Event Brief",
        bold: true,
        size: 30,
        color: "FFFFFF",
      })],
      alignment: AlignmentType.CENTER,
      shading: { fill: "1F4E79" },
      spacing: { before: 0, after: 90 },
    }),
    new Paragraph({
      children: [new TextRun({
        text: [metadata.org, metadata.date].filter(Boolean).join("  ·  "),
        size: 19,
        color: "555555",
      })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 180 },
    }),
    new Paragraph({
      text: title.toUpperCase(),
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 120, after: 120 },
    }),
  ];

  for (const line of content.split(/\r?\n/)) {
    const paragraph = makeParagraph(line);
    if (paragraph) children.push(paragraph);
  }

  return new Document({
    sections: [{
      properties: { page: { margin: { top: 720, bottom: 720, left: 720, right: 720 } } },
      children,
    }],
  });
}

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers: { Allow: "POST" }, body: "Method Not Allowed" };
  }

  if (!event.body || event.body.length > 150000) {
    return json(413, { error: "Export request is too large." });
  }

  let body;
  try {
    body = JSON.parse(event.body);
  } catch {
    return json(400, { error: "Request body must be valid JSON." });
  }

  const { docs, eventName = "Event", org = "", date = "" } = body || {};
  if (!Array.isArray(docs) || docs.length !== 3 || !docs.every(
    (doc) => typeof doc === "string" && doc.trim() && doc.length <= MAX_DOC_LENGTH
  )) {
    return json(400, { error: "Three valid documents are required for export." });
  }

  for (const [field, value] of [["eventName", eventName], ["org", org], ["date", date]]) {
    if (typeof value !== "string" || value.length > 200) {
      return json(400, { error: `${field} must be a text value of 200 characters or fewer.` });
    }
  }

  try {
    const metadata = { eventName: eventName.trim(), org: org.trim(), date: date.trim() };
    const buffers = await Promise.all(docs.map(async (content, index) => {
      const document = buildDocx(content, TITLES[index], metadata);
      return Packer.toBuffer(document);
    }));

    const folderName = safeFilename(metadata.eventName);
    const zip = new JSZip();
    const folder = zip.folder(folderName);
    buffers.forEach((buffer, index) => folder.file(FILES[index], buffer));

    const zipBuffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    const filename = `${folderName}_docs.zip`;

    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store",
      },
      body: zipBuffer.toString("base64"),
      isBase64Encoded: true,
    };
  } catch (error) {
    console.error("[export-word] Word export failed:", error?.message || "Unknown error");
    return json(500, { error: "Word export failed. Please try again." });
  }
};
