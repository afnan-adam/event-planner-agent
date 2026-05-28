const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = require("docx");

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  const { docs, eventName, org, date } = JSON.parse(event.body || "{}");
  if (!docs?.length) {
    return { statusCode: 400, body: JSON.stringify({ error: "No documents to export" }) };
  }

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

  return {
    statusCode: 200,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename="${filename}"`,
    },
    body: buffer.toString("base64"),
    isBase64Encoded: true,
  };
};
