exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  const { docs, eventName } = JSON.parse(event.body || "{}");
  if (!docs?.length) {
    return { statusCode: 400, body: JSON.stringify({ error: "No documents" }) };
  }

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
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url: urlMatch?.[0] || null, message: text }),
    };
  } catch (e) {
    console.error(e);
    return { statusCode: 500, body: JSON.stringify({ error: e.message }) };
  }
};
