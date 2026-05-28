const Anthropic = require("@anthropic-ai/sdk");

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  const { context, leadTime } = JSON.parse(event.body || "{}");
  if (!context) {
    return { statusCode: 400, body: JSON.stringify({ error: "Missing event context" }) };
  }

  try {
    const [brief, staffing, checklist] = await Promise.all([
      anthropic.messages.create({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 1024,
        messages: [{
          role: "user",
          content: `You are an expert event operations coordinator. Using the event details below, produce a comprehensive Volunteer HQ Brief. Include these emoji-headed sections:\n\n📌 Room Assignments\n🧑‍💻 Technical Support Assignments (by room, using volunteer names and roles provided)\n📺 A/V & Tech Setup (setup owner, equipment, hybrid link if applicable)\n🗓 Run-of-Show Timeline (build a timed schedule using the agenda segments and start time)\n🏢 Overflow Room Plan (include contingency if overflow is not yet activated)\n📸 Recording & Photography (opt-out seating note, photo team duties, recommended shot list)\n🏷️ Check-In & Attendance (method, name tag details, tracker name and snapshot time)\n🪧 Signage & Materials (who places what, handouts)\n🔧 Room Setup & Arrival Plan (staggered arrival times per role, upon-arrival checklist)\n🤝 Special Guests (specific handling notes per guest)\n📱 On-Site Contact\n📣 Post-Event Actions (follow-up email, sharing, debrief)\n\nUse all names provided. Be specific and actionable. Output ONLY the document.\n\nEVENT DATA:\n${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),

      anthropic.messages.create({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 1024,
        messages: [{
          role: "user",
          content: `You are an expert event operations coordinator. Using the event details below, produce a clear Staffing Plan:\n\n1. Roster table — Name | Role | Room | Key Responsibilities\n2. Room-by-room breakdown with specific duties\n3. Escalation chain\n4. Staggered arrival timeline per person\n5. Contingency notes (overflow, A/V issues, late arrivals)\n\nOutput ONLY the document.\n\nEVENT DATA:\n${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),

      anthropic.messages.create({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 1024,
        messages: [{
          role: "user",
          content: `You are an expert event operations coordinator. Using the event details below, produce a Day-of Logistics Checklist organized by time block:\n\n- Pre-arrival (T-${leadTime || 60}min)\n- Setup (T-30min)\n- Final sweep (T-10min)\n- Doors open\n- One section per agenda segment (by name)\n- Wrap-up\n- Post-event close-out\n\nEach item as [ ]. Include owner names where known. Output ONLY the checklist.\n\nEVENT DATA:\n${context}`,
        }],
      }).then(r => r.content.map(b => b.text || "").join("")),
    ]);

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ brief, staffing, checklist }),
    };
  } catch (e) {
    console.error(e);
    return { statusCode: 500, body: JSON.stringify({ error: e.message }) };
  }
};
