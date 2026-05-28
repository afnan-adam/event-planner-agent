# Event Planning Agent

AI-powered event planning tool — fill out a 6-step form, get a volunteer brief, staffing plan, and logistics checklist. Exports to Word or Google Drive.

## Setup

```bash
# 1. Install dependencies
npm install

# 2. Add your API key
cp .env.example .env
# Edit .env and paste your ANTHROPIC_API_KEY

# 3. Start the server
npm start
# or for auto-restart on changes:
npm run dev
```

Then open **http://localhost:3000** in your browser.

---

## Features

| Feature | Notes |
|---|---|
| 6-step intake form | Basics, Rooms & A/V, Agenda, Volunteers, Attendees, Logistics |
| Run-of-show timeline | Auto-generated from your agenda segments |
| Volunteer memory | Saved to `roster.json` — persists between sessions |
| Export to Word (.docx) | Downloads a formatted Word document |
| Save to Google Drive | Requires Google Drive connected in your Anthropic account |

---

## Google Drive export

The Drive export calls the Anthropic API with Google Drive MCP (`drivemcp.googleapis.com`).
This works when you have Google Drive connected via [claude.ai](https://claude.ai) (Settings → Integrations).

The API call is made server-side using your `ANTHROPIC_API_KEY`.

---

## Volunteer roster memory

Volunteers are saved to `roster.json` in the project root whenever you click **Save current** on step 4. On your next event, click **Load saved** to repopulate the roster — great for recurring events with the same team.

---

## Project structure

```
event-planner/
├── server.js          # Express server + all API routes
├── public/
│   └── index.html     # Full UI (vanilla JS, no build step)
├── package.json
├── .env               # Your API key (not committed)
├── roster.json        # Auto-created when you save a roster
└── README.md
```
