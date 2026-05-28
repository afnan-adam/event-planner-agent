# Event Planning Agent

An AI-powered event operations tool. Fill out a 6-step form with your event details and get three production-ready documents in seconds — a Volunteer HQ Brief, Staffing Plan, and Day-of Logistics Checklist. Export to a formatted Word document.

**Live demo:** https://event-planner-agent-production.up.railway.app

---

## What it does

Event coordinators typically spend 1-2 hours writing briefs before each event. This tool automates that by collecting structured input and using Claude to generate personalized, actionable documents that reference specific volunteers, rooms, agenda timings, and contingency plans.

**Generated documents:**
- **Volunteer HQ Brief** — room assignments, A/V setup, run-of-show timeline, check-in plan, photography notes, special guest handling, post-event actions
- **Staffing Plan** — roster table, room-by-room duties, escalation chain, staggered arrival times
- **Day-of Checklist** — time-blocked checklist from pre-arrival through post-event close-out

---

## Features

- 6-step intake form (basics, rooms & A/V, agenda, volunteers, attendees, logistics)
- Auto-generated run-of-show timeline from your agenda segments
- Volunteer roster saved to browser localStorage — reusable across events
- Form state auto-saved — pick up where you left off
- Export to formatted Word (.docx) with styled tables, headings, and title page

---

## Tech stack

- **Frontend** — Vanilla JS, no build step
- **Backend** — Node.js + Express
- **AI** — Anthropic Claude (claude-sonnet-4-6)
- **Deployment** — Railway

---

## Local setup

```bash
# 1. Install dependencies
npm install

# 2. Add your API key
cp .env.example .env
# Edit .env and add your ANTHROPIC_API_KEY

# 3. Start the server
npm start
```

Open **http://localhost:3000** in your browser.

---

## Project structure

```
event-planner/
├── server.js        # Express server — generate, export-word routes
├── public/
│   └── index.html   # Full UI (vanilla JS, no build step)
├── package.json
├── .env.example
└── railway.json     # Railway deployment config
```
