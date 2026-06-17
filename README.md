# Event Planning Agent

An AI-powered event operations tool. Fill out a 6-step form with your event details and get three production-ready documents in seconds — a Volunteer HQ Brief, Staffing Plan, and Day-of Logistics Checklist. Export to a formatted Word document.

**Live demo:** https://event-planner-agent-production.up.railway.app

---

## What it does

Event coordinators typically spend 1–2 hours writing briefs before each event. This tool automates that by collecting structured input and using Claude to generate personalized, actionable documents that reference specific volunteers, rooms, agenda timings, and contingency plans.

**Generated documents:**
- **Volunteer HQ Brief** — room assignments, A/V setup, run-of-show timeline, check-in plan, photography notes, special guest handling, post-event actions
- **Staffing Plan** — roster table, room-by-room duties, escalation chain, staggered arrival times
- **Day-of Checklist** — time-blocked checklist from pre-arrival through post-event close-out

---

## Architecture

```
Browser (public/index.html)
  └─ 6-step intake form (vanilla JS, no build step)
  └─ POST /api/generate  ──►  server.js
                                └─ 3× Claude API calls in parallel
                                     (claude-sonnet-4-6, ~2–4s total)
                                └─ JSON { brief, staffing, checklist }
  └─ POST /api/export-word ──►  server.js
                                └─ markdown → docx via `docx` npm package
                                └─ streams .docx file to browser
```

**State management:** Form state is persisted to `localStorage` between steps. The volunteer roster is saved separately so it can be reloaded across events.

---

## Features

- 6-step intake form (basics, rooms & A/V, agenda, volunteers, attendees, logistics)
- Auto-generated run-of-show timeline from agenda segments with live minute totals
- Volunteer roster saved to `localStorage` — reusable across events
- Form state auto-saved — pick up where you left off after a page refresh
- Export to formatted Word (.docx) with styled tables, section banners, and a title page
- Mobile-responsive — single-column layout under 640 px

---

## Tech stack

| Layer | Technology |
|---|---|
| Frontend | Vanilla JS, no build step |
| Backend | Node.js + Express |
| AI | Anthropic Claude (`claude-sonnet-4-6`) |
| Word export | `docx` npm package |
| Security | `helmet`, `express-rate-limit`, input validation |
| Deployment | Railway |

---

## Local setup

```bash
# 1. Install dependencies
npm install

# 2. Add your API key
cp .env.example .env
# Open .env and set ANTHROPIC_API_KEY=sk-ant-...

# 3. Start the server
npm start
# or, for auto-reload during development:
npm run dev
```

Open **http://localhost:3000** in your browser.

The server will exit immediately on startup if `ANTHROPIC_API_KEY` is not set, rather than failing silently on the first request.

---

## Environment variables

| Variable | Required | Default | Description |
|---|---|---|---|
| `ANTHROPIC_API_KEY` | Yes | — | Your Anthropic API key |
| `PORT` | No | `3000` | Port the server listens on |

---

## Project structure

```
event-planner/
├── server.js              # Express server — /api/generate and /api/export-word
├── public/
│   └── index.html         # Full UI (vanilla JS, no build step)
├── netlify/
│   └── functions/         # Serverless equivalents for Netlify deployment
│       ├── generate.js    # Same as /api/generate; uses Haiku to reduce cold-start cost
│       └── export-word.js # Simplified Word export for Netlify
├── package.json
├── .env.example
├── netlify.toml           # Netlify function routing
└── railway.json           # Railway deployment config
```

---

## Security

These weren't added as an afterthought — each one addresses a specific failure mode for a demo app that wraps a paid API:

- **Rate limiting** (20 req / 15 min per IP) — the biggest practical risk for a public demo is runaway API spend if the app is crawled or the key leaks. Rate limiting is the cheapest mitigation and the first thing a reviewer will look for on any app that proxies a paid service.

- **Input validation with hard length caps** — each Claude call costs money and takes time. Without length caps on `context` and `docs`, a single malformed or malicious request could trigger a max-token generation against a 20k-character prompt. Validation also means error messages stay meaningful: a bad request fails at the boundary with a 400, not somewhere inside the Claude SDK with a cryptic exception.

- **Sanitized error responses** — the catch blocks log full detail server-side but return only a generic message to the client. This prevents leaking internal information (model names, stack traces, API error codes) that could help an attacker map the system. It's a small change with no downside.

- **`helmet` security headers** — one line that sets `X-Content-Type-Options`, `X-Frame-Options`, `Referrer-Policy`, and others. No reason not to include it; skipping it is the kind of thing that shows up as a finding in even a light security review.

- **Same-origin CORS** — the frontend is served by the same Express process, so there's no legitimate cross-origin caller. Locking this down explicitly (rather than relying on browser defaults) means the policy is visible in code and doesn't depend on the deployment environment behaving a certain way.

- **Fail-fast API key check** — the server exits immediately on startup if `ANTHROPIC_API_KEY` is missing, with a clear console message. The alternative — failing silently on the first real request — makes the root cause much harder to diagnose in a deployed environment.

---

## Known limitations

- **No authentication** — this is a single-user demo tool. Do not expose it publicly with a shared API key without adding auth.
- **No streaming** — all three Claude calls complete before anything is returned to the browser (typically 2–4 s on Sonnet).
- **Netlify Word export** — the Netlify function version (`netlify/functions/export-word.js`) is an older, simpler implementation without styled tables. The Railway/local `server.js` version is the reference.
- **Context length** — the event context passed to Claude is capped at 20,000 characters. Extremely long agendas or notes fields will be truncated at the client validation step.
