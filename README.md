# Event Planning Agent

An AI-assisted event operations tool. Complete a six-step event intake form and generate three editable planning documents: a Volunteer HQ Brief, a Staffing Plan, and a Day-of Logistics Checklist. Export all three as separate Word documents in one ZIP file.

**Railway demo (existing):** https://event-planner-agent-production.up.railway.app

## What it does

- Collects event basics, rooms and A/V details, agenda segments, volunteer roles, attendee/check-in information, and logistics.
- Builds an operational brief with room assignments, volunteer roles, run of show, and key notes.
- Produces a staffing plan with room duties, escalation guidance, and contingencies.
- Creates a time-blocked day-of checklist.
- Lets coordinators preview each generated document and export three Word files in one ZIP.
- Saves form progress and a reusable volunteer roster in browser local storage.

## Architecture

```
Browser (public/index.html)
  ├─ Six-step event intake form (vanilla JavaScript)
  ├─ POST /api/generate ──► Netlify Function: generate
  │                           └─ Three parallel Claude API calls
  │                               └─ { brief, staffing, checklist }
  └─ POST /api/export-word ► Netlify Function: export-word
                              └─ docx + JSZip
                                  └─ Three .docx files in one .zip

Railway deployment continues to use server.js (Express).
```

## Tech stack

| Layer | Technology |
|---|---|
| Frontend | HTML, CSS, vanilla JavaScript |
| Netlify backend | Netlify Functions (Node.js) |
| Railway backend | Node.js + Express |
| AI | Anthropic Claude API (Claude Haiku 4.5 in Netlify functions) |
| Word export | `docx` |
| ZIP packaging | `jszip` |
| Railway security middleware | `helmet`, `express-rate-limit`, input validation |

## Deploy to Netlify

1. Import this GitHub repository into Netlify.
2. Set the base directory to the repository root.
3. Set the publish directory to `public`.
4. Set the functions directory to `netlify/functions`.
5. Add `ANTHROPIC_API_KEY` in Netlify environment variables and grant it Functions scope.
6. Deploy the `feature/netlify-deployment` branch as a branch deploy or merge it into `master` after reviewing and testing.

The repository's `netlify.toml` already configures the publish directory, functions directory, esbuild bundling, and the `/api/generate` and `/api/export-word` routes. No frontend build command is required.

**Important:** Keep `ANTHROPIC_API_KEY` in Netlify's server-side environment settings. Never place it in `public/index.html`, expose it to browser code, or commit a real key to Git.

## Test locally

Install dependencies and run the Netlify development server:

```bash
npm install
npx netlify dev
```

Add `ANTHROPIC_API_KEY` to a local `.env` file for local testing. Netlify Dev serves the static site and emulates the configured function routes.

Test both workflows:
- Submit a representative event and confirm all three document previews render.
- Export the documents and confirm the ZIP contains `1_Volunteer_Brief.docx`, `2_Staffing_Plan.docx`, and `3_Logistics_Checklist.docx`.

## Environment variables

| Variable | Required | Used by |
|---|---|---|
| `ANTHROPIC_API_KEY` | Yes | Netlify Functions and Railway server |
| `PORT` | No | Railway/local Express server; defaults to 3000 |

## Project structure

```
event-planner-agent/
├── server.js
├── public/
│   └── index.html
├── netlify/
│   └── functions/
│       ├── generate.js
│       ├── export-word.js
│       └── export-drive.js
├── netlify.toml
├── railway.json
└── package.json
```

## Security and limitations

- The Netlify generation function validates request size, event-context length, and arrival lead time; it returns generic errors rather than exposing provider error details.
- It includes a **best-effort, in-memory per-instance throttle**. Serverless instances are ephemeral, so this is not a durable/global rate limiter and should not be treated as sufficient abuse protection for unrestricted public use.
- The app has no user authentication. Before promoting it as a public production service, add authentication or a shared, durable rate-limiting mechanism and monitor API usage.
- Generated plans are drafts. Event coordinators should verify names, timings, assignments, access instructions, and contingency details before distributing them.
- The Netlify Word export is aligned with the frontend's expected ZIP response and produces three separate Word documents. Formatting is intentionally simpler than the full Railway export implementation.
- The Google Drive export route is present in the existing configuration; it is not part of the validated Netlify deployment path in this update.

## Local Railway setup

```bash
npm install
cp .env.example .env
# Set ANTHROPIC_API_KEY in .env
npm start
``

Open http://localhost:3000.
