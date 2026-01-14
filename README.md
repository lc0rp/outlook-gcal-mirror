# outlook-gcal-mirror

Mirror Outlook Web (OWA) calendar event **titles + attendee names** into a dedicated Google Calendar so Granola can see the real meeting details.

Constraints:
- No Outlook/Graph API access: we read events from **outlook.office.com in a real logged-in browser**.
- We connect to that browser via **CDP** (Chrome DevTools Protocol) using [`browser-keepalive`](https://github.com/lc0rp/browser-keepalive).
- Google Calendar writes use OAuth and **must not email attendees**.

## Architecture

See `SPEC.md` for the current mini-spec and decisions.

This repo is split conceptually into two layers:

1) **OWA in-browser client** (`src/owa/*`)
   - Extracts events by capturing OWA JSON responses (and can optionally run `fetch()` inside the tab using existing cookies/session).
   - Supports a **discovery** mode to help identify OWA’s internal JSON endpoints.

2) **Google Calendar sync** (`src/google/*`, `src/sync/*`)
   - Upserts events into a dedicated Google calendar (default name: `Outlook Mirror`).
   - Stores a stable source id in `extendedProperties.private` for idempotency.

## Setup

### Prereqs

Install deps in this repo:

```bash
npm install

# Choose ONE automation engine:
# - lighter install (uses your existing Chrome):
npm install playwright-core
# or:
npm install puppeteer-core

# Needed for Google Calendar API:
npm install googleapis
```

### 1) Start Outlook in a CDP-enabled browser

Use `browser-keepalive` to start Chromium with CDP enabled:

```bash
browser-keepalive https://outlook.office.com/calendar/view/week -p 9222 --only-if-idle
```

Log in and make sure the calendar week view is loaded.

### 2) Discover OWA internal requests (one-time)

In another terminal:

```bash
node src/cli.js discover-owa --cdp-port 9222 --engine playwright

# If you see "No candidates found", try:
# - closing extra tabs (so CDP attaches to the calendar tab)
# - increasing duration: --duration-ms 120000
# - lowering the score threshold: --min-score 1
# - disabling URL filtering (some tenants use different hosts): --no-url-filter
```

Then, in the Outlook tab:
- Click a calendar event to open its details.
- Optionally navigate between weeks.

The discovery command prints candidate request patterns (URL + method) and a suggested template config.

### 3) Google OAuth (one-time)

You’ll need a Google Cloud OAuth client for an “Installed app” (Desktop).

Provide credentials JSON via `--google-credentials /path/to/client_secret.json`.

### 4) Run a sync

```bash
node src/cli.js sync \
  --cdp-port 9222 \
  --engine playwright \
  --google-credentials /path/to/client_secret.json \
  --calendar-name "Outlook Mirror" \
  --window-days 14
```

Notes:
- The current `sync` implementation reads events from whatever JSON OWA loads during the capture window.
- If you need more coverage, increase `--capture-ms` and navigate weeks while it runs.
- Only use `--mark-cancelled` if you’re confident the capture covered the full time window.

## Safety: no attendee email

This tool **never adds guests** to Google events. Attendee names are written to the event description only.

When updating/creating events, it uses Google Calendar API options that suppress notifications.

## Scheduling

Target cadence is every 30 minutes. On macOS, use a LaunchAgent `StartInterval=1800`.

(We’ll add a sample `launchd` plist once the core sync is finalized.)
