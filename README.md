# outlook-gcal-mirror

Mirror Outlook Web (OWA) calendar event **titles + attendee names** into a dedicated Google Calendar so Granola can see the real meeting details.

Constraints:

- No Outlook/Graph API access: we read events from **outlook.office.com in a real logged-in browser**.
- We connect to that browser via **CDP** (Chrome DevTools Protocol) using the built-in keepalive command.
- Google Calendar writes use OAuth and **must not email attendees**.

## Architecture

See `SPEC.md` for the current mini-spec and decisions.

This repo is split conceptually into two layers:

1. **OWA in-browser client** (`src/owa/*`)

   - Extracts events by either:
     - capturing OWA JSON responses (passive; depends on what the UI loads), or
     - running `fetch()` inside the tab using existing cookies/session (deterministic once you have a stable internal endpoint).
   - Supports a **discovery** mode to help identify OWA’s internal JSON endpoints.

2. **Google Calendar sync** (`src/google/*`, `src/sync/*`)
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

Use the built-in keepalive command to start Chromium with CDP enabled:

```bash
node src/cli.js keepalive --target-url https://outlook.office.com/calendar/view/week -p 9222 --only-if-idle
```

Log in and make sure the calendar week view is loaded.

### 2) Discover OWA internal requests (one-time)

In another terminal:

```bash
node src/cli.js discover-owa --cdp-port 9222 --engine playwright -duration-ms 120000 --min-score 1 --no-url-filter

# If you see "No candidates found", try:
# - closing extra tabs (so CDP attaches to the calendar tab)
# - increasing duration: --duration-ms 120000
# - lowering the score threshold: --min-score 1
# - disabling URL filtering (some tenants use different hosts): --no-url-filter
```

Then, in the Outlook tab:

- Click a calendar event to open its details.
- Optionally navigate between weeks.

The discovery command prints candidate request patterns (URL + method) and a `suggestedTemplate`.

If live discovery misses the initial payload, you can record traffic and scan the log later:

```bash
node src/cli.js keepalive \
  --target-url https://outlook.office.com/calendar/view/week \
  -p 9222 \
  --only-if-idle \
  --record-network ~/.config/outlook-gcal-mirror/logs/owa.ndjson \
  --record-include outlook.office.com

node src/cli.js discover-owa-log --log ~/.config/outlook-gcal-mirror/logs/owa.ndjson --save-templates
```

If you use `--save-templates`, the file defaults to `~/.config/outlook-gcal-mirror/templates.json`. You can point to it via `outlook.owaTemplatesPath` in config.

Note: the log can include sensitive data; delete it when done or use `--no-record-body`.

To use template-based fetch mode, paste one of those templates into your config:

```json
{
  "outlook": {
    "owaRequestTemplate": {
      "method": "POST",
      "url": "https://outlook.office.com/...",
      "headers": {
        "accept": "application/json",
        "content-type": "application/json",
        "x-owa-canary": "{{owaCanary}}"
      },
      "body": "{{body}}"
    }
  }
}
```

Notes:

- The template supports placeholders: `{{start}}`, `{{end}}`, `{{owaCanary}}`.
- If your endpoint needs extra constants (folder ids, user ids, etc.), put them under `outlook.owaTemplateVars` and reference them as `{{myVar}}`.

### 3) Google OAuth (one-time)

You’ll need a Google Cloud OAuth client for an “Installed app” (Desktop).

Provide credentials JSON via `--google-credentials /path/to/client_secret.json`.

### 4) Verify extraction

Passive capture mode:

```bash
node src/cli.js capture-owa --cdp-port 9222 --engine playwright --json
```

Template fetch mode (requires `outlook.owaRequestTemplate` in config):

```bash
node src/cli.js fetch-owa --json
```

### 5) Run a sync

Capture mode (default):

```bash
node src/cli.js sync \
  --cdp-port 9222 \
  --engine playwright \
  --source capture \
  --capture-ms 30000 \
  --google-credentials /path/to/client_secret.json \
  --calendar "Outlook Mirror" \
  --window-days 14
```

Template mode (recommended once you have a stable endpoint):

```bash
node src/cli.js sync \
  --cdp-port 9222 \
  --engine playwright \
  --source template \
  --google-credentials /path/to/client_secret.json \
  --calendar "Outlook Mirror" \
  --window-days 14
```

Notes:

- `--calendar` accepts a calendar id or name; id match is attempted first. Default: `Outlook Mirror`.
- Capture mode only sees what OWA loads during the capture window. If you need more coverage, increase `--capture-ms` and navigate weeks while it runs.
- Only use `--mark-cancelled` in capture mode if you’re confident the capture covered the full time window.

### 6) Clear mirrored events

This is **dry-run by default**; add `--yes` to actually delete.

```bash
node src/cli.js clear \
  --google-credentials /path/to/client_secret.json \
  --calendar "Outlook Mirror"

# actually delete
node src/cli.js clear \
  --google-credentials /path/to/client_secret.json \
  --calendar "Outlook Mirror" \
  --yes
```

## Safety: no attendee email

This tool **does not invite Outlook attendees** to Google events. Attendee names are written to the event description only.

It **always adds `owner@example.com` as a single attendee** on mirrored events, but uses Google Calendar API options (`sendUpdates: "none"`) to suppress any notifications.

## Scheduling

Target cadence is every 30 minutes. On macOS, use a LaunchAgent `StartInterval=1800`.

(We’ll add a sample `launchd` plist once the core sync is finalized.)
