# outlook-gcal-mirror

Mirror Outlook Web (OWA) calendar event **titles + attendee names** into a dedicated Google Calendar so Granola can see the real meeting details.

Constraints:

- No Outlook/Graph API access: we read events from **outlook.office.com in a real logged-in browser**.
- We connect to that browser via **CDP** (Chrome DevTools Protocol) using the built-in keepalive command.
- Google Calendar writes use OAuth and **must not email attendees**.

## Architecture

See `SPEC.md` for the current mini-spec and decisions.

This repo is split conceptually into three layers:

1. **CLI providers** (`src/providers/*`)
   - Outlook reads/writes via `cli-365` subprocesses.
   - Google reads/writes via `gog` subprocesses (for bidirectional sync).

2. **OWA tooling** (`src/owa/*`)
   - Discovery/capture/template-based fetch helpers for debugging and endpoint discovery.
   - Uses CDP against a real logged-in Outlook Web tab.

3. **Google Calendar sync** (`src/google/*`, `src/sync/*`)
   - One-way `sync`: Outlook (`cli-365`) -> Google Calendar mirror.
   - Bidirectional `sync-bidir`: `cli-365` <-> `gog` reconciliation + link state.

## Quickstart (recommended)

### 0) Requirements

- Node **>= 20**
- A logged-in Outlook Web (OWA) account
- A Google OAuth client (type: **Installed app**)
- Chrome/Chromium available for Playwright/Puppeteer

### 1) Install dependencies

```bash
pnpm install
# or: npm install
```

Notes:

- Playwright + Puppeteer are already dependencies. Choose the runtime with `--engine`.
- If Playwright can’t find a browser, run `npx playwright install chromium`.
- If your package manager skips optional deps, install Google APIs:
  `pnpm add googleapis` (or `npm install googleapis`).

### 2) Create a config file (one-time)

```bash
node src/cli.js setup \
  --cdp-port 9222 \
  --engine playwright \
  --target-url https://outlook.office.com/calendar/view/week \
  --google-credentials /path/to/client_secret.json \
  --calendar "Outlook Mirror"
```

Default config path: `~/.config/outlook-gcal-mirror/config.json`

Pass `--config /path/to/config.json` if you want a different location.

### 3) Start Outlook in a CDP-enabled browser

```bash
DISPLAY=:1 XAUTHORITY=$HOME/.Xauthority node src/cli.js keepalive --target-url https://outlook.office.com/calendar/view/week -p 9222 --only-if-idle
```

Log in and make sure the calendar week view is loaded.

Tip: pass `--user-data-dir ~/.config/outlook-gcal-mirror/chrome` to keep a stable Chrome profile.

### 4) Discover OWA internal requests (one-time)

```bash
node src/cli.js discover-owa --cdp-port 9222 --engine playwright --duration-ms 120000 --min-score 1 --no-url-filter
```

Then:

- Click a calendar event to open its details.
- Optionally navigate between weeks.

The discovery command prints candidate request patterns and a `suggestedTemplate`.

If you see “No candidates found”:

- close extra tabs (so CDP attaches to the calendar tab)
- increase `--duration-ms`
- lower `--min-score`
- set `--no-url-filter` (some tenants use `outlook.cloud.microsoft`)

If live discovery misses the initial payload, you can record traffic and scan the log later:

```bash
DISPLAY=:1 XAUTHORITY=$HOME/.Xauthority node src/cli.js keepalive \
  --target-url https://outlook.office.com/calendar/view/week \
  -p 9222 \
  --only-if-idle \
  --record-network ~/.config/outlook-gcal-mirror/logs/owa.ndjson \
  --record-include outlook.office.com

node src/cli.js discover-owa-log --log ~/.config/outlook-gcal-mirror/logs/owa.ndjson --save-templates
```

If you use `--save-templates`, the file defaults to `~/.config/outlook-gcal-mirror/templates.json`. You can point to it via `outlook.owaTemplatesPath` in config.

Note: the log can include sensitive data; delete it when done or use `--no-record-body`.

To use template-based fetch mode, paste a template into your config:

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

### 5) Google OAuth (one-time)

Provide credentials JSON via `--google-credentials /path/to/client_secret.json`.

The first sync will open the OAuth flow and store a token at:
`~/.config/outlook-gcal-mirror/google-token.json` (override with `--google-token`).

### 6) Verify extraction

Passive capture mode:

```bash
node src/cli.js capture-owa --cdp-port 9222 --engine playwright --json
```

Template fetch mode (requires `outlook.owaRequestTemplate` in config):

```bash
node src/cli.js fetch-owa --json
```

### 7) Run a sync

CLI mode (default; uses `cli-365` on PATH):

```bash
node src/cli.js sync \
  --cli365-cdp-port 36429 \
  --cli365-ensure-cdp \
  --google-credentials /path/to/client_secret.json \
  --calendar "Outlook Mirror" \
  --window-days 14
```



Notes:

- `sync` reads Outlook events via `cli-365` only.
- `--calendar` accepts a calendar id or name; id match is attempted first. If the calendar doesn’t exist, it will be created.
- `--lookback-days` (default: 1) includes recently-started events.
- `--cli365-ensure-cdp` asks `cli-365` to start/connect CDP and wait for login.
- For direct OWA extraction/debugging, use `capture-owa` or `fetch-owa`.

## Bi-directional sync (WIP MVP)

This repo now includes `sync-bidir`, which syncs both ways using CLI tools:

- Outlook side: `cli-365` (called as a subprocess from PATH)
- Google side: `gog` (`gog calendar ...`)

### Requirements

- `gog` installed and authenticated (`gog auth ...`)
- Outlook session available to `cli-365` (typically with `--cli365-cdp-port` and optional `--cli365-ensure-cdp`).

### Run

```bash
node src/cli.js sync-bidir \
  --google-calendar primary \
  --cli365-config ~/.config/cli-365/config.yaml \
  --cli365-cdp-port 36429 \
  --cli365-ensure-cdp \
  --window-days 14 \
  --lookback-days 1
```

Useful flags:

- `--state-path <path>`: mapping state file (default: `~/.config/outlook-gcal-mirror/bidir-state.json`)
- `--gog-account <email>`: pass-through to `gog --account`
- `--dry-run`: compute plan without writes/state updates

Current MVP behavior:

- Create + update propagation both directions.
- Identity matching for pre-existing unmatched events (`summary + start + end`).
- Loop prevention via local state links (`outlookId <-> googleId`).
- Conflict rule: if both sides changed since last sync, Outlook wins.
- Deletion propagation is not implemented yet; missing linked events are recreated.
- Legacy one-way mirror events on Google (`Mirrored from Outlook (read-only)`) are skipped.

## Config file

Default path: `~/.config/outlook-gcal-mirror/config.json`

Notes:

- Paths are treated literally (no `~` expansion). Use absolute paths in config.
- Template files default to `~/.config/outlook-gcal-mirror/templates.json`.
- Bidir state defaults to `~/.config/outlook-gcal-mirror/bidir-state.json`.

A minimal config (what `setup` writes):

```json
{
  "outlook": {
    "cdpPort": 9222,
    "engine": "playwright",
    "targetUrl": "https://outlook.office.com/calendar/view/week"
  },
  "google": {
    "credentialsPath": "/path/to/client_secret.json",
    "tokenPath": "/home/you/.config/outlook-gcal-mirror/google-token.json",
    "calendarName": "Outlook Mirror"
  },
  "sync": {
    "windowDays": 14,
    "markCancelled": false
  }
}
```

Optional filters (drop these under `outlook`):

```json
{
  "includeCalendars": ["Team"],
  "skipCalendars": ["Holidays"],
  "includeOwnerEmails": ["me@company.com"],
  "skipOwnerEmails": ["no-reply@company.com"]
}
```

## Clear mirrored events

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

There’s also a helper script that runs keepalive + discovery + verify + sync in one shot:

```bash
pnpm run mirror:all
```

Override behavior via env vars (see `scripts/mirror-all.js` for the full list).

## Release (semantic-release)

- Conventional Commits required (feat/fix/docs/refactor/perf/test/build/ci/chore/style).
- Commit messages linted locally (husky) and in CI.
- Local pre-push runs lint + tests.
- CI on main creates tag + GitHub Release and writes CHANGELOG.md.
- Local dry-run: `pnpm run release -- --dry-run --no-ci`.
- No npm publish (private package).
