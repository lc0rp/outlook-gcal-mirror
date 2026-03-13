# outlook-gcal-mirror

Mirror Outlook calendar event **titles + attendee names** into a dedicated Google Calendar so Granola can see the real meeting details.

Constraints:

- Sync runtime reads/writes Outlook via `cli-365` subprocess calls.
- Google Calendar writes use OAuth and **must not email attendees**.

## Architecture

See `SPEC.md` for the current mini-spec and decisions.

This repo is split conceptually into two layers:

1. **CLI providers** (`src/providers/*`)
   - Outlook reads/writes via `cli-365` subprocesses.
   - Google reads/writes via `gog` subprocesses (for bidirectional sync).

2. **Google Calendar sync** (`src/google/*`, `src/sync/*`)
   - One-way `sync`: Outlook (`cli-365`) -> Google Calendar mirror.
   - Bidirectional `sync-bidir`: `cli-365` <-> `gog` reconciliation + link state.

## Quickstart (recommended)

### 0) Requirements

- Node **>= 20**
- `cli-365` installed and authenticated for Outlook calendar access
- A Google OAuth client (type: **Installed app**)
- `gog` installed/authenticated if you use `sync-bidir`

### 1) Install dependencies

```bash
pnpm install
# or: npm install
```

Notes:

- If your package manager skips optional deps, install Google APIs:
  `pnpm add googleapis` (or `npm install googleapis`).

### 2) Create a config file (one-time)

```bash
node src/cli.js setup \
  --google-credentials /path/to/client_secret.json \
  --calendar "Outlook Mirror"
```

Default config path: `~/.config/outlook-gcal-mirror/config.json`

Pass `--config /path/to/config.json` if you want a different location.

### 3) Google OAuth (one-time)

Provide credentials JSON via `--google-credentials /path/to/client_secret.json`.

The first sync will open the OAuth flow and store a token at:
`~/.config/outlook-gcal-mirror/google-token.json` (override with `--google-token`).

### 4) Run a sync

CLI mode (default; uses `cli-365` on PATH):

```bash
scripts/with-gog-keyring.sh node src/cli.js sync \
  --cli365-cdp-port 36429 \
  --cli365-ensure-cdp \
  --google-credentials /path/to/client_secret.json \
  --calendar "Outlook Mirror" \
  --window-days 14
```

Notes:

- `scripts/with-gog-keyring.sh` auto-exports `GOG_KEYRING_PASSWORD` from `/home/user/.config/.env` (fallback: `pass show openclaw/gog_keyring_password`) if not already set; mainly needed for `sync-bidir`/`gog`.
- `sync` reads Outlook events via `cli-365` only.
- `--calendar` accepts a calendar id or name; id match is attempted first. If the calendar doesn’t exist, it will be created.
- `--lookback-days` (default: 1) includes recently-started events.
- Optional `cli-365` pass-through flags: `--cli365-config`, `--cli365-cdp-port`, `--cli365-ensure-cdp`, `--cli365-ensure-cdp-timeout`, `--cli365-folder`.
- `--cli365-ensure-cdp` asks `cli-365` to start/connect CDP and wait for login.

## Bi-directional sync (WIP MVP)

This repo now includes `sync-bidir`, which syncs both ways using CLI tools:

- Outlook side: `cli-365` (called as a subprocess from PATH)
- Google side: `gog` (`gog calendar ...`)

### Requirements

- `gog` installed and authenticated (`gog auth ...`)
- Outlook session available to `cli-365` (typically with `--cli365-cdp-port` and optional `--cli365-ensure-cdp`).

### Run

```bash
scripts/with-gog-keyring.sh node src/cli.js sync-bidir \
  --google-calendar primary \
  --cli365-config ~/.config/cli-365/config.yaml \
  --cli365-cdp-port 36429 \
  --cli365-ensure-cdp \
  --window-days 14 \
  --lookback-days 1
```

Useful flags:

- `--google-calendar <idOrName>`: target calendar (default: `primary`; does not inherit one-way `google.calendarName`)
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
- Bidir state defaults to `~/.config/outlook-gcal-mirror/bidir-state.json`.

A minimal config (what `setup` writes):

```json
{
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

There’s also a helper script that runs a cli-365-first `sync` in one shot:

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

## TODO: Use red/green TDD with complete implementaton (no surface stubs). Track progress, hand-off when necessary, and mark tasks as done below as you complete them. Conventional commits + push after each task
1. - [x] Scrub any local code/path that doesn't rely on `cli-365`. cli-365 is now ready to replace all direct Outlook interactions.
2. - [x] Update README to remove local paths that do not rely on `cli-365`, e.g. --cdp-port, --engine, --targetUrl
3. - [x] Review code for dead paths that are no longer relevant
4. - [ ] Use repo-doc and user-doc skills to update dev and user documentation
