---
type: Reference
primary_audience: Users
owner: outlook-gcal-mirror maintainers
last_verified: 2026-03-13
next_review_by: 2026-04-13
source_of_truth: ../../README.md
read_when: >-
  You are setting up the Outlook mirror, running syncs, or debugging cli-365,
  gog, or Google OAuth issues.
---

# Users Start Here

## Read when

- You want first success fast.
- You need the current commands for one-way sync, `sync-bidir`, or `clear`.

## First success

1. Install Node 20+, `cli-365`, and optionally `gog` for `sync-bidir`.
2. Run `pnpm install`.
3. If the CLI reports `Missing dependency 'googleapis'`, run `pnpm add googleapis`.
4. Write config:

   ```bash
   node src/cli.js setup \
     --google-credentials /path/to/client_secret.json \
     --calendar "Outlook Mirror"
   ```

5. Run one-way mirror:

   ```bash
   scripts/with-gog-keyring.sh node src/cli.js sync \
     --cli365-cdp-port 36429 \
     --cli365-ensure-cdp \
     --google-credentials /path/to/client_secret.json \
     --calendar "Outlook Mirror" \
     --window-days 14
   ```

## Daily commands

- `sync`: Outlook (`cli-365`) -> Google mirror. Common pass-through flags:
  `--cli365-config`, `--cli365-cdp-port`, `--cli365-ensure-cdp`,
  `--cli365-ensure-cdp-timeout`, `--cli365-folder`.
- `sync-bidir`: two-way sync with `gog`. Use
  `scripts/with-gog-keyring.sh node src/cli.js sync-bidir --google-calendar primary ...`.
- `clear`: dry-run by default; add `--yes` to delete mirrored Google events.

## Config + state

- Config: `~/.config/outlook-gcal-mirror/config.json`
- Google token: `~/.config/outlook-gcal-mirror/google-token.json`
- Bidir link state: `~/.config/outlook-gcal-mirror/bidir-state.json`
- Paths stay literal. Use absolute paths; no `~` expansion inside config JSON.

## Troubleshooting

- `Missing dependency 'googleapis'`: install it in this repo via `pnpm add googleapis`.
- `cli-365` cannot see Outlook session: rerun with `--cli365-ensure-cdp`, set
  `--cli365-cdp-port`, and verify `cli-365` auth on the machine.
- `gog` prompts for keyring password in automation: use
  `scripts/with-gog-keyring.sh` or export `GOG_KEYRING_PASSWORD` first.
- Need bulk scheduling: `pnpm run mirror:all` wraps `sync` with env-driven flags.
