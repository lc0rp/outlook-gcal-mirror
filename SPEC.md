# Outlook → Google Calendar Mirror (Mini Spec)

## Problem

Some Outlook calendar integrations sync to Google Calendar as **free/busy only**, so downstream tools cannot see correct meeting titles and attendee names.

## Goal

Monitor Outlook calendar events and mirror **real event details** into a dedicated Google Calendar so Granola can read:
- Event title
- Attendee display names
- Start/end times

## Key Constraints

- Outlook IO in this repo is via `cli-365` subprocess calls (no direct Outlook SDK/API integration here).
- Google Calendar writes use OAuth.
- Google CLI reads/writes for bidirectional sync go through `gog`.
- **Do not email any attendee** (no invitations, no updates).
- Runs on a user-managed desktop or server.

## Decisions (Confirmed)

- Mirrored events live in a dedicated Google calendar (default name: **“Outlook Mirror”**).
- Existing Outlook→Google free/busy sync will be disabled (free/busy blocks will disappear).
- If an Outlook event is deleted/cancelled, the mirrored Google event is **not deleted**; it is **marked cancelled**.
- Calendar selection is **configurable** and/or selectable during an initial `setup`.
- There may be Google→Outlook syncing in the other direction; we must detect those Outlook items and **skip mirroring them**.
  - Detection signals can include Outlook calendar/folder name and owner email; these should be configurable.
- Granola only needs title + attendee names + time fields; we do not need to preserve every Outlook field.
- Initial cadence target: **every 30 minutes**.

## Architecture

Three major components, kept logically separate:

### 1) `cli-providers` (subprocess adapters)

Purpose: isolate external calendar CLIs from sync logic.

- Outlook IO via `cli-365`.
- Google CLI IO via `gog`.
- Normalize subprocess responses into stable sync inputs.

### 2) `gcal-sync` (one-way mirror writer)

Purpose: upsert the mirror calendar without sending any notifications.

- Authenticate via Google OAuth (Desktop/Installed app).
- Ensure the “Outlook Mirror” calendar exists (or select an existing one).
- Read Outlook events from `cli-365`, filter them, and upsert them into that calendar.

### 3) `bidir-sync` (CLI-to-CLI reconciler, MVP)

Purpose: sync Outlook and Google in both directions without direct SDK/API integration in this repo.

- Outlook IO via `cli-365` CLI subprocess calls.
- Google IO via `gog` CLI subprocess calls.
- Local link state (`bidir-state.json`) tracks `outlookId <-> googleId` mappings + last fingerprint.
- Existing unmatched events are linked by identity (`summary + start + end`) before creating duplicates.

## Data Model

Normalized Outlook event fields used for sync:
- `sourceKey` (stable idempotency key; ideally Outlook event id, else a hash)
- `subject`
- `start`, `end`
- `attendeeNames[]`
- Optional: `organizerEmail`, `sourceCalendarName`, `sourceOwnerEmail`

## Idempotency & Mapping

- Each mirrored Google event stores the Outlook identity in `extendedProperties.private`:
  - `ogm.sourceKey = <sourceKey>`
  - `ogm.status = active|cancelled`
- Sync should be repeatable without creating duplicates.

For bidirectional mode:

- Local state file stores link pairs.
- Google-created/updated events include private props:
  - `ogm.link.outlookId`
  - `ogm.link.version=1`

## Cancellation Behavior

When an Outlook event is no longer present in the scan window:
- Find the corresponding mirrored Google event by `ogm.sourceKey`.
- **Do not delete**.
- Mark as cancelled by:
  - Prefixing summary with `CANCELLED:` (idempotent)
  - Setting `ogm.status=cancelled`
  - Optionally appending a note to description.

## “No attendee email” guarantee

- Never add guests to Google events.
- Attendee names are written to the description only.
- All writes must suppress notifications (e.g. Google Calendar API `sendUpdates=none`).

## Setup & Configuration

Initial setup should:
- Configure Google OAuth and mirror calendar defaults.
- Let the user select/include/skip Outlook calendars (by name) and optionally owner emails.
- Save config under `~/.config/outlook-gcal-mirror/config.json` (paths overridable).

Per-run flags may still pass `cli-365` runtime inputs such as config path, optional CDP bootstrap params, and optional folder filter.

## Execution

A `sync` run should:
- Read Outlook events from `cli-365 calendar list` for a configurable time window (e.g. today-1…today+N days).
- Filter out Outlook items that appear to originate from Google (based on configured rules).
- Upsert active events into the mirror calendar.
- Mark missing prior mirrored events as cancelled.

For `sync-bidir` (current MVP):

- Read Outlook events from `cli-365 calendar list`.
- Read Google events from `gog calendar events`.
- Reconcile links:
  - existing links first
  - then identity-based linking
  - then create missing counterparts
- Propagate updates:
  - one-sided change: push to other side
  - two-sided change: Outlook wins (current policy)
- Skip legacy one-way mirrored Google events (`Mirrored from Outlook (read-only)`).

## Open Questions / TBD

- How to represent recurring events (instance ids vs series ids) for stable mapping.
- Deletion semantics for bidirectional sync (currently recreated, not propagated).
