# Outlook → Google Calendar Mirror (Mini Spec)

## Problem

User’s Outlook calendar syncs to Google Calendar as **free/busy only**, so Granola (transcription app) cannot see correct meeting titles and attendee names. Granola cannot connect to Office 365 directly.

## Goal

Monitor Outlook calendar events and mirror **real event details** into a dedicated Google Calendar so Granola can read:
- Event title
- Attendee display names
- Start/end times

## Key Constraints

- **No Outlook API / Graph access.**
- Outlook must be accessed via **Outlook Web (OWA)** in a real logged-in browser session.
- The tool must use [`browser-keepalive`](https://github.com/lc0rp/browser-keepalive) with **CDP** enabled to keep OWA open and controllable.
- Google Calendar writes use OAuth.
- **Do not email any attendee** (no invitations, no updates).
- Must support both **Playwright** and **Puppeteer** for CDP connection.
- Runs on User’s **macOS desktop**.

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

Two major components, kept logically separate:

### 1) `owa-client` (in-browser OWA API)

Purpose: read Outlook event details without relying on a first-party API.

- Connect to a running Chromium instance via CDP (port provided by `browser-keepalive`).
- Prefer **calling the same internal JSON endpoints OWA uses**:
  - Execute `fetch()` inside the Outlook tab context so OWA cookies/session tokens apply.
  - Parse JSON into a normalized event model.
- Fallback path if endpoints are too brittle:
  - Observe network responses and/or scrape DOM/flyouts as needed.

### 2) `gcal-sync` (Google Calendar writer)

Purpose: upsert the mirror calendar without sending any notifications.

- Authenticate via Google OAuth (Desktop/Installed app).
- Ensure the “Outlook Mirror” calendar exists (or select an existing one).
- Upsert events into that calendar.

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
- Connect to the CDP browser session and confirm OWA week view is open.
- Let the user select/include/skip Outlook calendars (by name) and optionally owner emails.
- Perform Google OAuth and select/create the destination calendar.
- Save config under `~/.config/outlook-gcal-mirror/config.json` (paths overridable).

## Execution

A `sync` run should:
- Connect to CDP (Playwright or Puppeteer).
- Capture/read Outlook events for a configurable time window (e.g. today-1…today+N days).
- Filter out Outlook items that appear to originate from Google (based on configured rules).
- Upsert active events into the mirror calendar.
- Mark missing prior mirrored events as cancelled.

## Open Questions / TBD

- Exact OWA internal endpoint(s) and required headers/tokens (to be determined via discovery on User’s Mac).
- Best strategy to enumerate events over a time range (OWA endpoint vs week-view paging).
- How to represent recurring events (instance ids vs series ids) for stable mapping.
