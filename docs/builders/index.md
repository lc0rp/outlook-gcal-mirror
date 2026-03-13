---
type: Reference
primary_audience: Builders
owner: outlook-gcal-mirror maintainers
last_verified: 2026-03-13
next_review_by: 2026-04-13
source_of_truth: ../../SPEC.md
read_when: >-
  You are changing CLI flags, provider behavior, sync policy, validation gates,
  or repository docs.
---

# Builders Start Here

## Read when

- You are editing sync behavior or external CLI adapters.
- You need the current module map before changing code.

## Module map

- `src/cli.js`: command surface for `setup`, `sync`, `sync-bidir`, `clear`.
- `src/config.js`: config path defaults and JSON load/save.
- `src/providers/cli365.js`: Outlook adapter. Only Outlook IO surface in this repo.
- `src/providers/gog.js`: Google CLI adapter for bidirectional sync.
- `src/google/client.js`: Google OAuth + Calendar API client for one-way sync and `clear`.
- `src/sync/normalized-event.js`: shared event contract between providers/sync logic.
- `src/sync/outlook.js`: maps `cli-365` events into normalized sync inputs.
- `src/sync/google.js`: one-way mirror upsert/cancel behavior.
- `src/sync/bidir.js`: link reconciliation, conflict policy, and create/update propagation.
- `scripts/mirror-all.js`: env-driven wrapper around `sync`.

## Guardrails

- No direct Outlook SDK, REST, browser, or OWA code here. Outlook access stays behind `cli-365`.
- Keep `README.md` and `docs/users/index.md` aligned with the live flag surface.
- Update `SPEC.md` when behavior, conflict rules, or persistence contracts change.
- Add regression coverage in the closest Vitest file; `src/cli.test.js` is the doc/CLI guardrail.

## Validation loop

- `pnpm lint`
- `pnpm test`
- `pnpm test -- src/cli.test.js` for focused CLI/doc regressions

## Release guardrails

- Semantic-release owns tags + GitHub Releases.
- Conventional Commits enforced locally and in CI.
- Pre-push runs lint + tests; keep branches green before handoff.
