import { describe, expect, it, vi } from "vitest";

import { runBidirectionalSync, _internal } from "./bidir.js";
import { BIDIR_STATE_VERSION } from "./bidir-state.js";

function dt(value) {
	return { dateTime: value };
}

function buildOutlook(id, summary, start, end, extra = {}) {
	return {
		id,
		summary,
		start: dt(start),
		end: dt(end),
		allDay: false,
		location: "",
		description: "",
		attendeeNames: [],
		isCancelled: false,
		...extra,
	};
}

function buildGoogle(id, summary, start, end, extra = {}) {
	return {
		id,
		summary,
		start: dt(start),
		end: dt(end),
		allDay: false,
		location: "",
		description: "",
		attendeeNames: [],
		status: "confirmed",
		...extra,
	};
}

function makeStateStore(state) {
	return {
		load: vi.fn().mockResolvedValue(state),
		save: vi.fn().mockResolvedValue(undefined),
	};
}

describe("sync/bidir", () => {
	it("creates google event for unmatched outlook event", async () => {
		const outlookClient = {
			listEvents: vi.fn().mockResolvedValue([buildOutlook("o1", "Meet", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z")]),
			createEvent: vi.fn(),
			updateEvent: vi.fn(),
		};
		const googleClient = {
			listEvents: vi.fn().mockResolvedValue([]),
			createEvent: vi.fn().mockResolvedValue(buildGoogle("g1", "Meet", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z")),
			updateEvent: vi.fn(),
		};
		const stateStore = makeStateStore({ version: BIDIR_STATE_VERSION, links: [] });

		const result = await runBidirectionalSync({
			outlookClient,
			googleClient,
			calendarId: "primary",
			range: { start: new Date("2026-01-01T00:00:00Z"), end: new Date("2026-01-07T00:00:00Z") },
			stateStore,
		});

		expect(result.summary.createdOnGoogle).toBe(1);
		expect(googleClient.createEvent).toHaveBeenCalledTimes(1);
		expect(stateStore.save).toHaveBeenCalledTimes(1);
	});

	it("creates outlook event for unmatched google event", async () => {
		const outlookClient = {
			listEvents: vi.fn().mockResolvedValue([]),
			createEvent: vi.fn().mockResolvedValue(buildOutlook("o1", "Meet", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z")),
			updateEvent: vi.fn(),
		};
		const googleClient = {
			listEvents: vi.fn().mockResolvedValue([buildGoogle("g1", "Meet", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z")]),
			createEvent: vi.fn(),
			updateEvent: vi.fn(),
		};
		const stateStore = makeStateStore({ version: BIDIR_STATE_VERSION, links: [] });

		const result = await runBidirectionalSync({
			outlookClient,
			googleClient,
			calendarId: "primary",
			range: { start: new Date("2026-01-01T00:00:00Z"), end: new Date("2026-01-07T00:00:00Z") },
			stateStore,
		});

		expect(result.summary.createdOnOutlook).toBe(1);
		expect(outlookClient.createEvent).toHaveBeenCalledTimes(1);
	});

	it("links by identity and updates google to outlook version", async () => {
		const outlookClient = {
			listEvents: vi
				.fn()
				.mockResolvedValue([buildOutlook("o1", "Meet", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z", { location: "A" })]),
			createEvent: vi.fn(),
			updateEvent: vi.fn(),
		};
		const googleClient = {
			listEvents: vi
				.fn()
				.mockResolvedValue([buildGoogle("g1", "Meet", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z", { location: "B" })]),
			createEvent: vi.fn(),
			updateEvent: vi.fn().mockResolvedValue(buildGoogle("g1", "Meet", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z", { location: "A" })),
		};
		const stateStore = makeStateStore({ version: BIDIR_STATE_VERSION, links: [] });

		const result = await runBidirectionalSync({
			outlookClient,
			googleClient,
			calendarId: "primary",
			range: { start: new Date("2026-01-01T00:00:00Z"), end: new Date("2026-01-07T00:00:00Z") },
			stateStore,
		});

		expect(result.summary.linkedByIdentity).toBe(1);
		expect(result.summary.updatedOnGoogle).toBe(1);
	});

	it("updates outlook when google changed relative to last fingerprint", async () => {
		const outlook = buildOutlook("o1", "Meet", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z");
		const google = buildGoogle("g1", "Meet v2", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z");
		const fp = _internal.eventFingerprint(outlook);
		const stateStore = makeStateStore({
			version: BIDIR_STATE_VERSION,
			links: [{ outlookId: "o1", googleId: "g1", lastFingerprint: fp }],
		});
		const outlookClient = {
			listEvents: vi.fn().mockResolvedValue([outlook]),
			createEvent: vi.fn(),
			updateEvent: vi.fn().mockResolvedValue({ action: "updated", event: { ...outlook, summary: "Meet v2" } }),
		};
		const googleClient = {
			listEvents: vi.fn().mockResolvedValue([google]),
			createEvent: vi.fn(),
			updateEvent: vi.fn(),
		};

		const result = await runBidirectionalSync({
			outlookClient,
			googleClient,
			calendarId: "primary",
			range: { start: new Date("2026-01-01T00:00:00Z"), end: new Date("2026-01-07T00:00:00Z") },
			stateStore,
		});

		expect(result.summary.updatedOnOutlook).toBe(1);
		expect(outlookClient.updateEvent).toHaveBeenCalledTimes(1);
	});

	it("skips legacy mirrored google events from one-way sync", async () => {
		const outlookClient = {
			listEvents: vi.fn().mockResolvedValue([]),
			createEvent: vi.fn(),
			updateEvent: vi.fn(),
		};
		const googleClient = {
			listEvents: vi.fn().mockResolvedValue([
				buildGoogle("g1", "Legacy", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z", {
					description: "Mirrored from Outlook (read-only)",
				}),
			]),
			createEvent: vi.fn(),
			updateEvent: vi.fn(),
		};
		const stateStore = makeStateStore({ version: BIDIR_STATE_VERSION, links: [] });

		const result = await runBidirectionalSync({
			outlookClient,
			googleClient,
			calendarId: "primary",
			range: { start: new Date("2026-01-01T00:00:00Z"), end: new Date("2026-01-07T00:00:00Z") },
			stateStore,
		});

		expect(result.summary.skippedLegacyGoogleEvents).toBe(1);
		expect(outlookClient.createEvent).not.toHaveBeenCalled();
	});

	it("dry-run does not call create/update or save", async () => {
		const outlookClient = {
			listEvents: vi.fn().mockResolvedValue([buildOutlook("o1", "Meet", "2026-01-01T10:00:00Z", "2026-01-01T11:00:00Z")]),
			createEvent: vi.fn(),
			updateEvent: vi.fn(),
		};
		const googleClient = {
			listEvents: vi.fn().mockResolvedValue([]),
			createEvent: vi.fn(),
			updateEvent: vi.fn(),
		};
		const stateStore = makeStateStore({ version: BIDIR_STATE_VERSION, links: [] });

		const result = await runBidirectionalSync({
			outlookClient,
			googleClient,
			calendarId: "primary",
			range: { start: new Date("2026-01-01T00:00:00Z"), end: new Date("2026-01-07T00:00:00Z") },
			stateStore,
			dryRun: true,
		});

		expect(result.summary.createdOnGoogle).toBe(1);
		expect(googleClient.createEvent).not.toHaveBeenCalled();
		expect(stateStore.save).not.toHaveBeenCalled();
	});
});
