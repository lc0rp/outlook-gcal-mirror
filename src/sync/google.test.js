import { describe, expect, it, vi, beforeEach } from "vitest";

vi.mock("../google/client.js", () => ({
	getGoogleCalendarClient: vi.fn().mockResolvedValue({
		calendar: { calendarList: { list: vi.fn() }, calendars: { insert: vi.fn() }, events: { list: vi.fn(), insert: vi.fn(), patch: vi.fn() } },
	}),
}));

import {
	PRIVATE_SOURCE_KEY,
	PRIVATE_STATUS,
	findCalendarId,
	ensureMirrorCalendar,
	listMirrorEvents,
	listMirrorEventsAll,
	getGoogleSyncContext,
	upsertMirroredEvent,
	markMirroredCancelled,
} from "./google.js";

describe("sync/google", () => {
	const mockCalendar = {
		calendarList: { list: vi.fn() },
		calendars: { insert: vi.fn() },
		events: { list: vi.fn(), insert: vi.fn(), patch: vi.fn() },
	};

	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe("findCalendarId", () => {
		it("finds by id", async () => {
			mockCalendar.calendarList.list.mockResolvedValue({
				data: { items: [{ id: "cal-1", summary: "Work" }] },
			});
			const id = await findCalendarId({ calendar: mockCalendar, calendarRef: "cal-1" });
			expect(id).toBe("cal-1");
		});

		it("finds by name", async () => {
			mockCalendar.calendarList.list.mockResolvedValue({
				data: { items: [{ id: "cal-2", summary: "Outlook Mirror" }] },
			});
			const id = await findCalendarId({ calendar: mockCalendar, calendarRef: "Outlook Mirror" });
			expect(id).toBe("cal-2");
		});

		it("returns null when not found", async () => {
			mockCalendar.calendarList.list.mockResolvedValue({ data: { items: [] } });
			const id = await findCalendarId({ calendar: mockCalendar, calendarRef: "Missing" });
			expect(id).toBeNull();
		});
	});

	describe("ensureMirrorCalendar", () => {
		it("returns existing calendar id", async () => {
			mockCalendar.calendarList.list.mockResolvedValue({
				data: { items: [{ id: "existing", summary: "Outlook Mirror" }] },
			});
			const id = await ensureMirrorCalendar({ calendar: mockCalendar, calendarRef: "Outlook Mirror" });
			expect(id).toBe("existing");
		});

		it("creates calendar when missing", async () => {
			mockCalendar.calendarList.list.mockResolvedValue({ data: { items: [] } });
			mockCalendar.calendars.insert.mockResolvedValue({ data: { id: "new-cal" } });
			const id = await ensureMirrorCalendar({ calendar: mockCalendar, calendarRef: "My Mirror" });
			expect(id).toBe("new-cal");
			expect(mockCalendar.calendars.insert).toHaveBeenCalled();
		});
	});

	describe("listMirrorEvents", () => {
		it("paginates through results", async () => {
			mockCalendar.events.list
				.mockResolvedValueOnce({ data: { items: [{ id: "e1" }], nextPageToken: "tok" } })
				.mockResolvedValueOnce({ data: { items: [{ id: "e2" }] } });

			const events = await listMirrorEvents({
				calendar: mockCalendar,
				calendarId: "cal",
				timeMin: "2026-01-01T00:00:00Z",
				timeMax: "2026-01-31T23:59:59Z",
			});
			expect(events).toHaveLength(2);
			expect(mockCalendar.events.list).toHaveBeenCalledTimes(2);
		});
	});

	describe("listMirrorEventsAll", () => {
		it("fetches all events without time bounds", async () => {
			mockCalendar.events.list.mockResolvedValue({ data: { items: [{ id: "e1" }] } });
			const events = await listMirrorEventsAll({ calendar: mockCalendar, calendarId: "cal" });
			expect(events).toHaveLength(1);
		});
	});

	describe("upsertMirroredEvent", () => {
		const ev = {
			subject: "Test Meeting",
			sourceKey: "key-1",
			sourceId: "id-1",
			attendeeNames: ["Alice"],
			organizerEmail: "org@example.com",
			start: { dateTime: "2026-01-15T10:00:00Z" },
			end: { dateTime: "2026-01-15T11:00:00Z" },
		};

		it("creates new event when not found", async () => {
			mockCalendar.events.list.mockResolvedValue({ data: { items: [] } });
			mockCalendar.events.insert.mockResolvedValue({});

			const result = await upsertMirroredEvent({ calendar: mockCalendar, calendarId: "cal", ev });
			expect(result.action).toBe("created");
			expect(mockCalendar.events.insert).toHaveBeenCalled();
		});

		it("updates existing event", async () => {
			mockCalendar.events.list.mockResolvedValue({
				data: { items: [{ id: "existing-event" }] },
			});
			mockCalendar.events.patch.mockResolvedValue({});

			const result = await upsertMirroredEvent({ calendar: mockCalendar, calendarId: "cal", ev });
			expect(result.action).toBe("updated");
			expect(mockCalendar.events.patch).toHaveBeenCalled();
		});

		it("finds event by identity fallback", async () => {
			// First call returns no match by sourceKey, second returns matching event by identity
			mockCalendar.events.list
				.mockResolvedValueOnce({ data: { items: [] } })
				.mockResolvedValueOnce({
					data: {
						items: [
							{
								id: "identity-match",
								summary: "Test Meeting",
								start: { dateTime: "2026-01-15T10:00:00Z" },
								end: { dateTime: "2026-01-15T11:00:00Z" },
							},
						],
					},
				});
			mockCalendar.events.patch.mockResolvedValue({});

			const result = await upsertMirroredEvent({ calendar: mockCalendar, calendarId: "cal", ev });
			expect(result.action).toBe("updated");
		});

		it("handles event with date-only start/end", async () => {
			const allDayEv = {
				...ev,
				start: { date: "2026-01-15" },
				end: { date: "2026-01-16" },
			};
			mockCalendar.events.list.mockResolvedValue({ data: { items: [] } });
			mockCalendar.events.insert.mockResolvedValue({});

			const result = await upsertMirroredEvent({ calendar: mockCalendar, calendarId: "cal", ev: allDayEv });
			expect(result.action).toBe("created");
		});
	});

	describe("markMirroredCancelled", () => {
		it("marks event as cancelled", async () => {
			mockCalendar.events.patch.mockResolvedValue({});
			const gcalEvent = { id: "ev-1", summary: "Meeting" };

			const result = await markMirroredCancelled({
				calendar: mockCalendar,
				calendarId: "cal",
				gcalEvent,
			});
			expect(result.action).toBe("cancelled");
			expect(mockCalendar.events.patch).toHaveBeenCalledWith(
				expect.objectContaining({
					requestBody: expect.objectContaining({
						summary: "CANCELLED: Meeting",
					}),
				})
			);
		});

		it("skips if already cancelled", async () => {
			const gcalEvent = {
				id: "ev-1",
				summary: "CANCELLED: Meeting",
				extendedProperties: { private: { [PRIVATE_STATUS]: "cancelled" } },
			};

			const result = await markMirroredCancelled({
				calendar: mockCalendar,
				calendarId: "cal",
				gcalEvent,
			});
			expect(result.action).toBe("skipped");
		});

		it("skips if no id", async () => {
			const result = await markMirroredCancelled({
				calendar: mockCalendar,
				calendarId: "cal",
				gcalEvent: {},
			});
			expect(result.action).toBe("skipped");
		});

		it("marks cancelled: prefixed event as cancelled", async () => {
			mockCalendar.events.patch.mockResolvedValue({});
			const gcalEvent = { id: "ev-1", summary: "cancelled: Old Meeting" };

			const result = await markMirroredCancelled({
				calendar: mockCalendar,
				calendarId: "cal",
				gcalEvent,
			});
			expect(result.action).toBe("cancelled");
		});

		it("preserves description when marking cancelled", async () => {
			mockCalendar.events.patch.mockResolvedValue({});
			const gcalEvent = { id: "ev-1", summary: "Meeting", description: "Original desc" };

			await markMirroredCancelled({
				calendar: mockCalendar,
				calendarId: "cal",
				gcalEvent,
			});

			const patchCall = mockCalendar.events.patch.mock.calls[0][0];
			expect(patchCall.requestBody.description).toContain("Original desc");
			expect(patchCall.requestBody.description).toContain("Marked CANCELLED");
		});
	});

	describe("getGoogleSyncContext", () => {
		it("returns calendar and calendarId", async () => {
			// This uses the mocked getGoogleCalendarClient
			const { getGoogleCalendarClient } = await import("../google/client.js");
			vi.mocked(getGoogleCalendarClient).mockResolvedValue({
				calendar: mockCalendar,
			});
			mockCalendar.calendarList.list.mockResolvedValue({
				data: { items: [{ id: "cal-id", summary: "Mirror" }] },
			});

			const ctx = await getGoogleSyncContext({
				credentialsPath: "creds.json",
				tokenPath: "token.json",
				calendarRef: "Mirror",
			});
			expect(ctx.calendarId).toBe("cal-id");
		});
	});
});
