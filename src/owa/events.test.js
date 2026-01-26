import { describe, it, expect, vi } from "vitest";

import {
	applyRangeToRequestBody,
	applyRangeToRequestHeaders,
	applyRangeToRequestUrl,
	fetchOwaEventsByTemplate,
	fetchOwaEventsByTemplates,
	getOwaEvents,
} from "./events.js";

describe("applyRangeToRequestBody", () => {
	it("overrides common start/end keys in object bodies", () => {
		const body = {
			Body: {
				StartDate: "2025-11-10T00:00:00.000Z",
				EndDate: "2026-01-16T00:00:00.000Z",
			},
			ViewStart: { DateTime: "2025-11-10T00:00:00.000Z" },
			ViewEnd: { DateTime: "2026-01-16T00:00:00.000Z" },
		};

		const range = {
			start: new Date("2026-01-21T00:00:00.000Z"),
			end: new Date("2026-02-04T00:00:00.000Z"),
		};

		const res = applyRangeToRequestBody(body, range);
		expect(res.matched).toBeGreaterThan(0);
		expect(res.body.Body.StartDate).toBe(range.start.toISOString());
		expect(res.body.Body.EndDate).toBe(range.end.toISOString());
		expect(res.body.ViewStart.DateTime).toBe(range.start.toISOString());
		expect(res.body.ViewEnd.DateTime).toBe(range.end.toISOString());
	});

	it("updates JSON-string bodies with DateTime fields", () => {
		const body = JSON.stringify({
			StartTime: { DateTime: "2025-11-10T00:00:00.000Z", TimeZone: "UTC" },
			EndTime: { DateTime: "2026-01-16T00:00:00.000Z", TimeZone: "UTC" },
		});

		const range = {
			start: new Date("2026-01-21T00:00:00.000Z"),
			end: new Date("2026-02-04T00:00:00.000Z"),
		};

		const res = applyRangeToRequestBody(body, range);
		expect(res.parsed).toBe(true);
		expect(res.body.StartTime.DateTime).toBe(range.start.toISOString());
		expect(res.body.EndTime.DateTime).toBe(range.end.toISOString());
	});

	it("updates URL query params for start/end ranges", () => {
		const range = {
			start: new Date("2026-01-21T00:00:00.000Z"),
			end: new Date("2026-02-04T00:00:00.000Z"),
		};

		const url =
			"https://example.test/api?StartDate=2025-11-10&EndDate=2026-01-16&foo=bar&StartDateTime=2025-11-10T00:00:00.000Z";
		const res = applyRangeToRequestUrl(url, range);
		expect(res.matched).toBe(3);
		expect(res.url).toContain("StartDate=2026-01-21");
		expect(res.url).toContain("EndDate=2026-02-04");
		expect(res.url).toContain(`StartDateTime=${encodeURIComponent(range.start.toISOString())}`);
	});

	it("updates URL-encoded JSON headers", () => {
		const range = {
			start: new Date("2026-01-21T00:00:00.000Z"),
			end: new Date("2026-02-04T00:00:00.000Z"),
		};

		const headerJson = {
			Body: {
				RangeStart: "2025-11-10T00:00:00.000Z",
				RangeEnd: "2026-01-16T00:00:00.000Z",
			},
		};
		const headers = {
			"x-owa-urlpostdata": encodeURIComponent(JSON.stringify(headerJson)),
		};

		const res = applyRangeToRequestHeaders(headers, range);
		expect(res.matched).toBeGreaterThan(0);
		const decoded = JSON.parse(decodeURIComponent(res.headers["x-owa-urlpostdata"]));
		expect(decoded.Body.RangeStart).toBe(range.start.toISOString());
		expect(decoded.Body.RangeEnd).toBe(range.end.toISOString());
	});

	it("handles null/undefined body", () => {
		const range = { start: new Date(), end: new Date() };
		expect(applyRangeToRequestBody(null, range)).toEqual({ body: null, matched: 0, parsed: false });
		expect(applyRangeToRequestBody(undefined, range)).toEqual({ body: undefined, matched: 0, parsed: false });
	});

	it("handles null/undefined url", () => {
		const range = { start: new Date(), end: new Date() };
		expect(applyRangeToRequestUrl(null, range)).toEqual({ url: null, matched: 0 });
	});

	it("handles null/undefined headers", () => {
		const range = { start: new Date(), end: new Date() };
		expect(applyRangeToRequestHeaders(null, range)).toEqual({ headers: null, matched: 0 });
	});
});

describe("fetchOwaEventsByTemplate", () => {
	it("throws on missing template", async () => {
		await expect(fetchOwaEventsByTemplate({ page: {}, range: { start: new Date(), end: new Date() } }))
			.rejects.toThrow(/Missing owaRequestTemplate/);
	});

	it("throws on template without url", async () => {
		await expect(fetchOwaEventsByTemplate({
			page: {},
			template: { method: "GET" },
			range: { start: new Date(), end: new Date() },
		})).rejects.toThrow(/must include url/);
	});

	it("fetches events using template", async () => {
		const page = {
			context: () => ({ cookies: () => Promise.resolve([{ name: "X-OWA-CANARY", value: "canary" }]) }),
			evaluate: vi.fn().mockResolvedValue({
				Subject: "Test Event",
				Start: "2026-01-15T10:00:00Z",
				End: "2026-01-15T11:00:00Z",
				Id: "event-1",
			}),
		};

		const events = await fetchOwaEventsByTemplate({
			page,
			template: { url: "https://outlook.office.com/api", method: "POST", body: {} },
			range: { start: new Date("2026-01-01"), end: new Date("2026-01-31") },
		});

		expect(events).toHaveLength(1);
		expect(events[0].subject).toBe("Test Event");
	});
});

describe("fetchOwaEventsByTemplates", () => {
	it("throws on missing viewTemplate", async () => {
		await expect(fetchOwaEventsByTemplates({ page: {}, range: { start: new Date(), end: new Date() } }))
			.rejects.toThrow(/Missing owaRequestTemplate/);
	});

	it("throws on viewTemplate without url", async () => {
		await expect(fetchOwaEventsByTemplates({
			page: {},
			viewTemplate: { method: "GET" },
			range: { start: new Date(), end: new Date() },
		})).rejects.toThrow(/must include url/);
	});

	it("fetches view events without detail template", async () => {
		const page = {
			context: () => ({ cookies: () => Promise.resolve([]) }),
			evaluate: vi.fn().mockResolvedValue({
				Subject: "View Event",
				Start: "2026-01-15T10:00:00Z",
				End: "2026-01-15T11:00:00Z",
				Id: "view-1",
			}),
		};

		const events = await fetchOwaEventsByTemplates({
			page,
			viewTemplate: { url: "https://outlook.office.com/api/view", method: "GET" },
			range: { start: new Date("2026-01-01"), end: new Date("2026-01-31") },
		});

		expect(events).toHaveLength(1);
	});

	it("fetches with event detail template and merges results", async () => {
		const page = {
			context: () => ({ cookies: () => Promise.resolve([{ name: "X-OWA-CANARY", value: "canary" }]) }),
			evaluate: vi.fn()
				// First call: view events
				.mockResolvedValueOnce({
					value: [{ Subject: "Event", Start: "2026-01-15T10:00:00Z", End: "2026-01-15T11:00:00Z", ItemId: { Id: "id-1" } }],
				})
				// Second call: event details
				.mockResolvedValueOnce({
					value: [{ Subject: "Event", Id: "id-1", Attendees: [{ EmailAddress: { Name: "Bob" } }] }],
				}),
		};

		const events = await fetchOwaEventsByTemplates({
			page,
			viewTemplate: { url: "https://outlook.office.com/api/view", method: "POST", body: {} },
			eventTemplate: {
				url: "https://outlook.office.com/api/events",
				method: "POST",
				body: { Body: { EventIds: [] } },
			},
			range: { start: new Date("2026-01-01"), end: new Date("2026-01-31") },
		});

		expect(events.length).toBeGreaterThanOrEqual(1);
	});
});

describe("getOwaEvents", () => {
	it("throws on missing capture options in capture mode", async () => {
		await expect(getOwaEvents({ mode: "capture", page: {} }))
			.rejects.toThrow(/Missing capture options/);
	});

	it("throws on missing template options in template mode", async () => {
		await expect(getOwaEvents({ mode: "template", page: {} }))
			.rejects.toThrow(/Missing template options/);
	});

	it("uses capture mode", async () => {
		const listeners = {};
		const page = {
			on: (event, handler) => { listeners[event] = handler; },
			off: vi.fn(),
		};

		const promise = getOwaEvents({
			mode: "capture",
			page,
			capture: { durationMs: 10 },
		});

		const events = await promise;
		expect(Array.isArray(events)).toBe(true);
	});

	it("uses template mode", async () => {
		const page = {
			context: () => ({ cookies: () => Promise.resolve([]) }),
			evaluate: vi.fn().mockResolvedValue({ Subject: "Test", Start: "2026-01-01", End: "2026-01-02", Id: "id" }),
		};

		const events = await getOwaEvents({
			mode: "template",
			page,
			template: {
				template: { url: "https://outlook.office.com/api", method: "GET" },
				range: { start: new Date(), end: new Date() },
			},
		});

		expect(events).toHaveLength(1);
	});

	it("uses template mode with detailTemplate", async () => {
		const page = {
			context: () => ({ cookies: () => Promise.resolve([]) }),
			evaluate: vi.fn().mockResolvedValue({
				Subject: "Test",
				Start: "2026-01-01",
				End: "2026-01-02",
				ItemId: { Id: "event-id" },
			}),
		};

		const events = await getOwaEvents({
			mode: "template",
			page,
			template: {
				template: { url: "https://outlook.office.com/api/view", method: "GET" },
				detailTemplate: {
					url: "https://outlook.office.com/api/events",
					method: "POST",
					body: { Body: { EventIds: [] } },
				},
				range: { start: new Date(), end: new Date() },
			},
		});

		expect(events.length).toBeGreaterThanOrEqual(1);
	});
});

describe("applyRangeToRequestBody edge cases", () => {
	it("handles non-JSON string body", () => {
		const range = { start: new Date(), end: new Date() };
		const result = applyRangeToRequestBody("not json", range);
		expect(result.matched).toBe(0);
	});

	it("handles primitive non-object parsed body", () => {
		const range = { start: new Date(), end: new Date() };
		const result = applyRangeToRequestBody("123", range);
		expect(result.matched).toBe(0);
	});

	it("handles arrays in body", () => {
		const range = {
			start: new Date("2026-01-01"),
			end: new Date("2026-01-31"),
		};
		const body = [{ StartDate: "2025-01-01", EndDate: "2025-12-31" }];
		const result = applyRangeToRequestBody(body, range);
		expect(result.matched).toBeGreaterThan(0);
	});
});

describe("applyRangeToRequestUrl edge cases", () => {
	it("handles invalid URL", () => {
		const range = { start: new Date(), end: new Date() };
		const result = applyRangeToRequestUrl("not a url", range);
		expect(result.matched).toBe(0);
		expect(result.url).toBe("not a url");
	});
});

describe("applyRangeToRequestHeaders edge cases", () => {
	it("handles non-string header values", () => {
		const range = { start: new Date(), end: new Date() };
		const headers = { "x-number": 123 };
		const result = applyRangeToRequestHeaders(headers, range);
		expect(result.matched).toBe(0);
	});

	it("handles non-JSON header string", () => {
		const range = { start: new Date(), end: new Date() };
		const headers = { "x-plain": "just text" };
		const result = applyRangeToRequestHeaders(headers, range);
		expect(result.matched).toBe(0);
	});

	it("handles JSON header without date keys", () => {
		const range = { start: new Date(), end: new Date() };
		const headers = { "x-json": JSON.stringify({ foo: "bar" }) };
		const result = applyRangeToRequestHeaders(headers, range);
		expect(result.matched).toBe(0);
	});

	it("handles non-URL-encoded JSON header", () => {
		const range = {
			start: new Date("2026-01-01"),
			end: new Date("2026-01-31"),
		};
		const headers = { "x-json": JSON.stringify({ StartDate: "2025-01-01" }) };
		const result = applyRangeToRequestHeaders(headers, range);
		expect(result.matched).toBeGreaterThan(0);
	});
});

describe("fetchOwaEventsByTemplate with placeholders", () => {
	it("resolves owaCanary placeholder", async () => {
		const page = {
			context: () => ({ cookies: () => Promise.resolve([{ name: "X-OWA-CANARY", value: "test-canary" }]) }),
			evaluate: vi.fn().mockResolvedValue({ Subject: "Test", Start: "2026-01-01", End: "2026-01-02", Id: "id" }),
		};

		const events = await fetchOwaEventsByTemplate({
			page,
			template: {
				url: "https://outlook.office.com/api",
				method: "POST",
				headers: { "x-owa-canary": "{{owaCanary}}" },
				body: {},
			},
			range: { start: new Date("2026-01-01"), end: new Date("2026-01-31") },
		});

		expect(events).toHaveLength(1);
	});

	it("resolves owaBearer placeholder", async () => {
		const page = {
			context: () => ({ cookies: () => Promise.resolve([]) }),
			evaluate: vi.fn()
				.mockResolvedValueOnce("Bearer test-token") // getOwaBearerToken
				.mockResolvedValueOnce({ Subject: "Test", Start: "2026-01-01", End: "2026-01-02", Id: "id" }),
		};

		const events = await fetchOwaEventsByTemplate({
			page,
			template: {
				url: "https://outlook.office.com/api",
				method: "POST",
				headers: { authorization: "Bearer {{owaBearer}}" },
				body: {},
			},
			range: { start: new Date("2026-01-01"), end: new Date("2026-01-31") },
		});

		expect(events).toHaveLength(1);
	});
});
