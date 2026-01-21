import { describe, it, expect } from "vitest";

import { applyRangeToRequestBody, applyRangeToRequestUrl } from "./events.js";

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
});
