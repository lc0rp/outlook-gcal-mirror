import { describe, it, expect } from "vitest";

import { applyRangeToRequestBody } from "./events.js";

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
});
