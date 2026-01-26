import { describe, expect, it, vi } from "vitest";
import { captureOwaEvents } from "./capture.js";

describe("owa/capture", () => {
	it("captures and parses OWA responses", async () => {
		const eventJson = {
			Subject: "Team Meeting",
			Start: "2026-01-15T10:00:00Z",
			End: "2026-01-15T11:00:00Z",
			Id: "AAMk-test",
		};

		const mockResponse = {
			url: () => "https://outlook.office.com/api/events",
			json: () => Promise.resolve(eventJson),
		};

		const listeners = {};
		const page = {
			on: (event, handler) => { listeners[event] = handler; },
			off: vi.fn(),
		};

		const promise = captureOwaEvents({ page, durationMs: 20 });
		// Simulate responses
		setTimeout(() => listeners.response?.(mockResponse), 5);
		setTimeout(() => listeners.response?.(mockResponse), 10); // duplicate

		const events = await promise;
		expect(events).toHaveLength(1); // de-duped
		expect(events[0].subject).toBe("Team Meeting");
		expect(events[0].sourceId).toBe("AAMk-test");
	});

	it("filters by urlIncludes", async () => {
		const listeners = {};
		const page = {
			on: (event, handler) => { listeners[event] = handler; },
			off: vi.fn(),
		};

		const promise = captureOwaEvents({ page, durationMs: 20, urlIncludes: "outlook.office.com" });
		setTimeout(() => {
			listeners.response?.({
				url: () => "https://other.example.com/api",
				json: () => Promise.resolve({ Subject: "Other" }),
			});
		}, 5);

		const events = await promise;
		expect(events).toHaveLength(0);
	});

	it("handles response.url as property", async () => {
		const listeners = {};
		const page = {
			on: (event, handler) => { listeners[event] = handler; },
			off: vi.fn(),
		};

		const promise = captureOwaEvents({ page, durationMs: 20 });
		setTimeout(() => {
			listeners.response?.({
				url: "https://outlook.office.com/api",
				json: () => Promise.resolve({ Subject: "Test", Start: "2026-01-01", End: "2026-01-02", Id: "id-1" }),
			});
		}, 5);

		const events = await promise;
		expect(events).toHaveLength(1);
	});

	it("handles json() throwing", async () => {
		const listeners = {};
		const page = {
			on: (event, handler) => { listeners[event] = handler; },
			off: vi.fn(),
		};

		const promise = captureOwaEvents({ page, durationMs: 20 });
		setTimeout(() => {
			listeners.response?.({
				url: () => "https://outlook.office.com/api",
				json: () => Promise.reject(new Error("not JSON")),
			});
		}, 5);

		const events = await promise;
		expect(events).toHaveLength(0); // gracefully handled
	});
});
