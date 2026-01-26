import { describe, expect, it, vi } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";

import {
	scoreOwaJson,
	discoverOwaCandidatesFromLog,
	discoverOwaCandidates,
	suggestTemplate,
	_internal,
} from "./discovery.js";

describe("owa/discovery", () => {
	describe("scoreOwaJson", () => {
		it("scores based on key hints", () => {
			const json = { Subject: "Meeting", Start: "2026-01-01", End: "2026-01-02", Attendees: [] };
			const { score, keys } = scoreOwaJson(json);
			expect(score).toBeGreaterThan(0);
			expect(keys).toContain("Subject");
		});

		it("returns 0 for empty object", () => {
			const { score } = scoreOwaJson({});
			expect(score).toBe(0);
		});

		it("handles arrays", () => {
			const json = [{ Subject: "Test", Organizer: {} }];
			const { score } = scoreOwaJson(json);
			expect(score).toBeGreaterThan(0);
		});
	});

	describe("discoverOwaCandidatesFromLog", () => {
		it("extracts candidates from NDJSON log file", async () => {
			const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-discover-"));
			const logPath = path.join(dir, "log.ndjson");
			const entry = {
				url: "https://outlook.office.com/api/events",
				method: "POST",
				body: JSON.stringify({ Subject: "Meeting", Start: "2026-01-01", End: "2026-01-02", Organizer: {}, Attendees: [] }),
				requestPostData: '{"folderId": "abc"}',
			};
			await fs.writeFile(logPath, JSON.stringify(entry) + "\n");

			const candidates = await discoverOwaCandidatesFromLog({ filePath: logPath, minScore: 2 });
			expect(candidates.length).toBeGreaterThan(0);
			expect(candidates[0].url).toContain("outlook.office.com");
		});

		it("skips lines without url", async () => {
			const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-discover-"));
			const logPath = path.join(dir, "log.ndjson");
			await fs.writeFile(logPath, JSON.stringify({ body: "{}" }) + "\n");

			const candidates = await discoverOwaCandidatesFromLog({ filePath: logPath });
			expect(candidates).toHaveLength(0);
		});
	});

	describe("discoverOwaCandidates", () => {
		it("captures responses from page", async () => {
			const mockResponse = {
				url: () => "https://outlook.office.com/api/events",
				json: () => Promise.resolve({ Subject: "Test", Start: "2026-01-01", End: "2026-01-02", Organizer: {} }),
				request: () => ({ method: () => "POST" }),
			};

			const listeners = {};
			const page = {
				on: (event, handler) => { listeners[event] = handler; },
				off: vi.fn(),
			};

			const promise = discoverOwaCandidates({ page, durationMs: 10, minScore: 2 });
			// Simulate response
			setTimeout(() => listeners.response?.(mockResponse), 5);

			const candidates = await promise;
			expect(candidates.length).toBeGreaterThan(0);
		});
	});

	describe("suggestTemplate", () => {
		it("builds template from candidate", () => {
			const candidate = {
				method: "POST",
				url: "https://outlook.office.com/api?start=2026-01-01&end=2026-01-02",
				requestBody: '{"folderId": "abc"}',
			};

			const template = suggestTemplate(candidate);
			expect(template.method).toBe("POST");
			expect(template.headers["x-owa-canary"]).toBe("{{owaCanary}}");
			expect(template.body).toEqual({ folderId: "abc" });
		});

		it("handles missing requestBody", () => {
			const candidate = { method: "GET", url: "https://example.com" };
			const template = suggestTemplate(candidate);
			expect(template.body).toBeNull();
		});
	});

	describe("_internal.redactHeaders", () => {
		it("redacts sensitive headers", () => {
			const headers = {
				cookie: "session=abc",
				authorization: "Bearer xyz",
				"x-owa-canary": "tok",
				"content-type": "application/json",
			};
			const redacted = _internal.redactHeaders(headers);
			expect(redacted.cookie).toBe("<redacted>");
			expect(redacted.authorization).toBe("<redacted>");
			expect(redacted["x-owa-canary"]).toBe("<redacted>");
			expect(redacted["content-type"]).toBe("application/json");
		});
	});
});
