import { describe, it, expect } from "vitest";

import { applyTemplate } from "./fetch.js";

describe("applyTemplate", () => {
	it("replaces placeholders in url/headers/body", () => {
		const template = {
			method: "POST",
			url: "https://example.test/api?start={{start}}&end={{end}}",
			headers: {
				accept: "application/json",
				"x-owa-canary": "{{owaCanary}}",
				"x-custom": "{{custom}}",
			},
			body: {
				start: "{{start}}",
				end: "{{end}}",
				folderId: "{{folderId}}",
			},
		};

		const req = applyTemplate({
			template,
			vars: {
				start: "2026-01-01T00:00:00.000Z",
				end: "2026-01-02T00:00:00.000Z",
				owaCanary: "CANARY",
				custom: "X",
				folderId: "FOLDER",
			},
		});

		expect(req.method).toBe("POST");
		expect(req.url).toBe(
			"https://example.test/api?start=2026-01-01T00:00:00.000Z&end=2026-01-02T00:00:00.000Z"
		);
		expect(req.headers["x-owa-canary"]).toBe("CANARY");
		expect(req.headers["x-custom"]).toBe("X");
		expect(req.body).toBe(
			'{"start":"2026-01-01T00:00:00.000Z","end":"2026-01-02T00:00:00.000Z","folderId":"FOLDER"}'
		);
	});
});
