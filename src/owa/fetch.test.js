import { describe, it, expect, vi } from "vitest";

import {
	applyTemplate,
	owaFetchJson,
	owaFetchJsonWithFallback,
	getOwaCanary,
	getOwaBearerToken,
} from "./fetch.js";

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

	it("handles undefined body", () => {
		const template = { method: "GET", url: "https://example.test" };
		const req = applyTemplate({ template, vars: {} });
		expect(req.body).toBeUndefined();
	});
});

describe("owaFetchJson", () => {
	it("calls page.evaluate with fetch", async () => {
		const page = {
			evaluate: vi.fn().mockResolvedValue({ data: "test" }),
		};

		const result = await owaFetchJson(page, { url: "https://outlook.office.com/api", method: "GET" });
		expect(result).toEqual({ data: "test" });
		expect(page.evaluate).toHaveBeenCalled();
	});
});

describe("owaFetchJsonWithFallback", () => {
	it("returns result on success", async () => {
		const page = {
			url: () => "https://outlook.office.com/mail",
			evaluate: vi.fn().mockResolvedValue({ ok: true }),
		};

		const result = await owaFetchJsonWithFallback(page, { url: "https://outlook.office.com/api" });
		expect(result).toEqual({ ok: true });
	});

	it("falls back to alternate host on 401", async () => {
		const page = {
			url: () => "https://outlook.office.com/mail",
			evaluate: vi
				.fn()
				.mockRejectedValueOnce(new Error("OWA fetch failed: HTTP 401"))
				.mockResolvedValueOnce({ fallback: true }),
		};

		const result = await owaFetchJsonWithFallback(page, { url: "https://outlook.office.com/api" });
		expect(result).toEqual({ fallback: true });
		expect(page.evaluate).toHaveBeenCalledTimes(2);
	});

	it("throws without fallback on other errors", async () => {
		const page = {
			url: () => "https://outlook.office.com/mail",
			evaluate: vi.fn().mockRejectedValue(new Error("Network error")),
		};

		await expect(
			owaFetchJsonWithFallback(page, { url: "https://outlook.office.com/api" })
		).rejects.toThrow("Network error");
	});
});

describe("getOwaCanary", () => {
	it("gets canary from context cookies (playwright)", async () => {
		const page = {
			context: () => ({
				cookies: () => Promise.resolve([{ name: "X-OWA-CANARY", value: "canary123" }]),
			}),
			evaluate: vi.fn(),
		};

		const canary = await getOwaCanary(page);
		expect(canary).toBe("canary123");
	});

	it("gets canary from browserContext cookies (puppeteer)", async () => {
		const page = {
			browserContext: () => ({
				cookies: () => Promise.resolve([{ name: "OWA-CANARY", value: "pup-canary" }]),
			}),
			evaluate: vi.fn(),
		};

		const canary = await getOwaCanary(page);
		expect(canary).toBe("pup-canary");
	});

	it("gets canary from page.cookies", async () => {
		const page = {
			cookies: () => Promise.resolve([{ name: "XOWACANARY", value: "direct-canary" }]),
			evaluate: vi.fn(),
		};

		const canary = await getOwaCanary(page);
		expect(canary).toBe("direct-canary");
	});

	it("falls back to page.evaluate", async () => {
		const page = {
			evaluate: vi.fn().mockResolvedValue("eval-canary"),
		};

		const canary = await getOwaCanary(page);
		expect(canary).toBe("eval-canary");
	});
});

describe("getOwaBearerToken", () => {
	it("calls page.evaluate", async () => {
		const page = {
			evaluate: vi.fn().mockResolvedValue("Bearer abc123"),
		};

		const token = await getOwaBearerToken(page);
		expect(token).toBe("Bearer abc123");
		expect(page.evaluate).toHaveBeenCalled();
	});
});

describe("owaFetchJson error handling", () => {
	it("handles page.evaluate throwing", async () => {
		const page = {
			evaluate: vi.fn().mockRejectedValue(new Error("eval failed")),
		};

		await expect(owaFetchJson(page, { url: "https://test.com" })).rejects.toThrow("eval failed");
	});
});

describe("owaFetchJsonWithFallback edge cases", () => {
	it("does not fallback on non-owa URLs", async () => {
		const page = {
			url: () => "https://example.com/",
			evaluate: vi.fn().mockRejectedValue(new Error("OWA fetch failed: HTTP 401")),
		};

		await expect(
			owaFetchJsonWithFallback(page, { url: "https://example.com/api" })
		).rejects.toThrow("HTTP 401");
	});

	it("normalizes URL to match page host", async () => {
		const page = {
			url: () => "https://outlook.cloud.microsoft/mail",
			evaluate: vi.fn().mockResolvedValue({ normalized: true }),
		};

		const result = await owaFetchJsonWithFallback(page, { url: "https://outlook.office.com/api" });
		expect(result).toEqual({ normalized: true });
		// The URL should have been normalized to outlook.cloud.microsoft
		const callUrl = page.evaluate.mock.calls[0][1].url;
		expect(callUrl).toContain("outlook.cloud.microsoft");
	});
});

describe("getOwaCanary edge cases", () => {
	it("handles context.cookies throwing", async () => {
		const page = {
			context: () => ({
				cookies: () => Promise.reject(new Error("no cookies")),
			}),
			evaluate: vi.fn().mockResolvedValue("fallback-canary"),
		};

		const canary = await getOwaCanary(page);
		expect(canary).toBe("fallback-canary");
	});

	it("handles null context", async () => {
		const page = {
			context: null,
			evaluate: vi.fn().mockResolvedValue("eval-canary"),
		};

		const canary = await getOwaCanary(page);
		expect(canary).toBe("eval-canary");
	});
});
