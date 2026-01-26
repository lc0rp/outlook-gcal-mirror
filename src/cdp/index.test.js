import { describe, expect, it, vi, beforeEach } from "vitest";

vi.mock("./playwright.js", () => ({
	connectPlaywrightOverCdp: vi.fn().mockResolvedValue({ engine: "playwright", page: {} }),
}));
vi.mock("./puppeteer.js", () => ({
	connectPuppeteerOverCdp: vi.fn().mockResolvedValue({ engine: "puppeteer", page: {} }),
}));

import { connectOverCdp } from "./index.js";
import { connectPlaywrightOverCdp } from "./playwright.js";
import { connectPuppeteerOverCdp } from "./puppeteer.js";

describe("connectOverCdp", () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	it("routes to playwright", async () => {
		const res = await connectOverCdp({ engine: "playwright", port: 9222 });
		expect(res.engine).toBe("playwright");
		expect(connectPlaywrightOverCdp).toHaveBeenCalledWith({ port: 9222, targetUrl: undefined });
	});

	it("routes to puppeteer", async () => {
		const res = await connectOverCdp({ engine: "puppeteer", port: 9222, targetUrl: "https://example.com" });
		expect(res.engine).toBe("puppeteer");
		expect(connectPuppeteerOverCdp).toHaveBeenCalledWith({ port: 9222, targetUrl: "https://example.com" });
	});

	it("throws on unknown engine", async () => {
		await expect(connectOverCdp({ engine: "safari", port: 9222 })).rejects.toThrow(/Unknown engine/);
	});
});
