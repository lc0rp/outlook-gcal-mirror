import { describe, expect, it, vi } from "vitest";

import {
	buildChromiumArgs,
	createSession,
	launchEngine,
	launchPlaywright,
	launchPuppeteer,
	normalizePort,
	withCause,
} from "./engines.js";

describe("keepalive engines", () => {
	it("normalizePort and buildChromiumArgs", () => {
		expect(normalizePort(null)).toBe(null);
		expect(normalizePort("9222")).toBe(9222);
		expect(() => normalizePort(0)).toThrow(/CDP port/);
		expect(buildChromiumArgs({ cdpPort: 9222 })).toEqual([
			"--remote-debugging-port=9222",
			"--remote-debugging-address=127.0.0.1",
		]);
	});

	it("withCause appends message", () => {
		expect(withCause("oops", new Error("boom"))).toContain("Cause: boom");
		expect(withCause("oops", "no")).toBe("oops");
	});

	it("createSession proxies page methods", async () => {
		const page = {
			goto: vi.fn().mockResolvedValue("ok"),
			reload: vi.fn().mockResolvedValue("reloaded"),
			url: vi.fn().mockReturnValue("https://example.com"),
		};
		const browser = { close: vi.fn() };
		const session = createSession("playwright", page, browser, 9222);
		await session.goto("https://example.com");
		await session.reload();
		expect(await session.currentUrl()).toBe("https://example.com");
		await session.close();
		expect(page.goto).toHaveBeenCalled();
		expect(browser.close).toHaveBeenCalled();
	});

	it("launchPlaywright falls back across channels", async () => {
		const page = {};
		const context = { newPage: vi.fn().mockResolvedValue(page) };
		const browser = { newContext: vi.fn().mockResolvedValue(context) };
		const chromium = {
			launch: vi.fn(async (opts) => {
				if (opts.channel === "chrome") {
					throw new Error("chromium distribution 'chrome' is not found");
				}
				return browser;
			}),
		};
		const _import = vi.fn(async () => ({ chromium }));

		const session = await launchPlaywright({ headless: true, cdpPort: 9222, userDataDir: null, _import });
		expect(session.page).toBe(page);
		expect(chromium.launch).toHaveBeenCalledTimes(2); // chrome fails, then null
	});

	it("launchPlaywright uses persistent context when userDataDir is set", async () => {
		const page = { url: () => "about:blank" };
		const context = { pages: () => [page], newPage: vi.fn() };
		const chromium = {
			launch: vi.fn(),
			launchPersistentContext: vi.fn(async () => context),
		};
		const _import = vi.fn(async () => ({ chromium }));

		const session = await launchPlaywright({ headless: true, cdpPort: 9222, userDataDir: "/tmp", _import });
		expect(session.page).toBe(page);
		expect(chromium.launchPersistentContext).toHaveBeenCalled();
	});

	it("launchPuppeteer falls back when chrome missing", async () => {
		const browser = { newPage: vi.fn().mockResolvedValue({}) };
		const puppeteer = {
			launch: vi.fn(async (opts) => {
				if (opts.channel === "chrome") {
					throw new Error("Could not find Chrome");
				}
				return browser;
			}),
		};
		const _import = vi.fn(async () => ({ default: puppeteer }));
		const session = await launchPuppeteer({ headless: true, cdpPort: null, userDataDir: null, _import });
		expect(session.engine).toBe("puppeteer");
		expect(puppeteer.launch).toHaveBeenCalledTimes(2);
	});

	it("launchEngine rejects unknown engine", async () => {
		await expect(launchEngine("safari")).rejects.toThrow(/Unknown engine/);
	});

	it("launchEngine routes to playwright", async () => {
		const page = {};
		const context = { newPage: vi.fn().mockResolvedValue(page) };
		const browser = { newContext: vi.fn().mockResolvedValue(context) };
		const chromium = { launch: vi.fn().mockResolvedValue(browser) };
		const _import = vi.fn(async () => ({ chromium }));

		const session = await launchEngine("playwright", { headless: true, _import });
		expect(session.engine).toBe("playwright");
	});

	it("launchEngine routes to puppeteer", async () => {
		const browser = { newPage: vi.fn().mockResolvedValue({}) };
		const puppeteer = { launch: vi.fn().mockResolvedValue(browser) };
		const _import = vi.fn(async () => ({ default: puppeteer }));

		const session = await launchEngine("puppeteer", { headless: true, _import });
		expect(session.engine).toBe("puppeteer");
	});

	it("launchPlaywright throws when import fails", async () => {
		const _import = vi.fn().mockRejectedValue(new Error("module not found"));
		await expect(launchPlaywright({ headless: true, cdpPort: null, _import }))
			.rejects.toThrow(/Failed to import 'playwright'/);
	});

	it("launchPlaywright throws when chromium is missing", async () => {
		const _import = vi.fn().mockResolvedValue({});
		await expect(launchPlaywright({ headless: true, cdpPort: null, _import }))
			.rejects.toThrow(/chromium.*not found/i);
	});

	it("launchPuppeteer throws when import fails", async () => {
		const _import = vi.fn().mockRejectedValue(new Error("module not found"));
		await expect(launchPuppeteer({ headless: true, cdpPort: null, _import }))
			.rejects.toThrow(/Failed to import 'puppeteer'/);
	});

	it("launchPuppeteer throws when launch is missing", async () => {
		const _import = vi.fn().mockResolvedValue({ default: {} });
		await expect(launchPuppeteer({ headless: true, cdpPort: null, _import }))
			.rejects.toThrow(/no.*launch/i);
	});

	it("launchPlaywright handles new page from persistent context", async () => {
		const page = { url: () => "https://example.com" };
		const context = { pages: () => [], newPage: vi.fn().mockResolvedValue(page) };
		const chromium = { launchPersistentContext: vi.fn().mockResolvedValue(context) };
		const _import = vi.fn(async () => ({ chromium }));

		const session = await launchPlaywright({ headless: true, cdpPort: null, userDataDir: "/tmp/data", _import });
		expect(session.page).toBe(page);
		expect(context.newPage).toHaveBeenCalled();
	});
});
