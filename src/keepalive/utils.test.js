import { describe, expect, it, vi } from "vitest";

import {
	parseInterval,
	validateEngine,
	validateUrlString,
	stripQueryParam,
	withCacheBuster,
	sleep,
	isMissingEngineError,
	isPlaywrightMissingBrowserError,
	isPuppeteerMissingBrowserError,
} from "./utils.js";

describe("keepalive utils", () => {
	it("parseInterval validates positive numbers", () => {
		expect(parseInterval("5")).toBe(5);
		expect(() => parseInterval("0")).toThrow(/positive/);
	});

	it("validateEngine accepts playwright/puppeteer", () => {
		expect(validateEngine("playwright")).toBe("playwright");
		expect(validateEngine("puppeteer")).toBe("puppeteer");
		expect(() => validateEngine("foo")).toThrow(/engine/);
	});

	it("validateUrlString and stripQueryParam", () => {
		expect(validateUrlString("https://example.com?a=1")).toContain("example.com");
		expect(stripQueryParam("https://example.com?a=1&b=2", "a")).toBe("https://example.com/?b=2");
		expect(stripQueryParam("not a url", "x")).toBe("not a url");
	});

	it("withCacheBuster adds _cb param", () => {
		const out = withCacheBuster("https://example.com/path");
		expect(out).toContain("_cb=");
		expect(out).toContain("example.com");
	});

	it("sleep resolves", async () => {
		vi.useFakeTimers();
		const p = sleep(10);
		vi.advanceTimersByTime(10);
		await p;
		vi.useRealTimers();
	});

	it("error helpers detect missing browsers", () => {
		expect(isMissingEngineError(new Error("Cannot find package 'playwright'"), "playwright")).toBe(true);
		expect(isPlaywrightMissingBrowserError("Playwright: executable doesn't exist")).toBe(true);
		expect(isPuppeteerMissingBrowserError("Could not find Chrome")).toBe(true);
	});
});
