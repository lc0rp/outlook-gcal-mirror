import { describe, expect, it, vi, beforeEach } from "vitest";

vi.mock("./imports.js");

import { connectPlaywrightOverCdp } from "./playwright.js";
import { importOptional } from "./imports.js";

describe("connectPlaywrightOverCdp", () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	it("connects and finds existing page matching targetUrl", async () => {
		const mockPage = { url: () => "https://outlook.office.com/", goto: vi.fn() };
		const mockContext = {
			pages: () => [mockPage],
			newPage: vi.fn().mockResolvedValue(mockPage),
		};
		const mockBrowser = {
			contexts: () => [mockContext],
			newContext: vi.fn().mockResolvedValue(mockContext),
		};
		const mockChromium = {
			connectOverCDP: vi.fn().mockResolvedValue(mockBrowser),
		};

		vi.mocked(importOptional).mockResolvedValue({ chromium: mockChromium });

		const conn = await connectPlaywrightOverCdp({ port: 9222, targetUrl: "https://outlook.office.com" });
		expect(conn.engine).toBe("playwright");
		expect(conn.page).toBe(mockPage);
		expect(importOptional).toHaveBeenCalledWith(["playwright", "playwright-core"]);
	});

	it("creates new page when none match", async () => {
		const mockPage = { url: () => "about:blank", goto: vi.fn() };
		const emptyContext = { pages: () => [], newPage: vi.fn().mockResolvedValue(mockPage) };
		const emptyBrowser = { contexts: () => [emptyContext], newContext: vi.fn() };
		const mockChromium = {
			connectOverCDP: vi.fn().mockResolvedValue(emptyBrowser),
		};

		vi.mocked(importOptional).mockResolvedValue({ chromium: mockChromium });

		const conn = await connectPlaywrightOverCdp({ port: 9222, targetUrl: "https://outlook.office.com" });
		expect(emptyContext.newPage).toHaveBeenCalled();
		expect(mockPage.goto).toHaveBeenCalled();
	});

	it("creates new context when no contexts exist", async () => {
		const mockPage = { url: () => "about:blank", goto: vi.fn() };
		const newContext = { pages: () => [], newPage: vi.fn().mockResolvedValue(mockPage) };
		const emptyBrowser = {
			contexts: () => [],
			newContext: vi.fn().mockResolvedValue(newContext),
		};
		const mockChromium = {
			connectOverCDP: vi.fn().mockResolvedValue(emptyBrowser),
		};

		vi.mocked(importOptional).mockResolvedValue({ chromium: mockChromium });

		const conn = await connectPlaywrightOverCdp({ port: 9222, targetUrl: "https://outlook.office.com" });
		expect(emptyBrowser.newContext).toHaveBeenCalled();
		expect(newContext.newPage).toHaveBeenCalled();
	});

	it("selects non-blank page when no targetUrl match", async () => {
		const mockPage1 = { url: () => "about:blank" };
		const mockPage2 = { url: () => "https://other.example.com/" };
		const mockContext = {
			pages: () => [mockPage1, mockPage2],
			newPage: vi.fn(),
		};
		const mockBrowser = {
			contexts: () => [mockContext],
			newContext: vi.fn(),
		};
		const mockChromium = {
			connectOverCDP: vi.fn().mockResolvedValue(mockBrowser),
		};

		vi.mocked(importOptional).mockResolvedValue({ chromium: mockChromium });

		const conn = await connectPlaywrightOverCdp({ port: 9222, targetUrl: "https://outlook.office.com" });
		expect(conn.page).toBe(mockPage2);
	});

	it("throws if connectOverCDP is missing", async () => {
		vi.mocked(importOptional).mockResolvedValue({ chromium: {} });
		await expect(connectPlaywrightOverCdp({ port: 9222 })).rejects.toThrow(/connectOverCDP/);
	});
});
