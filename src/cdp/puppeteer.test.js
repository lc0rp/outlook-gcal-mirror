import { describe, expect, it, vi, beforeEach } from "vitest";

vi.mock("./imports.js");

import { connectPuppeteerOverCdp } from "./puppeteer.js";
import { importOptional } from "./imports.js";

describe("connectPuppeteerOverCdp", () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	it("connects and finds existing page matching targetUrl", async () => {
		const mockPage = { url: () => "https://outlook.office.com/", goto: vi.fn() };
		const mockBrowser = {
			pages: vi.fn().mockResolvedValue([mockPage]),
			newPage: vi.fn().mockResolvedValue(mockPage),
		};
		const mockPuppeteer = {
			connect: vi.fn().mockResolvedValue(mockBrowser),
		};

		vi.mocked(importOptional).mockResolvedValue({ default: mockPuppeteer });

		const conn = await connectPuppeteerOverCdp({ port: 9222, targetUrl: "https://outlook.office.com" });
		expect(conn.engine).toBe("puppeteer");
		expect(conn.page).toBe(mockPage);
		expect(importOptional).toHaveBeenCalledWith(["puppeteer", "puppeteer-core"]);
	});

	it("creates new page when none match", async () => {
		const mockPage = { url: () => "about:blank", goto: vi.fn() };
		const mockBrowser = {
			pages: vi.fn().mockResolvedValue([]),
			newPage: vi.fn().mockResolvedValue(mockPage),
		};
		const mockPuppeteer = {
			connect: vi.fn().mockResolvedValue(mockBrowser),
		};

		vi.mocked(importOptional).mockResolvedValue({ default: mockPuppeteer });

		const conn = await connectPuppeteerOverCdp({ port: 9222, targetUrl: "https://outlook.office.com" });
		expect(mockBrowser.newPage).toHaveBeenCalled();
		expect(mockPage.goto).toHaveBeenCalled();
	});

	it("selects non-blank page when no targetUrl match", async () => {
		const mockPage1 = { url: () => "about:blank" };
		const mockPage2 = { url: () => "https://other.example.com/" };
		const mockBrowser = {
			pages: vi.fn().mockResolvedValue([mockPage1, mockPage2]),
			newPage: vi.fn(),
		};
		const mockPuppeteer = {
			connect: vi.fn().mockResolvedValue(mockBrowser),
		};

		vi.mocked(importOptional).mockResolvedValue({ default: mockPuppeteer });

		const conn = await connectPuppeteerOverCdp({ port: 9222, targetUrl: "https://outlook.office.com" });
		expect(conn.page).toBe(mockPage2);
	});

	it("creates new page without goto when no targetUrl", async () => {
		const mockPage = { url: () => "about:blank", goto: vi.fn() };
		const mockBrowser = {
			pages: vi.fn().mockResolvedValue([]),
			newPage: vi.fn().mockResolvedValue(mockPage),
		};
		const mockPuppeteer = {
			connect: vi.fn().mockResolvedValue(mockBrowser),
		};

		vi.mocked(importOptional).mockResolvedValue({ default: mockPuppeteer });

		const conn = await connectPuppeteerOverCdp({ port: 9222 });
		expect(mockBrowser.newPage).toHaveBeenCalled();
		expect(mockPage.goto).not.toHaveBeenCalled();
	});

	it("throws if connect is missing", async () => {
		vi.mocked(importOptional).mockResolvedValue({ default: {} });
		await expect(connectPuppeteerOverCdp({ port: 9222 })).rejects.toThrow(/connect/);
	});
});
