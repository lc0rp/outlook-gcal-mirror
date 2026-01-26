import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";

import { resolveUserDataDir } from "./keepalive.js";

describe("keepalive", () => {
	describe("resolveUserDataDir", () => {
		it("returns explicit dir when provided", () => {
			expect(resolveUserDataDir("/custom/path")).toBe("/custom/path");
			expect(resolveUserDataDir("  /trimmed  ")).toBe("/trimmed");
		});

		it("returns default dir when no explicit dir", () => {
			const result = resolveUserDataDir(null);
			expect(result).toContain(".config/outlook-gcal-mirror/chrome");
		});

		it("returns default dir for empty string", () => {
			const result = resolveUserDataDir("");
			expect(result).toContain("chrome");
		});
	});
});
