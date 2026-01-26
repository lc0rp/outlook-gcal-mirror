import { describe, expect, it } from "vitest";

import {
	assertNonEmptyString,
	validateAbsoluteUrl,
	validateCdpPort,
	errorMessage,
	stableStringify,
	sortKeysDeep,
	sha1Hex,
	templateReplace,
} from "./utils.js";

describe("utils", () => {
	it("assertNonEmptyString returns string and throws on empty", () => {
		expect(assertNonEmptyString(" hi ")).toBe(" hi ");
		expect(() => assertNonEmptyString("")).toThrow(/non-empty/);
		expect(() => assertNonEmptyString("   ")).toThrow(/non-empty/);
	});

	it("validateAbsoluteUrl normalizes and rejects invalid", () => {
		expect(validateAbsoluteUrl("https://example.com/abc")).toBe("https://example.com/abc");
		expect(() => validateAbsoluteUrl("not-a-url")).toThrow(/absolute URL/);
	});

	it("validateCdpPort handles defaults and bounds", () => {
		expect(validateCdpPort(undefined)).toBe(9222);
		expect(validateCdpPort("9223")).toBe(9223);
		expect(() => validateCdpPort(0)).toThrow(/between 1 and 65535/);
	});

	it("errorMessage returns message for Error", () => {
		expect(errorMessage(new Error("boom"))).toBe("boom");
		expect(errorMessage("yo")).toBe("yo");
	});

	it("stableStringify sorts keys deeply", () => {
		const input = { b: 2, a: { d: 4, c: 3 } };
		expect(stableStringify(input)).toBe('{"a":{"c":3,"d":4},"b":2}');
		const sorted = sortKeysDeep(input);
		expect(sorted).toEqual({ a: { c: 3, d: 4 }, b: 2 });
	});

	it("sha1Hex produces stable hash", () => {
		const hash = sha1Hex({ a: 1, b: 2 });
		expect(hash).toMatch(/^[a-f0-9]{40}$/);
		const hash2 = sha1Hex({ b: 2, a: 1 });
		expect(hash2).toBe(hash);
	});

	it("templateReplace swaps placeholders", () => {
		const out = templateReplace("Hello {{name}}", { name: "User" });
		expect(out).toBe("Hello User");
	});
});
