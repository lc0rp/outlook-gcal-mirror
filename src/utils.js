import crypto from "node:crypto";

/**
 * @param {string} message
 */
export function assertNonEmptyString(message) {
	if (typeof message !== "string" || message.trim().length === 0) {
		throw new Error("Expected a non-empty string");
	}
	return message;
}

/**
 * @param {string} value
 * @returns {string}
 */
export function validateAbsoluteUrl(value) {
	try {
		return new URL(value).toString();
	} catch {
		throw new Error("Expected an absolute URL (e.g. https://example.com)");
	}
}

/**
 * @param {unknown} value
 * @returns {number}
 */
export function validateCdpPort(value) {
	if (value === undefined || value === null || value === "") return 9222;
	const n = Number(value);
	if (!Number.isInteger(n) || n <= 0 || n > 65535) {
		throw new Error("CDP port must be an integer between 1 and 65535");
	}
	return n;
}

/**
 * @param {unknown} err
 * @returns {string}
 */
export function errorMessage(err) {
	if (err instanceof Error) return err.message;
	return String(err);
}

/**
 * Stable JSON stringify for hashing.
 * @param {unknown} value
 */
export function stableStringify(value) {
	return JSON.stringify(sortKeysDeep(value));
}

/**
 * @param {unknown} value
 * @returns {unknown}
 */
export function sortKeysDeep(value) {
	if (Array.isArray(value)) return value.map(sortKeysDeep);
	if (value && typeof value === "object") {
		/** @type {Record<string, unknown>} */
		const obj = value;
		const out = {};
		for (const key of Object.keys(obj).sort()) {
			out[key] = sortKeysDeep(obj[key]);
		}
		return out;
	}
	return value;
}

/**
 * @param {unknown} value
 */
export function sha1Hex(value) {
	return crypto.createHash("sha1").update(stableStringify(value)).digest("hex");
}

/**
 * Minimal templating: replaces `{{key}}` with string values.
 * @param {string} template
 * @param {Record<string, string>} vars
 */
export function templateReplace(template, vars) {
	let out = template;
	for (const [k, v] of Object.entries(vars)) {
		out = out.split(`{{${k}}}`).join(v);
	}
	return out;
}
