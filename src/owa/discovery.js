import { createReadStream } from "node:fs";
import { createInterface } from "node:readline";

import { templateReplace } from "../utils.js";

/**
 * @typedef {object} OwaCandidate
 * @property {string} method
 * @property {string} url
 * @property {string[]} interestingKeys
 * @property {number} score
 */

const DEFAULT_KEY_HINTS = [
	"Subject",
	"Attendees",
	"Organizer",
	"Start",
	"End",
	"StartTime",
	"EndTime",
	"Location",
	"OnlineMeeting",
	"Body",
];

/**
 * @param {unknown} value
 * @returns {string[]}
 */
function collectKeys(value) {
	if (!value || typeof value !== "object") return [];
	if (Array.isArray(value)) {
		for (const item of value) {
			const keys = collectKeys(item);
			if (keys.length) return keys;
		}
		return [];
	}
	return Object.keys(value);
}

/**
 * Best-effort score to identify "event detail" JSON.
 * @param {unknown} json
 */
export function scoreOwaJson(json) {
	const keys = collectKeys(json);
	const lower = keys.map((k) => k.toLowerCase());

	let score = 0;
	for (const hint of DEFAULT_KEY_HINTS) {
		if (lower.includes(hint.toLowerCase())) score += 1;
	}

	return { score, keys };
}

function parseJsonFromText(text) {
	if (!text || typeof text !== "string") return null;
	let trimmed = text.trim();
	if (!trimmed) return null;

	const firstBrace = trimmed.search(/[\[{]/);
	if (firstBrace > 0) {
		trimmed = trimmed.slice(firstBrace);
	}

	try {
		return JSON.parse(trimmed);
	} catch {
		return null;
	}
}

/**
 * @param {{ filePath: string, minScore?: number, urlIncludes?: string | null }} opts
 */
export async function discoverOwaCandidatesFromLog({
	filePath,
	minScore = 3,
	urlIncludes = "outlook.office.com",
}) {
	/** @type {OwaCandidate[]} */
	const candidates = [];

	const rl = createInterface({
		input: createReadStream(filePath),
		crlfDelay: Infinity,
	});

	for await (const line of rl) {
		const trimmed = line.trim();
		if (!trimmed) continue;
		let entry;
		try {
			entry = JSON.parse(trimmed);
		} catch {
			continue;
		}

		const url = entry?.url ? String(entry.url) : "";
		if (!url) continue;
		if (urlIncludes && !url.includes(urlIncludes)) continue;

		const body = entry?.body ?? entry?.bodyText ?? entry?.responseBody ?? null;
		if (!body || typeof body !== "string") continue;

		const json = parseJsonFromText(body);
		if (!json) continue;

		const { score, keys } = scoreOwaJson(json);
		if (score < minScore) continue;

		candidates.push({
			method: entry?.method ? String(entry.method) : "GET",
			url,
			interestingKeys: keys.slice(0, 20),
			score,
		});
	}

	const bestByKey = new Map();
	for (const c of candidates) {
		const key = `${c.method} ${c.url}`;
		const prev = bestByKey.get(key);
		if (!prev || c.score > prev.score) {
			bestByKey.set(key, c);
		}
	}

	return Array.from(bestByKey.values());
}

/**
 * @param {Record<string, string>} headers
 */
function redactHeaders(headers) {
	const out = {};
	for (const [k, v] of Object.entries(headers)) {
		const lk = k.toLowerCase();
		if (lk === "cookie" || lk === "authorization") {
			out[k] = "<redacted>";
			continue;
		}
		if (lk.includes("canary") || lk.includes("xsrf") || lk.includes("token")) {
			out[k] = "<redacted>";
			continue;
		}
		out[k] = v;
	}
	return out;
}

/**
 * Observe network activity and print candidate JSON endpoints.
 * Works for both Playwright and Puppeteer pages (duck-typed).
 *
 * @param {{ page: any, durationMs: number, minScore?: number, urlIncludes?: string | null }} opts
 */
export async function discoverOwaCandidates({ page, durationMs, minScore = 3, urlIncludes = "outlook.office.com" }) {
	/** @type {OwaCandidate[]} */
	const candidates = [];

	const onResponse = async (response) => {
		try {
			const url = typeof response.url === "function" ? response.url() : response.url;
			if (!url) return;
			if (urlIncludes && !String(url).includes(urlIncludes)) return;

			const json = typeof response.json === "function" ? await response.json() : null;
			const { score, keys } = scoreOwaJson(json);
			if (score < minScore) return;

			const request = typeof response.request === "function" ? response.request() : null;
			const method = request && typeof request.method === "function" ? request.method() : "GET";

			candidates.push({
				method,
				url: String(url),
				interestingKeys: keys.slice(0, 20),
				score,
			});
		} catch {
			// ignore
		}
	};

	page.on("response", onResponse);
	await new Promise((r) => setTimeout(r, durationMs));

	try {
		page.off("response", onResponse);
	} catch {
		// some implementations may not support off
	}

	// De-dupe by method+url (keep highest score)
	const bestByKey = new Map();
	for (const c of candidates) {
		const key = `${c.method} ${c.url}`;
		const prev = bestByKey.get(key);
		if (!prev || c.score > prev.score) {
			bestByKey.set(key, c);
		}
	}

	return Array.from(bestByKey.values());
}

/**
 * Render a suggested request template for config.
 * @param {{ method: string, url: string }} candidate
 */
export function suggestTemplate(candidate) {
	return {
		method: candidate.method,
		url: templateReplace(candidate.url, {
			start: "{{start}}",
			end: "{{end}}",
		}),
		headers: {
			accept: "application/json",
			"content-type": "application/json",
			"x-owa-canary": "{{owaCanary}}",
		},
		body: "{{body}}",
	};
}

export const _internal = { redactHeaders };
