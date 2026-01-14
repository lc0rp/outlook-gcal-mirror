import { templateReplace } from "../utils.js";

/**
 * @typedef {object} OwaCandidate
 * @property {string} method
 * @property {string} url
 * @property {string[]} interestingKeys
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
 * @param {{ page: any, durationMs: number, minScore?: number, includeHeaders?: boolean }} opts
 */
export async function discoverOwaCandidates({ page, durationMs, minScore = 3, includeHeaders = false }) {
	/** @type {OwaCandidate[]} */
	const candidates = [];

	const onResponse = async (response) => {
		try {
			const url = typeof response.url === "function" ? response.url() : response.url;
			if (!url || !String(url).includes("outlook.office.com")) return;

			const headers = typeof response.headers === "function" ? await response.headers() : response.headers;
			const ct = headers?.["content-type"] ?? headers?.["Content-Type"];
			if (!ct || !String(ct).includes("application/json")) return;

			const json = typeof response.json === "function" ? await response.json() : null;
			const { score, keys } = scoreOwaJson(json);
			if (score < minScore) return;

			const request = typeof response.request === "function" ? response.request() : null;
			const method = request && typeof request.method === "function" ? request.method() : "GET";

			candidates.push({
				method,
				url: String(url),
				interestingKeys: keys.slice(0, 20),
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

	// De-dupe by method+url
	const seen = new Set();
	const deduped = [];
	for (const c of candidates) {
		const key = `${c.method} ${c.url}`;
		if (seen.has(key)) continue;
		seen.add(key);
		deduped.push(c);
	}

	return deduped;
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
