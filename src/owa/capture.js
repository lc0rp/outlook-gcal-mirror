import { extractOutlookEventsFromJson } from "./extract.js";

/**
 * Capture and parse OWA JSON responses for a short window.
 *
 * @param {{ page: any, durationMs: number, urlIncludes?: string }} opts
 */
export async function captureOwaEvents({ page, durationMs, urlIncludes = "outlook.office.com" }) {
	/** @type {any[]} */
	const jsonPayloads = [];

	const onResponse = async (response) => {
		try {
			const url = typeof response.url === "function" ? response.url() : response.url;
			if (!url || !String(url).includes(urlIncludes)) return;

			const headers = typeof response.headers === "function" ? await response.headers() : response.headers;
			const ct = headers?.["content-type"] ?? headers?.["Content-Type"];
			if (!ct || !String(ct).includes("application/json")) return;

			const json = typeof response.json === "function" ? await response.json() : null;
			if (!json) return;
			jsonPayloads.push(json);
		} catch {
			// ignore
		}
	};

	page.on("response", onResponse);
	await new Promise((r) => setTimeout(r, durationMs));
	try {
		page.off("response", onResponse);
	} catch {
		// best-effort
	}

	/** @type {ReturnType<typeof extractOutlookEventsFromJson>} */
	const allEvents = [];
	for (const json of jsonPayloads) {
		try {
			allEvents.push(...extractOutlookEventsFromJson(json));
		} catch {
			// ignore
		}
	}

	// De-dupe by sourceKey
	const seen = new Set();
	const out = [];
	for (const e of allEvents) {
		if (seen.has(e.sourceKey)) continue;
		seen.add(e.sourceKey);
		out.push(e);
	}

	return out;
}
