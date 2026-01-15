import { templateReplace } from "../utils.js";

/**
 * Run a fetch inside the page context, so OWA cookies/session are used automatically.
 *
 * @param {any} page
 * @param {{ url: string, method?: string, headers?: Record<string,string>, body?: any }} req
 */
export async function owaFetchJson(page, req) {
	const method = req.method ?? "GET";
	const headers = req.headers ?? {};

	return await page.evaluate(
		async ({ url, method, headers, body }) => {
			const res = await fetch(url, {
				method,
				headers,
				body:
					body === undefined
						? undefined
						: typeof body === "string"
							? body
							: JSON.stringify(body),
				credentials: "include",
			});

			const ct = res.headers.get("content-type") || "";
			const text = await res.text();

			if (!res.ok) {
				const preview = text.slice(0, 500);
				throw new Error(
					`OWA fetch failed: HTTP ${res.status} (${ct})${preview ? `: ${preview}` : ""}`
				);
			}

			if (!text) return null;

			try {
				return JSON.parse(text);
			} catch {
				const preview = text.slice(0, 500);
				throw new Error(
					`OWA fetch returned non-JSON body (ct=${ct})${preview ? `: ${preview}` : ""}`
				);
			}
		},
		{ url: req.url, method, headers, body: req.body }
	);
}

/**
 * Best-effort way to find OWA canary/XSRF token.
 * This may need adjustment once we inspect the real OWA runtime.
 * @param {any} page
 */
export async function getOwaCanary(page) {
	return await page.evaluate(() => {
		const cookie = document.cookie || "";
		// Common cookie key in OWA variants.
		for (const key of ["X-OWA-CANARY", "OWA-CANARY", "XOWACANARY"]) {
			const match = cookie.match(new RegExp(`${key}=([^;]+)`));
			if (match?.[1]) return decodeURIComponent(match[1]);
		}

		// Fallback: try known globals.
		const w = /** @type {any} */ (window);
		return (
			w?.owa?.canary ||
			w?.owaSettings?.canary ||
			w?.__owa?.canary ||
			w?.__OWA_CANARY__ ||
			null
		);
	});
}

/**
 * @param {{ template: { url: string, method: string, headers?: Record<string,string>, body?: any }, vars: Record<string,string> }} opts
 */
export function applyTemplate({ template, vars }) {
	const headers = {};
	for (const [k, v] of Object.entries(template.headers ?? {})) {
		headers[k] = templateReplace(String(v), vars);
	}

	return {
		url: templateReplace(template.url, vars),
		method: template.method,
		headers,
		body:
			template.body === undefined || template.body === null
				? undefined
				: templateReplace(
					typeof template.body === "string" ? template.body : JSON.stringify(template.body),
					vars
				),
	};
}
