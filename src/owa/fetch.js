import { templateReplace } from "../utils.js";

const OWA_HOSTS = ["outlook.office.com", "outlook.cloud.microsoft"];

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
		/* c8 ignore start */
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
		/* c8 ignore stop */
		{ url: req.url, method, headers, body: req.body }
	);
}

function getPageUrl(page) {
	try {
		const raw = typeof page?.url === "function" ? page.url() : page?.url;
		return typeof raw === "string" ? raw : null;
	} catch {
		return null;
	}
}

function isOwaHost(host) {
	return !!host && OWA_HOSTS.includes(host);
}

function normalizeOwaUrlForPage(url, page) {
	if (!url) return url;
	const pageUrl = getPageUrl(page);
	if (!pageUrl) return url;
	let pageHost;
	try {
		pageHost = new URL(pageUrl).host;
	} catch {
		return url;
	}
	if (!isOwaHost(pageHost)) return url;
	try {
		const reqUrl = new URL(url);
		if (!isOwaHost(reqUrl.host) || reqUrl.host === pageHost) return url;
		reqUrl.host = pageHost;
		return reqUrl.toString();
	} catch {
		return url;
	}
}

function swapOwaHost(url) {
	try {
		const reqUrl = new URL(url);
		if (!isOwaHost(reqUrl.host)) return null;
		const next = OWA_HOSTS.find((h) => h !== reqUrl.host);
		if (!next) return null;
		reqUrl.host = next;
		return reqUrl.toString();
	} catch {
		return null;
	}
}

export async function owaFetchJsonWithFallback(page, req) {
	const normalizedUrl = normalizeOwaUrlForPage(req.url, page);
	const primaryReq = normalizedUrl === req.url ? req : { ...req, url: normalizedUrl };

	try {
		return await owaFetchJson(page, primaryReq);
	} catch (err) {
		const msg = String(err?.message ?? err);
		const statusMatch = msg.match(/OWA fetch failed: HTTP (\d+)/i);
		const status = statusMatch ? Number(statusMatch[1]) : null;
		const shouldFallback =
			msg.includes("Failed to fetch") || (status !== null && [401, 403, 404].includes(status));
		if (!shouldFallback) throw err;
		const altUrl = swapOwaHost(primaryReq.url);
		if (!altUrl || altUrl === primaryReq.url) throw err;
		console.info(`OWA fetch failed; retrying with ${altUrl}`);
		return await owaFetchJson(page, { ...primaryReq, url: altUrl });
	}
}

/**
 * Best-effort way to find OWA canary/XSRF token.
 * This may need adjustment once we inspect the real OWA runtime.
 * @param {any} page
 */
export async function getOwaCanary(page) {
	const cookieNames = ["X-OWA-CANARY", "OWA-CANARY", "XOWACANARY"];

	try {
		if (page?.context && typeof page.context === "function") {
			const context = page.context();
			if (context && typeof context.cookies === "function") {
				const cookies = await context.cookies();
				for (const c of cookies) {
					if (cookieNames.includes(c.name) && c.value) return c.value;
				}
			}
		} else if (page?.browserContext && typeof page.browserContext === "function") {
			const context = page.browserContext();
			if (context && typeof context.cookies === "function") {
				const cookies = await context.cookies();
				for (const c of cookies) {
					if (cookieNames.includes(c.name) && c.value) return c.value;
				}
			}
		} else if (page?.cookies && typeof page.cookies === "function") {
			const cookies = await page.cookies();
			for (const c of cookies) {
				if (cookieNames.includes(c.name) && c.value) return c.value;
			}
		}
	} catch {
		// ignore and try in-page lookup
	}

	return await page.evaluate(/* c8 ignore start */() => {
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
	}/* c8 ignore stop */);
}

/**
 * Best-effort way to find a bearer token for Outlook Web requests.
 * @param {any} page
 */
export async function getOwaBearerToken(page) {
	return await page.evaluate(() => {
		const tokens = [];
		const matchesTarget = (key) =>
			/https:\/\/outlook\.office\.com|https:\/\/outlook\.cloud\.microsoft/i.test(key);

		for (const key of Object.keys(localStorage || {})) {
			if (!/accesstoken/i.test(key)) continue;
			if (!matchesTarget(key)) continue;
			const raw = localStorage.getItem(key);
			if (!raw) continue;
			try {
				const parsed = JSON.parse(raw);
				if (parsed?.secret && parsed?.tokenType) {
					tokens.push(`${parsed.tokenType} ${parsed.secret}`);
				}
			} catch {
				// ignore
			}
		}

		for (const key of Object.keys(sessionStorage || {})) {
			if (!/token|auth/i.test(key)) continue;
			const raw = sessionStorage.getItem(key);
			if (!raw) continue;
			try {
				const parsed = JSON.parse(raw);
				if (parsed?.token && parsed?.tokenType) {
					tokens.push(`${parsed.tokenType} ${parsed.token}`);
				}
			} catch {
				// ignore
			}
		}

		return tokens[0] ?? null;
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
