import { captureOwaEvents } from "./capture.js";
import { applyTemplate, getOwaCanary, owaFetchJson } from "./fetch.js";
import { extractOutlookEventsFromJson } from "./extract.js";

/**
 * @typedef {{
 *   url: string,
 *   method: string,
 *   headers?: Record<string,string>,
 *   body?: any,
 * }} OwaRequestTemplate
 */

/**
 * @param {OwaRequestTemplate} template
 * @param {string} placeholder
 */
function templateContainsPlaceholder(template, placeholder) {
	const needle = `{{${placeholder}}}`;
	if (template.url?.includes(needle)) return true;
	if (template.method?.includes(needle)) return true;
	for (const value of Object.values(template.headers ?? {})) {
		if (String(value).includes(needle)) return true;
	}
	if (template.body === undefined || template.body === null) return false;
	if (typeof template.body === "string") return template.body.includes(needle);
	try {
		return JSON.stringify(template.body).includes(needle);
	} catch {
		return false;
	}
}

/**
 * Fetch OWA JSON via a request template and extract events.
 *
 * Template supports placeholders:
 * - `{{start}}` / `{{end}}` (ISO strings)
 * - `{{owaCanary}}` (best-effort; pulled from cookie or globals)
 * - plus any `templateVars` keys
 *
 * @param {{
 *   page: any,
 *   template: OwaRequestTemplate,
 *   range: { start: Date, end: Date },
 *   templateVars?: Record<string,string>,
 * }} opts
 */
export async function fetchOwaEventsByTemplate({ page, template, range, templateVars }) {
	if (!template || typeof template !== "object") {
		throw new Error("Missing owaRequestTemplate");
	}
	if (!template.url || !template.method) {
		throw new Error("owaRequestTemplate must include url + method");
	}

	const vars = {
		start: range.start.toISOString(),
		end: range.end.toISOString(),
		...(templateVars ?? {}),
	};

	const needsCanary = templateContainsPlaceholder(template, "owaCanary");
	if (needsCanary) {
		const canary = await getOwaCanary(page);
		if (!canary) {
			throw new Error(
				"OWA canary was required (template uses {{owaCanary}}) but could not be found. Make sure the Outlook tab is fully loaded and logged in."
			);
		}
		vars.owaCanary = canary;
	} else {
		const canary = await getOwaCanary(page);
		if (canary) vars.owaCanary = canary;
	}

	const req = applyTemplate({ template, vars });
	const json = await owaFetchJson(page, req);
	return extractOutlookEventsFromJson(json);
}

/**
 * Unified helper for future use.
 *
 * @param {{
 *   mode: "capture" | "template",
 *   page: any,
 *   capture?: { durationMs: number, urlIncludes?: string | null },
 *   template?: { template: OwaRequestTemplate, range: { start: Date, end: Date }, templateVars?: Record<string,string> },
 * }} opts
 */
export async function getOwaEvents(opts) {
	if (opts.mode === "capture") {
		if (!opts.capture) throw new Error("Missing capture options");
		return await captureOwaEvents({
			page: opts.page,
			durationMs: opts.capture.durationMs,
			urlIncludes: opts.capture.urlIncludes,
		});
	}

	if (!opts.template) throw new Error("Missing template options");
	return await fetchOwaEventsByTemplate({
		page: opts.page,
		template: opts.template.template,
		range: opts.template.range,
		templateVars: opts.template.templateVars,
	});
}
