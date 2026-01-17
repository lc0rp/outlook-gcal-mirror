import { captureOwaEvents } from "./capture.js";
import { applyTemplate, getOwaBearerToken, getOwaCanary, owaFetchJson } from "./fetch.js";
import { extractOutlookEventIdsFromJson, extractOutlookEventsFromJson } from "./extract.js";

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
 * @param {any} body
 */
function hasEventIds(body) {
	if (!body || typeof body !== "object") return false;
	if (body.EventIds && Array.isArray(body.EventIds)) return true;
	if (body.Body && typeof body.Body === "object" && Array.isArray(body.Body.EventIds)) return true;
	return false;
}

/**
 * @param {any} body
 * @param {Array<{__type: string, Id: string}>} eventIds
 */
function injectEventIdsIntoBody(body, eventIds) {
	if (!body || typeof body !== "object") return body;
	const cloned = JSON.parse(JSON.stringify(body));
	if (cloned.Body && typeof cloned.Body === "object") {
		cloned.Body.EventIds = eventIds;
		return cloned;
	}
	cloned.EventIds = eventIds;
	return cloned;
}

/**
 * @param {string} value
 */
function tryParseJsonString(value) {
	if (!value || typeof value !== "string") return null;
	try {
		return JSON.parse(value);
	} catch {
		return null;
	}
}

/**
 * @param {any[]} items
 * @param {number} size
 */
function chunk(items, size) {
	if (!Array.isArray(items) || items.length === 0) return [];
	const out = [];
	for (let i = 0; i < items.length; i += size) {
		out.push(items.slice(i, i + size));
	}
	return out;
}

/**
 * @param {{ page: any, template: OwaRequestTemplate, range: { start: Date, end: Date }, templateVars?: Record<string,string> }} opts
 */
async function resolveTemplateVars({ page, template, range, templateVars }) {
	const vars = {
		start: range.start.toISOString(),
		end: range.end.toISOString(),
		...(templateVars ?? {}),
	};

	const needsCanary = templateContainsPlaceholder(template, "owaCanary");
	if (needsCanary && !vars.owaCanary) {
		const canary = await getOwaCanary(page);
		if (canary) {
			vars.owaCanary = canary;
		} else {
			// Some tenants rely on bearer auth headers instead of canary cookies.
			vars.owaCanary = "";
		}
	}

	const needsBearer = templateContainsPlaceholder(template, "owaBearer");
	if (needsBearer && !vars.owaBearer) {
		const bearer = await getOwaBearerToken(page);
		vars.owaBearer = bearer ?? "";
	}

	return vars;
}

/**
 * @param {{ template: OwaRequestTemplate, vars: Record<string,string>, eventIds: string[] }} opts
 */
function buildEventDetailsRequest({ template, vars, eventIds }) {
	if (!template || typeof template !== "object") {
		throw new Error("Missing owaEventRequestTemplate");
	}
	if (!template.url || !template.method) {
		throw new Error("owaEventRequestTemplate must include url + method");
	}

	const payload = eventIds.map((id) => ({ __type: "ItemId:#Exchange", Id: id }));
	const varsWithIds = { ...vars, eventIds: JSON.stringify(payload) };

	let body = template.body;
	if (body === undefined || body === null) {
		throw new Error("owaEventRequestTemplate must include a request body (captured requestPostData)");
	}

	if (typeof body === "string") {
		if (!body.includes("{{eventIds}}")) {
			const parsed = tryParseJsonString(body);
			if (!parsed) {
				throw new Error(
					"owaEventRequestTemplate body must be JSON or include {{eventIds}} so EventIds can be injected"
				);
			}
			body = injectEventIdsIntoBody(parsed, payload);
		}
	} else if (typeof body === "object") {
		body = injectEventIdsIntoBody(body, payload);
	}

	if (typeof body === "object" && !hasEventIds(body)) {
		throw new Error("owaEventRequestTemplate body missing EventIds; re-run discover-owa-log");
	}

	return applyTemplate({ template: { ...template, body }, vars: varsWithIds });
}

/**
 * @param {{ base: any[], details: any[] }} opts
 */
function mergeEvents({ base, details }) {
	const byId = new Map();
	const noId = [];

	for (const event of base ?? []) {
		if (event?.sourceId) {
			byId.set(event.sourceId, event);
		} else {
			noId.push(event);
		}
	}

	for (const event of details ?? []) {
		if (!event?.sourceId) {
			noId.push(event);
			continue;
		}
		const existing = byId.get(event.sourceId);
		if (!existing) {
			byId.set(event.sourceId, event);
			continue;
		}
		byId.set(event.sourceId, {
			...existing,
			...event,
			attendeeNames:
				event.attendeeNames && event.attendeeNames.length
					? event.attendeeNames
					: existing.attendeeNames,
			organizerEmail: event.organizerEmail ?? existing.organizerEmail,
			sourceCalendarName: event.sourceCalendarName ?? existing.sourceCalendarName,
			sourceOwnerEmail: event.sourceOwnerEmail ?? existing.sourceOwnerEmail,
		});
	}

	return [...byId.values(), ...noId];
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

	const vars = await resolveTemplateVars({ page, template, range, templateVars });
	const req = applyTemplate({ template, vars });
	const json = await owaFetchJson(page, req);
	return extractOutlookEventsFromJson(json);
}

/**
 * Fetch list events using a view template, then hydrate attendees via an event details template.
 *
 * @param {{
 *   page: any,
 *   viewTemplate: OwaRequestTemplate,
 *   eventTemplate?: OwaRequestTemplate | null,
 *   range: { start: Date, end: Date },
 *   templateVars?: Record<string,string>,
 *   batchSize?: number,
 * }} opts
 */
export async function fetchOwaEventsByTemplates({
	page,
	viewTemplate,
	eventTemplate,
	range,
	templateVars,
	batchSize = 50,
}) {
	if (!viewTemplate || typeof viewTemplate !== "object") {
		throw new Error("Missing owaRequestTemplate");
	}
	if (!viewTemplate.url || !viewTemplate.method) {
		throw new Error("owaRequestTemplate must include url + method");
	}

	const viewVars = await resolveTemplateVars({ page, template: viewTemplate, range, templateVars });
	const viewReq = applyTemplate({ template: viewTemplate, vars: viewVars });
	const viewJson = await owaFetchJson(page, viewReq);
	const viewEvents = extractOutlookEventsFromJson(viewJson);

	if (!eventTemplate) return viewEvents;

	const eventIds = extractOutlookEventIdsFromJson(viewJson);
	if (!eventIds.length) return viewEvents;

	const detailVars = await resolveTemplateVars({ page, template: eventTemplate, range, templateVars });
	const uniqueIds = Array.from(new Set(eventIds));
	const batches = chunk(uniqueIds, Math.max(1, batchSize));

	const detailEvents = [];
	for (const batch of batches) {
		const detailReq = buildEventDetailsRequest({
			template: eventTemplate,
			vars: detailVars,
			eventIds: batch,
		});
		const detailJson = await owaFetchJson(page, detailReq);
		detailEvents.push(...extractOutlookEventsFromJson(detailJson));
	}

	return mergeEvents({ base: viewEvents, details: detailEvents });
}

/**
 * Unified helper for future use.
 *
 * @param {{
 *   mode: "capture" | "template",
 *   page: any,
 *   capture?: { durationMs: number, urlIncludes?: string | null },
 *   template?: {
 *     template: OwaRequestTemplate,
 *     detailTemplate?: OwaRequestTemplate | null,
 *     range: { start: Date, end: Date },
 *     templateVars?: Record<string,string>,
 *   },
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
	if (opts.template.detailTemplate) {
		return await fetchOwaEventsByTemplates({
			page: opts.page,
			viewTemplate: opts.template.template,
			eventTemplate: opts.template.detailTemplate,
			range: opts.template.range,
			templateVars: opts.template.templateVars,
		});
	}

	return await fetchOwaEventsByTemplate({
		page: opts.page,
		template: opts.template.template,
		range: opts.template.range,
		templateVars: opts.template.templateVars,
	});
}
