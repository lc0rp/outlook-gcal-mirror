import { captureOwaEvents } from "./capture.js";
import { applyTemplate, getOwaBearerToken, getOwaCanary, owaFetchJsonWithFallback } from "./fetch.js";
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

const RANGE_START_KEYS = new Set([
	"start",
	"startdate",
	"startdatetime",
	"starttime",
	"startutc",
	"startdateutc",
	"startdatetimeutc",
	"viewstart",
	"rangestart",
	"from",
]);

const RANGE_END_KEYS = new Set([
	"end",
	"enddate",
	"enddatetime",
	"endtime",
	"endutc",
	"enddateutc",
	"enddatetimeutc",
	"viewend",
	"rangeend",
	"to",
]);

function normalizeKey(key) {
	return String(key ?? "").trim().toLowerCase();
}

function applyRangeToValue(value, iso) {
	if (!value || typeof value !== "object") {
		return iso;
	}

	const out = Array.isArray(value) ? [...value] : { ...value };
	if ("DateTime" in out) out.DateTime = iso;
	if ("dateTime" in out) out.dateTime = iso;
	if ("date" in out) out.date = iso.slice(0, 10);
	if ("Date" in out) out.Date = iso.slice(0, 10);
	return out;
}

function applyRangeToNode(node, startIso, endIso, state) {
	if (Array.isArray(node)) {
		return node.map((item) => applyRangeToNode(item, startIso, endIso, state));
	}
	if (!node || typeof node !== "object") return node;

	const out = Array.isArray(node) ? [...node] : { ...node };
	for (const [key, value] of Object.entries(node)) {
		const normalized = normalizeKey(key);
		if (RANGE_START_KEYS.has(normalized)) {
			out[key] = applyRangeToValue(value, startIso);
			state.matched += 1;
			continue;
		}
		if (RANGE_END_KEYS.has(normalized)) {
			out[key] = applyRangeToValue(value, endIso);
			state.matched += 1;
			continue;
		}
		out[key] = applyRangeToNode(value, startIso, endIso, state);
	}
	return out;
}

function parseJsonBody(body) {
	if (typeof body !== "string") return { parsed: false, value: body };
	const parsed = tryParseJsonString(body);
	if (!parsed) return { parsed: false, value: body };
	return { parsed: true, value: parsed };
}

function parseJsonHeader(value) {
	if (typeof value !== "string") return { parsed: false, value };
	let decoded = value;
	let wasEncoded = false;
	try {
		decoded = decodeURIComponent(value);
		wasEncoded = decoded !== value;
	} catch {
		decoded = value;
		wasEncoded = false;
	}

	const parsedDecoded = tryParseJsonString(decoded);
	if (parsedDecoded && typeof parsedDecoded === "object") {
		return { parsed: true, value: parsedDecoded, encoded: wasEncoded };
	}

	const parsedRaw = tryParseJsonString(value);
	if (parsedRaw && typeof parsedRaw === "object") {
		return { parsed: true, value: parsedRaw, encoded: false };
	}

	return { parsed: false, value };
}

export function applyRangeToRequestBody(body, range) {
	if (body === undefined || body === null) {
		return { body, matched: 0, parsed: false };
	}
	const { parsed, value } = parseJsonBody(body);
	if (!value || typeof value !== "object") {
		return { body, matched: 0, parsed };
	}

	const state = { matched: 0 };
	const startIso = range.start.toISOString();
	const endIso = range.end.toISOString();
	const updated = applyRangeToNode(value, startIso, endIso, state);

	if (!state.matched) {
		return { body, matched: 0, parsed };
	}

	return { body: updated, matched: state.matched, parsed };
}

function rangeValueForKey(normalizedKey, iso) {
	if (normalizedKey.includes("date") && !normalizedKey.includes("datetime")) {
		return iso.slice(0, 10);
	}
	return iso;
}

export function applyRangeToRequestUrl(urlString, range) {
	if (!urlString) return { url: urlString, matched: 0 };
	let matched = 0;
	try {
		const url = new URL(urlString);
		for (const key of url.searchParams.keys()) {
			const normalized = normalizeKey(key);
			if (RANGE_START_KEYS.has(normalized)) {
				url.searchParams.set(key, rangeValueForKey(normalized, range.start.toISOString()));
				matched += 1;
				continue;
			}
			if (RANGE_END_KEYS.has(normalized)) {
				url.searchParams.set(key, rangeValueForKey(normalized, range.end.toISOString()));
				matched += 1;
			}
		}
		return { url: url.toString(), matched };
	} catch {
		return { url: urlString, matched: 0 };
	}
}

export function applyRangeToRequestHeaders(headers, range) {
	if (!headers || typeof headers !== "object") return { headers, matched: 0 };
	const startIso = range.start.toISOString();
	const endIso = range.end.toISOString();
	let matched = 0;
	const out = { ...headers };

	for (const [key, value] of Object.entries(headers)) {
		if (typeof value !== "string") continue;
		const parsed = parseJsonHeader(value);
		if (!parsed.parsed || !parsed.value || typeof parsed.value !== "object") continue;

		const state = { matched: 0 };
		const updated = applyRangeToNode(parsed.value, startIso, endIso, state);
		if (!state.matched) continue;

		matched += state.matched;
		const serialized = JSON.stringify(updated);
		out[key] = parsed.encoded ? encodeURIComponent(serialized) : serialized;
	}

	return { headers: out, matched };
}

function templateUsesBearerPrefix(template) {
	try {
		const blob = JSON.stringify(template).toLowerCase();
		return blob.includes("bearer {{owabearer}}".toLowerCase());
	} catch {
		return false;
	}
}

function stripBearerPrefix(value) {
	if (typeof value !== "string") return value;
	return value.replace(/^bearer\s+/i, "");
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

	if (needsBearer && vars.owaBearer && templateUsesBearerPrefix(template)) {
		vars.owaBearer = stripBearerPrefix(vars.owaBearer);
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
	const rangeResult = applyRangeToRequestBody(req.body, range);
	const urlRangeResult = applyRangeToRequestUrl(req.url, range);
	const headerRangeResult = applyRangeToRequestHeaders(req.headers ?? {}, range);
	const requestWithRange = {
		...req,
		url: urlRangeResult.url,
		headers: headerRangeResult.headers,
		body: rangeResult.body,
	};
	const matchedRangeKeys = rangeResult.matched + urlRangeResult.matched + headerRangeResult.matched;
	if (!matchedRangeKeys && !templateContainsPlaceholder(template, "start") && !templateContainsPlaceholder(template, "end")) {
		console.info("Note: request template does not include start/end placeholders or range keys; using captured range if present.");
	}
	const json = await owaFetchJsonWithFallback(page, requestWithRange);
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
	const viewRangeResult = applyRangeToRequestBody(viewReq.body, range);
	const viewUrlRangeResult = applyRangeToRequestUrl(viewReq.url, range);
	const viewHeaderRangeResult = applyRangeToRequestHeaders(viewReq.headers ?? {}, range);
	const viewRequestWithRange = {
		...viewReq,
		url: viewUrlRangeResult.url,
		headers: viewHeaderRangeResult.headers,
		body: viewRangeResult.body,
	};
	const viewMatchedRangeKeys =
		viewRangeResult.matched + viewUrlRangeResult.matched + viewHeaderRangeResult.matched;
	if (!viewMatchedRangeKeys && !templateContainsPlaceholder(viewTemplate, "start") && !templateContainsPlaceholder(viewTemplate, "end")) {
		console.info("Note: view template does not include start/end placeholders or range keys; using captured range if present.");
	}
	const viewJson = await owaFetchJsonWithFallback(page, viewRequestWithRange);
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
		const detailJson = await owaFetchJsonWithFallback(page, detailReq);
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
