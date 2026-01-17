import { sha1Hex } from "../utils.js";

/**
 * @typedef {object} NormalizedEvent
 * @property {string} sourceKey
 * @property {string | null} sourceId
 * @property {string} subject
 * @property {{ dateTime: string, timeZone?: string } | { date: string }} start
 * @property {{ dateTime: string, timeZone?: string } | { date: string }} end
 * @property {string[]} attendeeNames
 * @property {string | null} organizerEmail
 * @property {string | null} sourceCalendarName
 * @property {string | null} sourceOwnerEmail
 */

/**
 * @param {unknown} value
 * @returns {value is Record<string, any>}
 */
function isObject(value) {
	return !!value && typeof value === "object" && !Array.isArray(value);
}

/**
 * @param {unknown} value
 * @returns {string | null}
 */
function asString(value) {
	if (typeof value === "string" && value.trim()) return value;
	return null;
}

/**
 * @param {Record<string, any>} node
 */
function extractSourceId(node) {
	return (
		asString(node.Id ?? node.id) ??
		asString(node.ItemId?.Id ?? node.ItemId?.id ?? node.itemId?.Id ?? node.itemId?.id) ??
		asString(node.ItemId ?? node.itemId) ??
		null
	);
}

/**
 * @param {unknown} dt
 */
function parseDateTime(dt) {
	if (typeof dt === "string") {
		const d = new Date(dt);
		if (Number.isNaN(d.getTime())) return null;
		return { dateTime: d.toISOString() };
	}

	if (isObject(dt)) {
		const dateTime = asString(dt.DateTime ?? dt.dateTime ?? dt.datetime);
		const timeZone = asString(dt.TimeZone ?? dt.timeZone ?? dt.tz);
		if (dateTime) {
			// Some OWA payloads use local time without offset; keep as-is.
			return timeZone ? { dateTime, timeZone } : { dateTime };
		}
	}

	return null;
}

/**
 * Extract attendee display names from common OWA shapes.
 * @param {unknown} value
 * @returns {string[]}
 */
function extractAttendeeNames(value) {
	/** @type {string[]} */
	const names = [];

	const pushName = (v) => {
		const s = asString(v);
		if (s) names.push(s);
	};

	const visit = (v) => {
		if (!v) return;
		if (Array.isArray(v)) {
			for (const item of v) visit(item);
			return;
		}
		if (typeof v === "string") {
			pushName(v);
			return;
		}
		if (isObject(v)) {
			pushName(v.Name ?? v.name ?? v.DisplayName ?? v.displayName);
			pushName(v.EmailAddress?.Name ?? v.EmailAddress?.name);
			if (v.EmailAddress && typeof v.EmailAddress === "string") pushName(v.EmailAddress);
			if (v.Attendee && isObject(v.Attendee)) {
				pushName(v.Attendee.Name ?? v.Attendee.DisplayName);
			}
			if (v.attendee && isObject(v.attendee)) {
				pushName(v.attendee.name ?? v.attendee.displayName);
			}
			if (v.attendees) visit(v.attendees);
			if (v.Attendees) visit(v.Attendees);
			if (v.RequiredAttendees) visit(v.RequiredAttendees);
			if (v.OptionalAttendees) visit(v.OptionalAttendees);
			return;
		}
	};

	visit(value);

	// Dedupe (preserve order)
	const seen = new Set();
	const out = [];
	for (const n of names.map((x) => x.trim()).filter(Boolean)) {
		if (seen.has(n)) continue;
		seen.add(n);
		out.push(n);
	}
	return out;
}

/**
 * Best-effort extraction of event-like objects from arbitrary JSON.
 * Looks for common keys: Subject + Start/End.
 *
 * @param {unknown} json
 * @returns {NormalizedEvent[]}
 */
export function extractOutlookEventsFromJson(json) {
	/** @type {NormalizedEvent[]} */
	const events = [];

	const visit = (node) => {
		if (!node) return;
		if (Array.isArray(node)) {
			for (const item of node) visit(item);
			return;
		}
		if (!isObject(node)) return;

		const subject = asString(node.Subject ?? node.subject);
		const start = parseDateTime(node.Start ?? node.start);
		const end = parseDateTime(node.End ?? node.end);

		if (subject && start && end) {
			const sourceId = extractSourceId(node);
			const organizerEmail =
				asString(node.Organizer?.EmailAddress?.Address ?? node.organizer?.email ?? node.organizerEmail) ?? null;
			const sourceCalendarName = asString(node.CalendarName ?? node.calendarName) ?? null;
			const sourceOwnerEmail = asString(node.Owner?.EmailAddress?.Address ?? node.ownerEmail) ?? null;

			const attendeeNames = extractAttendeeNames(
				node.Attendees ?? node.attendees ?? node.RequiredAttendees ?? node.OptionalAttendees
			);

			const sourceKey =
				sourceId ??
				sha1Hex({
					subject,
					start,
					end,
					attendeeNames,
				});

			events.push({
				sourceKey,
				sourceId,
				subject,
				start,
				end,
				attendeeNames,
				organizerEmail,
				sourceCalendarName,
				sourceOwnerEmail,
			});
		}

		for (const v of Object.values(node)) visit(v);
	};

	visit(json);

	// De-dupe by sourceKey (preserve order)
	const seen = new Set();
	const out = [];
	for (const e of events) {
		if (seen.has(e.sourceKey)) continue;
		seen.add(e.sourceKey);
		out.push(e);
	}
	return out;
}

/**
 * Extract event ItemId values from calendar view payloads.
 * @param {unknown} json
 */
export function extractOutlookEventIdsFromJson(json) {
	const ids = [];

	const visit = (node) => {
		if (!node) return;
		if (Array.isArray(node)) {
			for (const item of node) visit(item);
			return;
		}
		if (!isObject(node)) return;

		const subject = asString(node.Subject ?? node.subject);
		const start = parseDateTime(node.Start ?? node.start);
		const end = parseDateTime(node.End ?? node.end);

		if (subject && start && end) {
			const sourceId = extractSourceId(node);
			if (sourceId) ids.push(sourceId);
		}

		for (const v of Object.values(node)) visit(v);
	};

	visit(json);

	const unique = new Set();
	const out = [];
	for (const id of ids) {
		if (unique.has(id)) continue;
		unique.add(id);
		out.push(id);
	}
	return out;
}

export const _internal = { parseDateTime, extractAttendeeNames, extractSourceId };
