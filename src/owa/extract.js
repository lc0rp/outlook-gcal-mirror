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
 * @property {string | null} location
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

function buildAddressString(address) {
	if (!isObject(address)) return null;
	const street = asString(address.Street ?? address.street ?? address.StreetAddress ?? address.streetAddress ?? address.Address ?? address.address);
	const city = asString(address.City ?? address.city ?? address.Town ?? address.town);
	const state = asString(address.State ?? address.state ?? address.Region ?? address.region);
	const postal = asString(address.PostalCode ?? address.postalCode ?? address.Zip ?? address.zip);
	const country = asString(address.CountryOrRegion ?? address.countryOrRegion ?? address.Country ?? address.country);

	const cityState = [city, state].filter(Boolean).join(", ");
	const cityStatePostal = [cityState, postal].filter(Boolean).join(" ");
	const parts = [street, cityStatePostal, country].filter(Boolean);
	if (!parts.length) return null;
	return parts.join(", ");
}

function extractLocationCandidate(value) {
	if (!value) return null;
	if (typeof value === "string") return asString(value);
	if (!isObject(value)) return null;

	const address = buildAddressString(value.Address ?? value.address);
	if (address) return address;

	return (
		asString(value.DisplayName ?? value.displayName) ??
		asString(value.Name ?? value.name) ??
		asString(value.Location ?? value.location) ??
		asString(value.LocationName ?? value.locationName) ??
		null
	);
}

function extractLocation(node) {
	if (!isObject(node)) return null;

	const candidates = [];
	const push = (value) => {
		const next = extractLocationCandidate(value);
		if (next) candidates.push(next);
	};

	push(node.Location ?? node.location);
	push(node.EnhancedLocation ?? node.enhancedLocation);

	const locations = node.Locations ?? node.locations ?? node.EnhancedLocations ?? node.enhancedLocations;
	if (Array.isArray(locations)) {
		for (const loc of locations) {
			push(loc);
			if (candidates.length) break;
		}
	}

	push(node.LocationDisplayName ?? node.locationDisplayName);
	push(node.LocationName ?? node.locationName);

	if (!candidates.length) return null;
	const seen = new Set();
	for (const item of candidates) {
		const key = item.trim();
		if (!key || seen.has(key)) continue;
		seen.add(key);
		return item;
	}
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

			const mailbox = v.Mailbox;
			if (mailbox && isObject(mailbox)) {
				const mailboxName =
					mailbox.Name ??
					mailbox.DisplayName ??
					mailbox.name ??
					mailbox.displayName ??
					null;
				if (mailboxName) {
					pushName(mailboxName);
				} else if (mailbox.EmailAddress) {
					if (isObject(mailbox.EmailAddress)) {
						pushName(
							mailbox.EmailAddress.Name ??
							mailbox.EmailAddress.Address ??
							mailbox.EmailAddress.address ??
							mailbox.EmailAddress.name
						);
					} else {
						pushName(mailbox.EmailAddress);
					}
				} else {
					pushName(
						mailbox.Address ?? mailbox.address ?? mailbox.emailAddress ?? mailbox.Email
					);
				}
			}
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
			const location = extractLocation(node);

			const attendeeNames = extractAttendeeNames(
				node.Attendees ?? node.attendees ?? node.RequiredAttendees ?? node.OptionalAttendees
			);

			const sourceKey =
				sourceId ??
				sha1Hex({
					subject,
					start,
					end,
					organizerEmail,
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
				location,
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
