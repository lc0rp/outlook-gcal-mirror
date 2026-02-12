import { UserError } from "../errors.js";
import { getGoogleCalendarClient } from "../google/client.js";
import { cancelledSummary, buildGoogleDescription } from "./format.js";

export const PRIVATE_SOURCE_KEY = "ogm.sourceKey";
export const PRIVATE_STATUS = "ogm.status";

const ALWAYS_ATTENDEE = "owner@example.com";

/**
 * @param {unknown} attendees
 */
function normalizeAttendees(attendees) {
	/** @type {Array<{ email: string } & Record<string, any>>} */
	const out = [];
	const seen = new Set();
	if (!Array.isArray(attendees)) return { out, seen };
	for (const att of attendees) {
		if (!att) continue;
		if (typeof att === "string") {
			const email = att.trim();
			if (!email) continue;
			const key = email.toLowerCase();
			if (seen.has(key)) continue;
			seen.add(key);
			out.push({ email });
			continue;
		}
		if (typeof att === "object") {
			const email = String(att.email ?? att.Email ?? "").trim();
			if (!email) continue;
			const key = email.toLowerCase();
			if (seen.has(key)) continue;
			seen.add(key);
			out.push({ ...att, email });
		}
	}
	return { out, seen };
}

/**
 * @param {unknown} attendees
 */
function withAlwaysAttendee(attendees) {
	const { out, seen } = normalizeAttendees(attendees);
	const key = ALWAYS_ATTENDEE.toLowerCase();
	if (!seen.has(key)) out.push({ email: ALWAYS_ATTENDEE });
	return out;
}

function normalizeSummary(summary) {
	if (!summary) return "";
	return String(summary).trim().toLowerCase();
}

function normalizeEventTime(value) {
	if (!value) return "";
	if (typeof value === "string") {
		const trimmed = value.trim();
		const d = new Date(trimmed);
		return Number.isNaN(d.getTime()) ? trimmed : d.toISOString();
	}
	if (typeof value !== "object") return "";

	const date = value.date ?? value.Date;
	if (typeof date === "string" && date.trim()) return date.trim();

	const dateTime = value.dateTime ?? value.DateTime;
	if (typeof dateTime === "string" && dateTime.trim()) {
		const trimmed = dateTime.trim();
		const d = new Date(trimmed);
		if (!Number.isNaN(d.getTime())) return d.toISOString();
		const timeZone = value.timeZone ?? value.TimeZone;
		return timeZone ? `${trimmed}|${timeZone}` : trimmed;
	}

	return "";
}

const CANCEL_PREFIX_RE = /^(canceled|cancelled)\s*:\s*/i;

function stripCancelPrefix(summary) {
	if (!summary) return null;
	const trimmed = String(summary).trim();
	const next = trimmed.replace(CANCEL_PREFIX_RE, "");
	if (!next || next === trimmed) return null;
	return next.trim();
}

function buildEventIdentity({ summary, start, end }) {
	const summaryKey = normalizeSummary(summary);
	const startKey = normalizeEventTime(start);
	const endKey = normalizeEventTime(end);
	if (!summaryKey || !startKey || !endKey) return "";
	return `${summaryKey}|${startKey}|${endKey}`;
}

function buildEventIdentityCandidates({ summary, start, end }) {
	const identities = [];
	const base = buildEventIdentity({ summary, start, end });
	if (base) identities.push(base);
	const stripped = stripCancelPrefix(summary);
	if (stripped) {
		const alt = buildEventIdentity({ summary: stripped, start, end });
		if (alt && !identities.includes(alt)) identities.push(alt);
	}
	return identities;
}

function findEventByIdentities(items, identities) {
	if (!Array.isArray(identities) || identities.length === 0) return null;
	const wanted = new Set(identities.filter(Boolean));
	if (wanted.size === 0) return null;
	for (const item of items ?? []) {
		const itemIdentity = buildEventIdentity({
			summary: item?.summary,
			start: item?.start,
			end: item?.end,
		});
		if (itemIdentity && wanted.has(itemIdentity)) return item;
	}
	return null;
}

/**
 * @param {any} calendar
 * @param {string} calendarRef
 */
export async function findCalendarId({ calendar, calendarRef }) {
	const res = await calendar.calendarList.list();
	const items = res.data.items ?? [];
	const desired = typeof calendarRef === "string" && calendarRef.trim() ? calendarRef.trim() : "Outlook Mirror";
	const foundById = items.find((c) => c.id === desired);
	if (foundById?.id) return foundById.id;
	const foundByName = items.find((c) => (c.summary ?? "").trim() === desired);
	if (foundByName?.id) return foundByName.id;
	return null;
}

export async function ensureMirrorCalendar({ calendar, calendarRef }) {
	const foundId = await findCalendarId({ calendar, calendarRef });
	if (foundId) return foundId;
	const desired = typeof calendarRef === "string" && calendarRef.trim() ? calendarRef.trim() : "Outlook Mirror";

	const created = await calendar.calendars.insert({
		requestBody: { summary: desired, timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone },
	});

	if (!created.data.id) throw new UserError("Failed to create mirror calendar");
	return created.data.id;
}

/**
 * @param {any} calendar
 * @param {{ calendarId: string, timeMin: string, timeMax: string, query?: string }} opts
 */
export async function listMirrorEvents({ calendar, calendarId, timeMin, timeMax, query }) {
	const events = [];
	let pageToken = undefined;

	while (true) {
		const res = await calendar.events.list({
			calendarId,
			timeMin,
			timeMax,
			singleEvents: true,
			maxResults: 2500,
			pageToken,
			...(query ? { q: query } : {}),
		});
		events.push(...(res.data.items ?? []));
		pageToken = res.data.nextPageToken;
		if (!pageToken) break;
	}

	return events;
}

/**
 * @param {any} calendar
 * @param {{ calendarId: string, query?: string }} opts
 */
export async function listMirrorEventsAll({ calendar, calendarId, query }) {
	const events = [];
	let pageToken = undefined;

	while (true) {
		const res = await calendar.events.list({
			calendarId,
			singleEvents: true,
			maxResults: 2500,
			pageToken,
			...(query ? { q: query } : {}),
		});
		events.push(...(res.data.items ?? []));
		pageToken = res.data.nextPageToken;
		if (!pageToken) break;
	}

	return events;
}

/**
 * @param {{ credentialsPath: string, tokenPath: string, calendarRef: string }} opts
 */
export async function getGoogleSyncContext({ credentialsPath, tokenPath, calendarRef }) {
	const { calendar } = await getGoogleCalendarClient({ credentialsPath, tokenPath });
	const calendarId = await ensureMirrorCalendar({ calendar, calendarRef });
	return { calendar, calendarId };
}

/**
 * @param {{ calendar: any, calendarId: string, ev: import("../owa/extract.js").NormalizedEvent }} opts
 */
export async function upsertMirroredEvent({ calendar, calendarId, ev }) {
	// Search existing by private extended prop within a narrow range around the event.
	const timeMin = typeof ev.start.dateTime === "string" ? new Date(ev.start.dateTime).toISOString() : `${ev.start.date}T00:00:00Z`;
	const timeMax = typeof ev.end.dateTime === "string" ? new Date(ev.end.dateTime).toISOString() : `${ev.end.date}T23:59:59Z`;

	const res = await calendar.events.list({
		calendarId,
		timeMin,
		timeMax,
		singleEvents: true,
		maxResults: 50,
		privateExtendedProperty: `${PRIVATE_SOURCE_KEY}=${ev.sourceKey}`,
	});

	let existing = (res.data.items ?? [])[0] ?? null;
	if (!existing) {
		const identities = buildEventIdentityCandidates({
			summary: ev.subject,
			start: ev.start,
			end: ev.end,
		});
		if (identities.length) {
			const fallbackRes = await calendar.events.list({
				calendarId,
				timeMin,
				timeMax,
				singleEvents: true,
				maxResults: 50,
			});
			existing = findEventByIdentities(fallbackRes.data.items ?? [], identities);
		}
	}

	const description = buildGoogleDescription(ev);
	const attendees = withAlwaysAttendee(existing?.attendees ?? []);

	const requestBody = {
		summary: ev.subject,
		description,
		start: ev.start,
		end: ev.end,
		attendees,
		extendedProperties: {
			private: {
				[PRIVATE_SOURCE_KEY]: ev.sourceKey,
				[PRIVATE_STATUS]: "active",
			},
		},
	};

	if (ev.location) {
		requestBody.location = ev.location;
	}

	if (!existing?.id) {
		await calendar.events.insert({
			calendarId,
			sendUpdates: "none",
			requestBody,
		});
		return { action: "created" };
	}

	await calendar.events.patch({
		calendarId,
		eventId: existing.id,
		sendUpdates: "none",
		requestBody,
	});
	return { action: "updated" };
}

/**
 * Mark a mirrored event as cancelled (without deleting it).
 *
 * @param {{ calendar: any, calendarId: string, gcalEvent: any }} opts
 */
export async function markMirroredCancelled({ calendar, calendarId, gcalEvent }) {
	if (!gcalEvent?.id) return { action: "skipped" };
	const summary = gcalEvent.summary ?? "(no title)";
	const nextSummary = cancelledSummary(summary);
	if (nextSummary === summary && gcalEvent.extendedProperties?.private?.[PRIVATE_STATUS] === "cancelled") {
		return { action: "skipped" };
	}

	const requestBody = {
		summary: nextSummary,
		extendedProperties: {
			private: {
				...(gcalEvent.extendedProperties?.private ?? {}),
				[PRIVATE_STATUS]: "cancelled",
			},
		},
		description: (gcalEvent.description ? `${gcalEvent.description}\n\n` : "") + "Marked CANCELLED because the Outlook source event is missing.",
	};

	await calendar.events.patch({
		calendarId,
		eventId: gcalEvent.id,
		sendUpdates: "none",
		requestBody,
	});

	return { action: "cancelled" };
}
