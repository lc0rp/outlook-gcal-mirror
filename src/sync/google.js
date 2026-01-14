import { UserError } from "../errors.js";
import { getGoogleCalendarClient } from "../google/client.js";
import { cancelledSummary, buildGoogleDescription } from "./format.js";

export const PRIVATE_SOURCE_KEY = "ogm.sourceKey";
export const PRIVATE_STATUS = "ogm.status";

/**
 * @param {any} calendar
 * @param {string} calendarName
 */
export async function ensureMirrorCalendar({ calendar, calendarName }) {
	const res = await calendar.calendarList.list();
	const items = res.data.items ?? [];
	const found = items.find((c) => (c.summary ?? "").trim() === calendarName);
	if (found?.id) return found.id;

	const created = await calendar.calendars.insert({
		requestBody: { summary: calendarName, timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone },
	});

	if (!created.data.id) throw new UserError("Failed to create mirror calendar");
	return created.data.id;
}

/**
 * @param {any} calendar
 * @param {{ calendarId: string, timeMin: string, timeMax: string }} opts
 */
export async function listMirrorEvents({ calendar, calendarId, timeMin, timeMax }) {
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
			privateExtendedProperty: `${PRIVATE_SOURCE_KEY}*`,
		});
		events.push(...(res.data.items ?? []));
		pageToken = res.data.nextPageToken;
		if (!pageToken) break;
	}

	return events;
}

/**
 * @param {{ credentialsPath: string, tokenPath: string, calendarName: string }} opts
 */
export async function getGoogleSyncContext({ credentialsPath, tokenPath, calendarName }) {
	const { calendar } = await getGoogleCalendarClient({ credentialsPath, tokenPath });
	const calendarId = await ensureMirrorCalendar({ calendar, calendarName });
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

	const existing = (res.data.items ?? [])[0] ?? null;
	const description = buildGoogleDescription(ev);

	const requestBody = {
		summary: ev.subject,
		description,
		start: ev.start,
		end: ev.end,
		extendedProperties: {
			private: {
				[PRIVATE_SOURCE_KEY]: ev.sourceKey,
				[PRIVATE_STATUS]: "active",
			},
		},
	};

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
