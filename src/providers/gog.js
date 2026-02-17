import { runJsonCommand } from "../subprocess.js";

export const GOOGLE_LINK_OUTLOOK_ID = "ogm.link.outlookId";
export const GOOGLE_LINK_VERSION = "ogm.link.version";
export const GOOGLE_LINK_VERSION_VALUE = "1";

function asString(value) {
	if (typeof value === "string" && value.trim()) return value.trim();
	return "";
}

function normalizeDateTime(raw) {
	const value = asString(raw);
	if (!value) return null;
	const parsed = new Date(value);
	if (!Number.isNaN(parsed.getTime())) return { dateTime: parsed.toISOString() };
	return { dateTime: value };
}

function normalizeDateBlock(raw) {
	if (!raw || typeof raw !== "object") return null;
	if (raw.date) return { date: String(raw.date) };
	if (raw.dateTime) return normalizeDateTime(raw.dateTime);
	return null;
}

function normalizeGoogleAttendees(rawAttendees) {
	const attendees = [];
	for (const attendee of rawAttendees ?? []) {
		const email = asString(attendee?.email);
		const name = asString(attendee?.displayName);
		if (!email && !name) continue;
		attendees.push({ email, name });
	}
	return attendees;
}

export function normalizeGoogleEvent(rawEvent) {
	const start = normalizeDateBlock(rawEvent?.start);
	const end = normalizeDateBlock(rawEvent?.end);
	const attendees = normalizeGoogleAttendees(rawEvent?.attendees);
	const attendeeNames = attendees.map((att) => att.name || att.email).filter(Boolean);
	const attendeeEmails = attendees.map((att) => att.email).filter(Boolean);

	return {
		id: asString(rawEvent?.id),
		summary: asString(rawEvent?.summary) || "(untitled)",
		start,
		end,
		allDay: !!start?.date,
		location: asString(rawEvent?.location),
		description: asString(rawEvent?.description),
		attendeeNames,
		attendeeEmails,
		status: asString(rawEvent?.status),
		updatedAt: asString(rawEvent?.updated),
		privateProps: { ...(rawEvent?.extendedProperties?.private ?? {}) },
		raw: rawEvent,
	};
}

function appendArg(args, flag, value) {
	const text = asString(value);
	if (!text) return;
	args.push(flag, text);
}

function formatGoogleTime(value, allDay) {
	if (!value || typeof value !== "object") return "";
	if (value.date) return String(value.date);
	if (value.dateTime) {
		const text = String(value.dateTime);
		if (allDay && /^\d{4}-\d{2}-\d{2}/.test(text)) return text.slice(0, 10);
		return text;
	}
	return "";
}

function pushPrivateProps(args, privateProps) {
	for (const [k, v] of Object.entries(privateProps ?? {})) {
		const key = asString(k);
		const value = asString(v);
		if (!key || !value) continue;
		args.push("--private-prop", `${key}=${value}`);
	}
}

/**
 * @param {{
 *  command?: string,
 *  account?: string,
 *  client?: string,
 *  timeoutMs?: number,
 *  env?: Record<string,string>,
 *  runJson?: typeof runJsonCommand,
 * }} options
 */
export function createGogClient(options = {}) {
	const command = options.command ?? "gog";
	const runJson = options.runJson ?? runJsonCommand;
	const timeoutMs = options.timeoutMs ?? 120000;

	const rootArgs = ["--json"];
	if (options.account) rootArgs.push("--account", String(options.account));
	if (options.client) rootArgs.push("--client", String(options.client));

	const execJson = (args, label) =>
		runJson({
			command,
			args: [...rootArgs, ...args],
			env: options.env,
			timeoutMs,
			label,
		});

	return {
		async listCalendars() {
			const result = await execJson(["calendar", "calendars", "--max", "500"], "gog calendar calendars");
			return Array.isArray(result?.calendars) ? result.calendars : [];
		},

		async resolveCalendarId(calendarRef) {
			const ref = asString(calendarRef);
			if (!ref || ref === "primary") return "primary";
			const calendars = await this.listCalendars();
			const byId = calendars.find((cal) => asString(cal?.id) === ref);
			if (byId?.id) return byId.id;
			const byName = calendars.find((cal) => asString(cal?.summary).toLowerCase() === ref.toLowerCase());
			if (byName?.id) return byName.id;
			return ref;
		},

		async listEvents({ calendarId, from, to, max = 500, query, privatePropFilter, sharedPropFilter }) {
			const args = ["calendar", "events", String(calendarId), "--max", String(max)];
			appendArg(args, "--from", from);
			appendArg(args, "--to", to);
			appendArg(args, "--query", query);
			appendArg(args, "--private-prop-filter", privatePropFilter);
			appendArg(args, "--shared-prop-filter", sharedPropFilter);
			const result = await execJson(args, "gog calendar events");
			const events = Array.isArray(result?.events) ? result.events : [];
			return events.map(normalizeGoogleEvent).filter((ev) => ev.id);
		},

		async createEvent({ calendarId, event, sendUpdates = "none" }) {
			const args = ["calendar", "create", String(calendarId)];
			args.push("--summary", String(event.summary ?? "(untitled)"));
			args.push("--from", formatGoogleTime(event.start, !!event.allDay));
			args.push("--to", formatGoogleTime(event.end, !!event.allDay));
			if (event.allDay) args.push("--all-day");
			appendArg(args, "--location", event.location);
			appendArg(args, "--description", event.description);
			if (sendUpdates) args.push("--send-updates", String(sendUpdates));
			pushPrivateProps(args, event.privateProps);

			const result = await execJson(args, "gog calendar create");
			const created = result?.event ?? result;
			return normalizeGoogleEvent(created);
		},

		async updateEvent({ calendarId, eventId, patch }) {
			const args = ["calendar", "update", String(calendarId), String(eventId)];
			if (patch.summary !== undefined) args.push("--summary", String(patch.summary ?? ""));
			if (patch.start) args.push("--from", formatGoogleTime(patch.start, !!patch.allDay));
			if (patch.end) args.push("--to", formatGoogleTime(patch.end, !!patch.allDay));
			if (patch.description !== undefined) args.push("--description", String(patch.description ?? ""));
			if (patch.location !== undefined) args.push("--location", String(patch.location ?? ""));
			if (patch.allDay === true) args.push("--all-day");
			pushPrivateProps(args, patch.privateProps);

			const result = await execJson(args, "gog calendar update");
			const updated = result?.event ?? result;
			return normalizeGoogleEvent(updated);
		},

		async deleteEvent({ calendarId, eventId }) {
			const result = await execJson(["calendar", "delete", String(calendarId), String(eventId)], "gog calendar delete");
			return {
				deleted: !!result?.deleted,
				eventId: asString(result?.eventId) || String(eventId),
			};
		},
	};
}

export const _internal = { formatGoogleTime, pushPrivateProps, normalizeDateBlock };
