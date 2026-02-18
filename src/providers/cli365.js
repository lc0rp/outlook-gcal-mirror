import { runCommand, runJsonCommand } from "../subprocess.js";

export const DEFAULT_CLI365_WORKDIR = "/path/to/projects/cli-365";

function asString(value) {
	if (typeof value === "string" && value.trim()) return value.trim();
	return "";
}

function normalizeDateValue(value, allDay = false) {
	const raw = asString(value);
	if (!raw) return null;
	const isDate = allDay || /^\d{4}-\d{2}-\d{2}$/.test(raw);
	if (isDate) return { date: raw.slice(0, 10) };
	const parsed = new Date(raw);
	if (!Number.isNaN(parsed.getTime())) return { dateTime: parsed.toISOString() };
	return { dateTime: raw };
}

function normalizeLocation(value) {
	if (!value) return "";
	if (typeof value === "string") return value.trim();
	if (typeof value === "object") {
		return asString(value.DisplayName ?? value.displayName ?? value.Name ?? value.name);
	}
	return "";
}

function collectAttendees(rawAttendees) {
	const attendees = [];
	for (const attendee of rawAttendees ?? []) {
		const mailbox = attendee?.Mailbox ?? attendee?.mailbox ?? attendee;
		const email = asString(mailbox?.Address ?? mailbox?.address ?? mailbox?.EmailAddress ?? mailbox?.emailAddress);
		const name = asString(mailbox?.Name ?? mailbox?.DisplayName ?? mailbox?.displayName);
		if (!email && !name) continue;
		attendees.push({ email, name });
	}
	return attendees;
}

export function normalizeOutlookEvent(rawEvent) {
	const allDay = !!rawEvent?.IsAllDayEvent;
	const required = collectAttendees(rawEvent?.RequiredAttendees);
	const optional = collectAttendees(rawEvent?.OptionalAttendees);
	const attendees = [...required, ...optional];
	const attendeeNames = attendees
		.map((att) => att.name || att.email)
		.filter(Boolean);
	const attendeeEmails = attendees
		.map((att) => att.email)
		.filter(Boolean);

	return {
		id: asString(rawEvent?.ItemId ?? rawEvent?.id ?? rawEvent?.ID),
		summary: asString(rawEvent?.Subject) || "(untitled)",
		start: normalizeDateValue(rawEvent?.Start, allDay),
		end: normalizeDateValue(rawEvent?.End, allDay),
		allDay,
		location: normalizeLocation(rawEvent?.Location),
		description: asString(rawEvent?.Body?.Value ?? rawEvent?.Body),
		attendeeNames,
		attendeeEmails,
		isCancelled: !!rawEvent?.IsCancelled,
		updatedAt: asString(rawEvent?.LastModifiedTime),
		raw: rawEvent,
	};
}

function formatCliDate(value, allDay) {
	if (!value || typeof value !== "object") return "";
	if ("date" in value && value.date) return String(value.date);
	if ("dateTime" in value && value.dateTime) {
		if (allDay && /^\d{4}-\d{2}-\d{2}/.test(String(value.dateTime))) {
			return String(value.dateTime).slice(0, 10);
		}
		return String(value.dateTime);
	}
	return "";
}

function appendArg(args, flag, value) {
	const text = asString(value);
	if (!text) return;
	args.push(flag, text);
}

function buildBaseArgs(options) {
	const args = ["--json"];
	if (options.configPath) args.push("--config", String(options.configPath));
	if (options.cdpPort !== undefined && options.cdpPort !== null) {
		args.push("--cdp-port", String(options.cdpPort));
	}
	if (options.ensureCdp) {
		args.push("--ensure-cdp");
		if (options.ensureCdpTimeout) {
			args.push("--ensure-cdp-timeout", String(options.ensureCdpTimeout));
		}
	}
	return args;
}

/**
 * @param {{
 *  command?: string,
 *  commandArgs?: string[],
 *  workdir?: string,
 *  configPath?: string,
 *  cdpPort?: number,
 *  ensureCdp?: boolean,
 *  ensureCdpTimeout?: string,
 *  timeoutMs?: number,
 *  env?: Record<string,string>,
 *  runJson?: typeof runJsonCommand,
 *  runText?: typeof runCommand,
 *  recreateOnUpdateFailure?: boolean,
 * }} options
 */
export function createCli365Client(options = {}) {
	const command = options.command ?? "cli-365";
	const commandArgs = options.commandArgs ?? [];
	const workdir =
		options.workdir ??
		(command === "go" || commandArgs.includes("./cmd/cli-365") ? DEFAULT_CLI365_WORKDIR : undefined);
	const runJson = options.runJson ?? runJsonCommand;
	const runText = options.runText ?? runCommand;
	const baseArgs = buildBaseArgs(options);
	const timeoutMs = options.timeoutMs ?? 120000;
	const recreateOnUpdateFailure = options.recreateOnUpdateFailure !== false;

	const execJson = (args, label) =>
		runJson({
			command,
			args: [...commandArgs, ...baseArgs, ...args],
			cwd: workdir,
			env: options.env,
			timeoutMs,
			label,
		});

	const execText = (args) =>
		runText({
			command,
			args: [...commandArgs, ...baseArgs, ...args],
			cwd: workdir,
			env: options.env,
			timeoutMs,
		});

	return {
		async listEvents({ start, end, limit = 200, folder } = {}) {
			const args = ["calendar", "list", "--limit", String(limit)];
			appendArg(args, "--start", start);
			appendArg(args, "--end", end);
			appendArg(args, "--folder", folder);
			const result = await execJson(args, "cli-365 calendar list");
			const events = Array.isArray(result?.Events)
				? result.Events
				: Array.isArray(result?.events)
					? result.events
					: [];
			return events.map(normalizeOutlookEvent).filter((ev) => ev.id);
		},

		async getEvent({ eventId }) {
			const result = await execJson(["calendar", "get", String(eventId)], "cli-365 calendar get");
			return normalizeOutlookEvent(result);
		},

		async createEvent(event) {
			const args = ["calendar", "create", "--subject", String(event.summary ?? "(untitled)")];
			const allDay = !!event.allDay;
			appendArg(args, "--start", formatCliDate(event.start, allDay));
			appendArg(args, "--end", formatCliDate(event.end, allDay));
			if (allDay) args.push("--all-day");
			appendArg(args, "--location", event.location);
			appendArg(args, "--body", event.description);

			const created = await execJson(args, "cli-365 calendar create");
			return normalizeOutlookEvent(created);
		},

		async updateEvent({ eventId, patch }) {
			const args = ["calendar", "update", String(eventId)];
			if (patch.summary !== undefined) args.push("--subject", String(patch.summary ?? ""));
			if (patch.start) args.push("--start", formatCliDate(patch.start, !!patch.allDay));
			if (patch.end) args.push("--end", formatCliDate(patch.end, !!patch.allDay));
			if (patch.allDay === true) args.push("--all-day");
			if (patch.allDay === false) args.push("--timed");
			if (patch.location !== undefined) args.push("--location", String(patch.location ?? ""));
			if (patch.description !== undefined) args.push("--body", String(patch.description ?? ""));

			try {
				await execJson(args, "cli-365 calendar update");
				const refreshed = await this.getEvent({ eventId });
				return { action: "updated", event: refreshed };
			} catch (err) {
				const message = String(err?.message ?? "");
				const updateNotSupported = message.includes("ErrorSendMeetingInvitationsOrCancellationsRequired");
				if (!recreateOnUpdateFailure || !updateNotSupported) throw err;
				if (!patch.summary || !patch.start || !patch.end) throw err;

				await this.deleteEvent({ eventId });
				const created = await this.createEvent({
					summary: patch.summary,
					start: patch.start,
					end: patch.end,
					allDay: !!patch.allDay,
					location: patch.location,
					description: patch.description,
				});
				return { action: "recreated", event: created, previousId: String(eventId) };
			}
		},

		async deleteEvent({ eventId }) {
			await execText(["calendar", "delete", String(eventId)]);
			return { deleted: true, eventId: String(eventId) };
		},
	};
}

export const _internal = {
	normalizeDateValue,
	formatCliDate,
	collectAttendees,
};
