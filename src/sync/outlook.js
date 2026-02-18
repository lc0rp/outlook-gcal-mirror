import { sha1Hex } from "../utils.js";

function asString(value) {
	if (typeof value === "string" && value.trim()) return value.trim();
	return null;
}

function cloneTimeBlock(block, fallbackDateTime) {
	if (block && typeof block === "object") {
		if (typeof block.date === "string" && block.date.trim()) {
			return { date: block.date.trim() };
		}
		if (typeof block.dateTime === "string" && block.dateTime.trim()) {
			return { dateTime: block.dateTime.trim() };
		}
	}
	if (fallbackDateTime) return { dateTime: fallbackDateTime };
	return { dateTime: new Date().toISOString() };
}

/**
 * @param {import("../providers/cli365.js").normalizeOutlookEvent extends (...args: any) => infer R ? R : any} event
 */
export function cli365EventToNormalized(event) {
	const sourceId = asString(event?.id);
	const subject = asString(event?.summary) ?? "(untitled)";
	const organizerEmail =
		asString(event?.raw?.Organizer?.Mailbox?.Address) ??
		asString(event?.raw?.Organizer?.Mailbox?.EmailAddress) ??
		null;
	const sourceCalendarName =
		asString(event?.raw?.ParentFolder?.DisplayName) ??
		asString(event?.raw?.ParentFolder?.displayName) ??
		null;
	const sourceOwnerEmail =
		asString(event?.raw?.Owner?.EmailAddress?.Address) ??
		asString(event?.raw?.Owner?.EmailAddress) ??
		null;

	const start = cloneTimeBlock(event?.start, event?.raw?.Start);
	const end = cloneTimeBlock(event?.end, event?.raw?.End);

	const sourceKey =
		sourceId ??
		sha1Hex({
			subject,
			start,
			end,
			organizerEmail,
		});

	const attendeeNames = Array.isArray(event?.attendeeNames)
		? event.attendeeNames.map((name) => String(name).trim()).filter(Boolean)
		: [];

	return {
		sourceKey,
		sourceId,
		subject,
		start,
		end,
		attendeeNames,
		organizerEmail,
		sourceCalendarName,
		sourceOwnerEmail,
	};
}
