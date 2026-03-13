/**
 * @param {import("./normalized-event.js").NormalizedEvent} ev
 */
export function buildGoogleDescription(ev) {
	const lines = [];
	lines.push("Mirrored from Outlook (read-only)");

	if (ev.location) {
		lines.push("");
		lines.push(`Location: ${ev.location}`);
	}

	if (ev.attendeeNames.length) {
		lines.push("");
		lines.push("Attendees:");
		for (const name of ev.attendeeNames) {
			lines.push(`- ${name}`);
		}
	}

	if (ev.organizerEmail) {
		lines.push("");
		lines.push(`Organizer: ${ev.organizerEmail}`);
	}

	if (ev.sourceCalendarName || ev.sourceOwnerEmail) {
		lines.push("");
		lines.push(
			`Source calendar: ${ev.sourceCalendarName ?? "?"}${ev.sourceOwnerEmail ? ` (${ev.sourceOwnerEmail})` : ""}`
		);
	}

	lines.push("");
	lines.push(`Source key: ${ev.sourceKey}`);
	if (ev.sourceId) lines.push(`Source id: ${ev.sourceId}`);

	return lines.join("\n");
}

/**
 * @param {string} summary
 */
export function cancelledSummary(summary) {
	if (summary.toLowerCase().startsWith("cancelled:")) return summary;
	return `CANCELLED: ${summary}`;
}
