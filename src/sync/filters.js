/**
 * @param {import("../owa/extract.js").NormalizedEvent} ev
 * @param {{ includeCalendars?: string[], skipCalendars?: string[], includeOwnerEmails?: string[], skipOwnerEmails?: string[] }} rules
 */
export function shouldSyncEvent(ev, rules) {
	const includeCalendars = rules.includeCalendars ?? [];
	const skipCalendars = rules.skipCalendars ?? [];
	const includeOwnerEmails = (rules.includeOwnerEmails ?? []).map((s) => s.toLowerCase());
	const skipOwnerEmails = (rules.skipOwnerEmails ?? []).map((s) => s.toLowerCase());

	if (skipCalendars.length && ev.sourceCalendarName) {
		for (const c of skipCalendars) {
			if (ev.sourceCalendarName.toLowerCase().includes(c.toLowerCase())) return false;
		}
	}

	if (includeCalendars.length && ev.sourceCalendarName) {
		let ok = false;
		for (const c of includeCalendars) {
			if (ev.sourceCalendarName.toLowerCase().includes(c.toLowerCase())) ok = true;
		}
		if (!ok) return false;
	}

	if (skipOwnerEmails.length && ev.sourceOwnerEmail) {
		if (skipOwnerEmails.includes(ev.sourceOwnerEmail.toLowerCase())) return false;
	}

	if (includeOwnerEmails.length && ev.sourceOwnerEmail) {
		if (!includeOwnerEmails.includes(ev.sourceOwnerEmail.toLowerCase())) return false;
	}

	return true;
}
