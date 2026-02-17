import { sha1Hex } from "../utils.js";
import {
	normalizeBidirState,
	createEmptyBidirState,
} from "./bidir-state.js";
import {
	GOOGLE_LINK_OUTLOOK_ID,
	GOOGLE_LINK_VERSION,
	GOOGLE_LINK_VERSION_VALUE,
} from "../providers/gog.js";

const LEGACY_MIRROR_MARKER = "Mirrored from Outlook (read-only)";

function logLine(logger, line) {
	if (!logger) return;
	if (typeof logger === "function") {
		logger(line);
		return;
	}
	if (typeof logger.info === "function") logger.info(line);
}

function canonicalText(value) {
	return String(value ?? "")
		.replace(/\s+/g, " ")
		.trim();
}

function canonicalTime(value) {
	if (!value || typeof value !== "object") return "";
	if (value.date) return `D:${value.date}`;
	const raw = String(value.dateTime ?? "").trim();
	if (!raw) return "";
	const parsed = new Date(raw);
	if (!Number.isNaN(parsed.getTime())) return `T:${parsed.toISOString()}`;
	return `T:${raw}`;
}

function eventIdentityKey(event) {
	return [
		canonicalText(event.summary).toLowerCase(),
		canonicalTime(event.start),
		canonicalTime(event.end),
	]
		.filter(Boolean)
		.join("|");
}

function eventFingerprint(event) {
	const cleaned = {
		summary: canonicalText(event.summary),
		start: canonicalTime(event.start),
		end: canonicalTime(event.end),
		allDay: !!event.allDay,
		location: canonicalText(event.location),
		description: canonicalText(event.description),
	};
	return sha1Hex(cleaned);
}

function dedupeLinks(links) {
	const seenOutlook = new Set();
	const seenGoogle = new Set();
	const out = [];
	for (const link of links) {
		if (!link?.outlookId || !link?.googleId) continue;
		if (seenOutlook.has(link.outlookId) || seenGoogle.has(link.googleId)) continue;
		seenOutlook.add(link.outlookId);
		seenGoogle.add(link.googleId);
		out.push(link);
	}
	return out;
}

function buildDescription(baseDescription, attendeeNames, sourceLabel) {
	const base = canonicalText(baseDescription);
	const names = Array.isArray(attendeeNames)
		? attendeeNames.map((n) => canonicalText(n)).filter(Boolean)
		: [];
	if (!names.length) return base;
	const attendees = names.join(", ");
	if (!base) return `Attendees (${sourceLabel}): ${attendees}`;
	if (base.includes(attendees)) return base;
	return `${base}\n\nAttendees (${sourceLabel}): ${attendees}`;
}

function isLegacyMirroredGoogleEvent(event) {
	const description = canonicalText(event?.description);
	if (!description) return false;
	return description.includes(LEGACY_MIRROR_MARKER);
}

function toGoogleEvent(outlookEvent) {
	return {
		summary: outlookEvent.summary,
		start: outlookEvent.start,
		end: outlookEvent.end,
		allDay: !!outlookEvent.allDay,
		location: outlookEvent.location,
		description: buildDescription(outlookEvent.description, outlookEvent.attendeeNames, "Outlook"),
		privateProps: {
			[GOOGLE_LINK_OUTLOOK_ID]: outlookEvent.id,
			[GOOGLE_LINK_VERSION]: GOOGLE_LINK_VERSION_VALUE,
		},
	};
}

function toOutlookPatch(googleEvent) {
	return {
		summary: googleEvent.summary,
		start: googleEvent.start,
		end: googleEvent.end,
		allDay: !!googleEvent.allDay,
		location: googleEvent.location,
		description: buildDescription(googleEvent.description, googleEvent.attendeeNames, "Google"),
	};
}

function initSummary() {
	return {
		outlookTotal: 0,
		googleTotal: 0,
		linkedByIdentity: 0,
		createdOnGoogle: 0,
		updatedOnGoogle: 0,
		createdOnOutlook: 0,
		updatedOnOutlook: 0,
		skippedLegacyGoogleEvents: 0,
		removedStaleLinks: 0,
		actions: [],
	};
}

function pushAction(summary, action, data) {
	summary.actions.push({ action, ...data });
}

/**
 * @param {{
 *  outlookClient: { listEvents: Function, createEvent: Function, updateEvent: Function },
 *  googleClient: { listEvents: Function, createEvent: Function, updateEvent: Function },
 *  calendarId: string,
 *  range: { start: Date, end: Date },
 *  stateStore: { load: Function, save: Function },
 *  dryRun?: boolean,
 *  logger?: any,
 *  outlookFolder?: string,
 * }} opts
 */
export async function runBidirectionalSync(opts) {
	const dryRun = !!opts.dryRun;
	const summary = initSummary();
	const stateRaw = opts.stateStore?.load ? await opts.stateStore.load() : createEmptyBidirState();
	const state = normalizeBidirState(stateRaw);

	const start = opts.range.start.toISOString();
	const end = opts.range.end.toISOString();

	const outlookEvents = (await opts.outlookClient.listEvents({
		start,
		end,
		limit: 1000,
		folder: opts.outlookFolder,
	}))
		.filter((ev) => ev.id)
		.filter((ev) => !ev.isCancelled);

	const googleEvents = (await opts.googleClient.listEvents({
		calendarId: opts.calendarId,
		from: start,
		to: end,
		max: 1000,
	}))
		.filter((ev) => ev.id)
		.filter((ev) => ev.status !== "cancelled");

	summary.outlookTotal = outlookEvents.length;
	summary.googleTotal = googleEvents.length;

	const outlookById = new Map(outlookEvents.map((ev) => [ev.id, ev]));
	const googleById = new Map(googleEvents.map((ev) => [ev.id, ev]));
	const usedOutlook = new Set();
	const usedGoogle = new Set();

	const links = state.links.map((link) => ({ ...link }));
	const keptLinks = [];

	for (const link of links) {
		let outlookEvent = outlookById.get(link.outlookId) ?? null;
		let googleEvent = googleById.get(link.googleId) ?? null;

		if (!outlookEvent && !googleEvent) {
			summary.removedStaleLinks += 1;
			continue;
		}

		if (outlookEvent && !googleEvent) {
			if (!dryRun) {
				googleEvent = await opts.googleClient.createEvent({
					calendarId: opts.calendarId,
					event: toGoogleEvent(outlookEvent),
				});
			}
			link.googleId = googleEvent?.id ?? `dry-google-${outlookEvent.id}`;
			summary.createdOnGoogle += 1;
			pushAction(summary, "create_google_missing", { outlookId: outlookEvent.id, googleId: link.googleId });
		}

		if (!outlookEvent && googleEvent) {
			if (!dryRun) {
				outlookEvent = await opts.outlookClient.createEvent(toOutlookPatch(googleEvent));
			}
			link.outlookId = outlookEvent?.id ?? `dry-outlook-${googleEvent.id}`;
			summary.createdOnOutlook += 1;
			pushAction(summary, "create_outlook_missing", { outlookId: link.outlookId, googleId: googleEvent.id });
		}

		if (!outlookEvent || !googleEvent) {
			continue;
		}

		const outlookFp = eventFingerprint(outlookEvent);
		const googleFp = eventFingerprint(googleEvent);

		if (!link.lastFingerprint) {
			if (outlookFp !== googleFp) {
				if (!dryRun) {
					googleEvent = await opts.googleClient.updateEvent({
						calendarId: opts.calendarId,
						eventId: googleEvent.id,
						patch: toGoogleEvent(outlookEvent),
					});
				}
				summary.updatedOnGoogle += 1;
				pushAction(summary, "update_google_initial", { outlookId: outlookEvent.id, googleId: googleEvent.id });
				link.lastFingerprint = outlookFp;
			} else {
				link.lastFingerprint = outlookFp;
			}
		} else if (outlookFp === googleFp) {
			link.lastFingerprint = outlookFp;
		} else if (link.lastFingerprint === outlookFp) {
			if (!dryRun) {
				const updated = await opts.outlookClient.updateEvent({
					eventId: outlookEvent.id,
					patch: toOutlookPatch(googleEvent),
				});
				if (updated?.event?.id && updated.event.id !== outlookEvent.id) {
					link.outlookId = updated.event.id;
					outlookEvent = updated.event;
				}
			}
			summary.updatedOnOutlook += 1;
			pushAction(summary, "update_outlook", { outlookId: link.outlookId, googleId: googleEvent.id });
			link.lastFingerprint = googleFp;
		} else {
			if (!dryRun) {
				googleEvent = await opts.googleClient.updateEvent({
					calendarId: opts.calendarId,
					eventId: googleEvent.id,
					patch: toGoogleEvent(outlookEvent),
				});
			}
			summary.updatedOnGoogle += 1;
			pushAction(summary, "update_google", { outlookId: outlookEvent.id, googleId: googleEvent.id });
			link.lastFingerprint = outlookFp;
		}

		link.updatedAt = new Date().toISOString();
		usedOutlook.add(link.outlookId);
		usedGoogle.add(link.googleId);
		keptLinks.push(link);
	}

	const unmatchedOutlook = outlookEvents.filter((ev) => !usedOutlook.has(ev.id));
	const unmatchedGoogle = googleEvents.filter((ev) => !usedGoogle.has(ev.id));
	const googleByIdentity = new Map();
	for (const googleEvent of unmatchedGoogle) {
		if (isLegacyMirroredGoogleEvent(googleEvent)) continue;
		const key = eventIdentityKey(googleEvent);
		if (!key || googleByIdentity.has(key)) continue;
		googleByIdentity.set(key, googleEvent);
	}

	for (const outlookEvent of unmatchedOutlook) {
		const key = eventIdentityKey(outlookEvent);
		if (!key) continue;
		const googleEvent = googleByIdentity.get(key);
		if (!googleEvent || usedGoogle.has(googleEvent.id)) continue;

		const link = {
			outlookId: outlookEvent.id,
			googleId: googleEvent.id,
			lastFingerprint: eventFingerprint(outlookEvent),
			updatedAt: new Date().toISOString(),
		};
		keptLinks.push(link);
		usedOutlook.add(outlookEvent.id);
		usedGoogle.add(googleEvent.id);
		summary.linkedByIdentity += 1;
		pushAction(summary, "link_identity", { outlookId: outlookEvent.id, googleId: googleEvent.id });

		if (eventFingerprint(outlookEvent) !== eventFingerprint(googleEvent)) {
			if (!dryRun) {
				await opts.googleClient.updateEvent({
					calendarId: opts.calendarId,
					eventId: googleEvent.id,
					patch: toGoogleEvent(outlookEvent),
				});
			}
			summary.updatedOnGoogle += 1;
			pushAction(summary, "update_google_identity", { outlookId: outlookEvent.id, googleId: googleEvent.id });
		}
	}

	for (const outlookEvent of unmatchedOutlook) {
		if (usedOutlook.has(outlookEvent.id)) continue;
		let created = null;
		if (!dryRun) {
			created = await opts.googleClient.createEvent({
				calendarId: opts.calendarId,
				event: toGoogleEvent(outlookEvent),
			});
		}
		const googleId = created?.id ?? `dry-google-${outlookEvent.id}`;
		keptLinks.push({
			outlookId: outlookEvent.id,
			googleId,
			lastFingerprint: eventFingerprint(outlookEvent),
			updatedAt: new Date().toISOString(),
		});
		usedOutlook.add(outlookEvent.id);
		usedGoogle.add(googleId);
		summary.createdOnGoogle += 1;
		pushAction(summary, "create_google", { outlookId: outlookEvent.id, googleId });
	}

	for (const googleEvent of unmatchedGoogle) {
		if (usedGoogle.has(googleEvent.id)) continue;
		if (isLegacyMirroredGoogleEvent(googleEvent)) {
			summary.skippedLegacyGoogleEvents += 1;
			pushAction(summary, "skip_legacy_google", { googleId: googleEvent.id });
			continue;
		}

		let created = null;
		if (!dryRun) {
			created = await opts.outlookClient.createEvent(toOutlookPatch(googleEvent));
		}
		const outlookId = created?.id ?? `dry-outlook-${googleEvent.id}`;
		keptLinks.push({
			outlookId,
			googleId: googleEvent.id,
			lastFingerprint: eventFingerprint(googleEvent),
			updatedAt: new Date().toISOString(),
		});
		usedOutlook.add(outlookId);
		usedGoogle.add(googleEvent.id);
		summary.createdOnOutlook += 1;
		pushAction(summary, "create_outlook", { outlookId, googleId: googleEvent.id });
	}

	const nextState = {
		version: state.version,
		links: dedupeLinks(keptLinks),
	};

	if (!dryRun && opts.stateStore?.save) {
		await opts.stateStore.save(nextState);
	}

	logLine(
		opts.logger,
		`bidir sync complete: google created=${summary.createdOnGoogle} updated=${summary.updatedOnGoogle}; outlook created=${summary.createdOnOutlook} updated=${summary.updatedOnOutlook}`
	);

	return { summary, state: nextState };
}

export const _internal = {
	eventIdentityKey,
	eventFingerprint,
	buildDescription,
	isLegacyMirroredGoogleEvent,
	canonicalTime,
};
