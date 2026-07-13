#!/usr/bin/env node

import { Command } from "commander";
import process from "node:process";
import { pathToFileURL } from "node:url";

import { UserError } from "./errors.js";
import {
	DEFAULT_CONFIG_PATH,
	DEFAULT_BIDIR_STATE_PATH,
	DEFAULT_TOKEN_PATH,
	loadConfig,
	saveConfig,
} from "./config.js";
import { validateCdpPort, errorMessage } from "./utils.js";
import { getGoogleCalendarClient } from "./google/client.js";
import { shouldSyncEvent } from "./sync/filters.js";
import { cli365EventToNormalized } from "./sync/outlook.js";
import {
	findCalendarId,
	getGoogleSyncContext,
	listMirrorEvents,
	listMirrorEventsAll,
	markMirroredCancelled,
	upsertMirroredEvent,
	PRIVATE_SOURCE_KEY,
} from "./sync/google.js";
import { createCli365Client } from "./providers/cli365.js";
import { createGogClient } from "./providers/gog.js";
import { loadBidirState, saveBidirState } from "./sync/bidir-state.js";
import { runBidirectionalSync } from "./sync/bidir.js";

/**
 * @param {unknown} value
 * @param {string} label
 */
function parsePositiveInt(value, label) {
	if (value === undefined || value === null || value === "") return null;
	const n = Number(value);
	if (!Number.isInteger(n) || n <= 0) throw new UserError(`${label} must be a positive integer`);
	return n;
}

function formatEventTimestamp(value) {
	if (!value || typeof value !== "object") return "";
	if ("dateTime" in value && value.dateTime) {
		return value.timeZone ? `${value.dateTime} ${value.timeZone}` : value.dateTime;
	}
	if ("date" in value && value.date) {
		return value.date;
	}
	return "";
}

function formatEventSummary(ev) {
	const start = formatEventTimestamp(ev.start) || "?";
	const end = formatEventTimestamp(ev.end) || "?";
	const subject = ev.subject ? String(ev.subject) : "(untitled)";
	return `${start} → ${end} | ${subject}`;
}

const MIRROR_MARKER = "Mirrored from Outlook (read-only)";

function looksLikeMirroredEvent(event, { requireMarker = true } = {}) {
	const sourceKey = event?.extendedProperties?.private?.[PRIVATE_SOURCE_KEY];
	if (!sourceKey || typeof sourceKey !== "string" || !sourceKey.trim()) return false;
	if (!requireMarker) return true;
	const desc = String(event?.description ?? "");
	if (!desc) return false;
	const normalized = desc.toLowerCase();
	return (
		normalized.includes(MIRROR_MARKER.toLowerCase()) &&
		normalized.includes("source key:")
	);
}

/**
 * @param {string | undefined} maybe
 */
function normalizePath(maybe) {
	if (!maybe) return null;
	return String(maybe);
}

/**
 * @param {{ start: Date, end: Date }} range
 */
function toTimeRangeIso(range) {
	return { timeMin: range.start.toISOString(), timeMax: range.end.toISOString() };
}

function buildMirrorWindowRange({ lookbackDays, windowDays }) {
	const start = new Date();
	start.setDate(start.getDate() - lookbackDays);
	const end = new Date();
	end.setDate(end.getDate() + windowDays);
	return { start, end };
}

export function buildProgram() {
	const program = new Command();
	program
		.name("outlook-gcal-mirror")
		.description("Mirror Outlook calendar details into a dedicated Google Calendar")
		.option("--config <path>", "Config path", process.env.OGM_CONFIG || DEFAULT_CONFIG_PATH)
		.showHelpAfterError(true);

	program
		.command("setup")
		.description("Write a config file (non-interactive)")
		.option("--google-credentials <path>", "Google OAuth credentials JSON (Installed app)")
		.option("--google-token <path>", "Google token JSON path", DEFAULT_TOKEN_PATH)
		.option("--calendar <idOrName>", "Destination Google calendar (id or name)")
		.option("--calendar-name <name>", "Deprecated: use --calendar")
		.option("--window-days <n>", "Days ahead to mirror", "14")
		.option("--mark-cancelled", "Mark missing mirrored events as CANCELLED (unsafe unless full window captured)")
		.action(async (opts) => {
			const cfgPath = program.opts().config;

			const credentialsPath = normalizePath(opts.googleCredentials ?? process.env.OGM_GOOGLE_CREDS);
			const tokenPath = normalizePath(opts.googleToken ?? process.env.OGM_GOOGLE_TOKEN) ?? DEFAULT_TOKEN_PATH;
			const calendarRef = String(opts.calendar ?? opts.calendarName ?? "Outlook Mirror");
			const windowDays = parsePositiveInt(opts.windowDays, "--window-days") ?? 14;

			/** @type {import('./config.js').MirrorConfig} */
			const cfg = {
				google: {
					credentialsPath: credentialsPath ?? "",
					tokenPath,
					calendarName: calendarRef,
				},
				sync: {
					windowDays,
					markCancelled: !!opts.markCancelled,
				},
			};

			await saveConfig(cfgPath, cfg);
			console.info(`Wrote config: ${cfgPath}`);
			if (!credentialsPath) {
				console.info("Note: set google.credentialsPath before running sync (or pass --google-credentials).");
			}
		});

	program
		.command("sync")
		.description("Read events from Outlook via cli-365 and mirror them to Google Calendar")
		.option("--cli365-bin <path>", "cli-365 binary on PATH", "cli-365")
		.option("--cli365-workdir <path>", "Working directory when running cli-365 from source")
		.option("--cli365-config <path>", "cli-365 config path")
		.option("--cli365-cdp-port <port>", "cli-365 CDP port")
		.option("--cli365-folder <id>", "cli-365 calendar folder id")
		.option("--cli365-ensure-cdp", "Pass --ensure-cdp to cli-365")
		.option("--cli365-ensure-cdp-timeout <duration>", "Pass --ensure-cdp-timeout to cli-365")
		.option("--google-credentials <path>", "Google OAuth credentials JSON (overrides config)")
		.option("--google-token <path>", "Google token JSON path (overrides config)")
		.option("--calendar <idOrName>", "Destination Google calendar (id or name; overrides config)")
		.option("--calendar-name <name>", "Deprecated: use --calendar")
		.option("--window-days <n>", "Days ahead to mirror (overrides config)")
		.option("--lookback-days <n>", "Days back to include when listing mirror events", "1")
		.option("--mark-cancelled", "Mark missing mirrored events as CANCELLED")
		.option("--dry-run", "Do not write to Google; print what would happen")
		.option("--no-log-events", "Disable per-event logging")
		.action(async (opts) => {
			const cfgPath = program.opts().config;
			/** @type {import('./config.js').MirrorConfig | null} */
			let cfg = null;
			try {
				cfg = await loadConfig(cfgPath);
			} catch {
				// allow running without config if flags are provided
			}

			const windowDays =
				parsePositiveInt(opts.windowDays ?? cfg?.sync?.windowDays, "--window-days") ?? 14;
			const lookbackDays = parsePositiveInt(opts.lookbackDays, "--lookback-days") ?? 1;
			const range = buildMirrorWindowRange({ lookbackDays, windowDays });

			const credentialsPath = normalizePath(opts.googleCredentials ?? process.env.OGM_GOOGLE_CREDS ?? cfg?.google?.credentialsPath) ?? null;
			const tokenPath = normalizePath(opts.googleToken ?? process.env.OGM_GOOGLE_TOKEN ?? cfg?.google?.tokenPath) ?? DEFAULT_TOKEN_PATH;
			const calendarRef = String(
				opts.calendar ?? opts.calendarName ?? cfg?.google?.calendarId ?? cfg?.google?.calendarName ?? "Outlook Mirror"
			);

			const markCancelled = opts.markCancelled !== undefined ? !!opts.markCancelled : !!cfg?.sync?.markCancelled;
			const logEvents = opts.logEvents !== false;

			if (!opts.dryRun && !credentialsPath) {
				throw new UserError(
					"Missing Google credentials. Run 'setup' or pass --google-credentials /path/to/credentials.json"
				);
			}

			let events = [];

			const cli365Bin = normalizePath(opts.cli365Bin) ?? "cli-365";
			const cli365ConfigPath = normalizePath(opts.cli365Config ?? cfg?.bidir?.cli365ConfigPath);
			const cli365CdpPort =
				opts.cli365CdpPort !== undefined ? validateCdpPort(opts.cli365CdpPort) : undefined;
			const cli365Folder = normalizePath(opts.cli365Folder);
			const cli365EnsureCdp = opts.cli365EnsureCdp === true;
			const cli365EnsureCdpTimeout =
				opts.cli365EnsureCdpTimeout !== undefined
					? normalizePath(opts.cli365EnsureCdpTimeout)
					: null;

			console.info("Fetching events from Outlook via cli-365.");
			const outlookClient = createCli365Client({
				command: cli365Bin,
				workdir: normalizePath(opts.cli365Workdir ?? process.env.OGM_CLI365_WORKDIR) ?? undefined,
				configPath: cli365ConfigPath ?? undefined,
				cdpPort: cli365CdpPort,
				ensureCdp: cli365EnsureCdp,
				ensureCdpTimeout: cli365EnsureCdpTimeout ?? undefined,
			});

			const pulled = await outlookClient.listEvents({
				start: range.start.toISOString(),
				end: range.end.toISOString(),
				limit: 1000,
				folder: cli365Folder ?? undefined,
			});
			events = pulled.map(cli365EventToNormalized);

			if (cfg?.outlook) {
				events = events.filter((ev) =>
					shouldSyncEvent(ev, {
						includeCalendars: cfg.outlook.includeCalendars,
						skipCalendars: cfg.outlook.skipCalendars,
						includeOwnerEmails: cfg.outlook.includeOwnerEmails,
						skipOwnerEmails: cfg.outlook.skipOwnerEmails,
					})
				);
			}

			console.info(`Extracted ${events.length} event(s) after filtering.`);

			if (logEvents) {
				for (const ev of events) {
					console.info(`PULL: ${formatEventSummary(ev)}`);
				}
			}

			if (opts.dryRun) {
				for (const ev of events) {
					console.info(`DRY RUN: would mirror: ${formatEventSummary(ev)} [${ev.sourceKey}]`);
				}
				return;
			}

			const { calendar, calendarId } = await getGoogleSyncContext({
				credentialsPath: /** @type {string} */ (credentialsPath),
				tokenPath,
				calendarRef,
			});

			let created = 0;
			let updated = 0;
			for (const ev of events) {
				const res = await upsertMirroredEvent({ calendar, calendarId, ev });
				if (res.action === "created") created += 1;
				if (res.action === "updated") updated += 1;
				if (logEvents) {
					console.info(`SYNC (${res.action}): ${formatEventSummary(ev)}`);
				}
			}

			console.info(`Upsert complete. created=${created} updated=${updated}`);

			if (!markCancelled) return;

			console.info("mark-cancelled enabled: listing mirror events and cancelling missing source keys.");

			const { timeMin, timeMax } = toTimeRangeIso(range);
			const mirrorEvents = await listMirrorEvents({
				calendar,
				calendarId,
				timeMin,
				timeMax,
				query: "Mirrored from Outlook",
			});

			const present = new Set(events.map((e) => e.sourceKey));
			let cancelled = 0;
			for (const gcalEv of mirrorEvents) {
				const sourceKey = gcalEv.extendedProperties?.private?.[PRIVATE_SOURCE_KEY];
				if (!sourceKey || typeof sourceKey !== "string") continue;
				if (present.has(sourceKey)) continue;
				const res = await markMirroredCancelled({ calendar, calendarId, gcalEvent: gcalEv });
				if (res.action === "cancelled") cancelled += 1;
			}
			console.info(`Cancelled ${cancelled} mirror event(s) not present in scan.`);
		});

	program
		.command("sync-bidir")
		.description("Bi-directional sync between Outlook (cli-365) and Google Calendar (gog)")
		.option("--google-calendar <idOrName>", "Google calendar id/name (default: primary)")
		.option("--gog-account <email>", "Google account email for gog --account")
		.option("--gog-bin <path>", "gog binary path", "gog")
		.option("--state-path <path>", "Path to bidirectional state JSON")
		.option("--window-days <n>", "Days ahead to sync")
		.option("--lookback-days <n>", "Days back to sync", "1")
		.option("--cli365-workdir <path>", "cli-365 working directory")
		.option("--cli365-config <path>", "cli-365 config path")
		.option("--cli365-cdp-port <port>", "cli-365 CDP port")
		.option("--cli365-folder <id>", "cli-365 calendar folder id")
		.option("--cli365-ensure-cdp", "Pass --ensure-cdp to cli-365")
		.option("--cli365-ensure-cdp-timeout <duration>", "Duration string passed to cli-365 --ensure-cdp-timeout")
		.option("--dry-run", "Plan only; do not write events or state")
		.option("--no-log-events", "Disable per-action logs")
		.action(async (opts) => {
			const cfgPath = program.opts().config;
			/** @type {import('./config.js').MirrorConfig | null} */
			let cfg = null;
			try {
				cfg = await loadConfig(cfgPath);
			} catch {
				// allow running without config
			}

			const windowDays =
				parsePositiveInt(opts.windowDays ?? cfg?.sync?.windowDays, "--window-days") ?? 14;
			const lookbackDays = parsePositiveInt(opts.lookbackDays, "--lookback-days") ?? 1;
			const range = buildMirrorWindowRange({ lookbackDays, windowDays });
			const statePath = normalizePath(opts.statePath ?? cfg?.bidir?.statePath) ?? DEFAULT_BIDIR_STATE_PATH;
			const logEvents = opts.logEvents !== false;

			const cli365Workdir = normalizePath(opts.cli365Workdir ?? cfg?.bidir?.cli365Workdir);
			const cli365ConfigPath = normalizePath(opts.cli365Config ?? cfg?.bidir?.cli365ConfigPath);
			const cli365EnsureCdp =
				opts.cli365EnsureCdp !== undefined ? !!opts.cli365EnsureCdp : !!cfg?.bidir?.cli365EnsureCdp;
			const cli365EnsureCdpTimeout = normalizePath(
				opts.cli365EnsureCdpTimeout ?? cfg?.bidir?.cli365EnsureCdpTimeout
			);
			const cli365CdpPortRaw = opts.cli365CdpPort ?? cfg?.bidir?.cli365CdpPort;
			const cli365CdpPort =
				cli365CdpPortRaw !== undefined && cli365CdpPortRaw !== null && cli365CdpPortRaw !== ""
					? validateCdpPort(cli365CdpPortRaw)
					: undefined;
			const cli365Folder = normalizePath(opts.cli365Folder);

			const gogBin = normalizePath(opts.gogBin ?? cfg?.bidir?.gogBin) ?? "gog";
			const gogAccount = normalizePath(opts.gogAccount ?? cfg?.bidir?.gogAccount);
			const calendarRef = String(opts.googleCalendar ?? cfg?.bidir?.calendarId ?? "primary");

			const outlookClient = createCli365Client({
				workdir: cli365Workdir ?? undefined,
				configPath: cli365ConfigPath ?? undefined,
				cdpPort: cli365CdpPort,
				ensureCdp: cli365EnsureCdp,
				ensureCdpTimeout: cli365EnsureCdpTimeout ?? undefined,
			});
			const googleClient = createGogClient({
				command: gogBin,
				account: gogAccount ?? undefined,
			});
			const calendarId = await googleClient.resolveCalendarId(calendarRef);

			const stateStore = {
				load: () => loadBidirState(statePath),
				save: (state) => saveBidirState(statePath, state),
			};

			const { summary } = await runBidirectionalSync({
				outlookClient,
				googleClient,
				calendarId,
				range,
				stateStore,
				outlookFolder: cli365Folder ?? undefined,
				dryRun: !!opts.dryRun,
				logger: logEvents ? console : null,
			});

			console.info(
				`Bi-directional sync complete. google(created=${summary.createdOnGoogle},updated=${summary.updatedOnGoogle}) outlook(created=${summary.createdOnOutlook},updated=${summary.updatedOnOutlook}) linked=${summary.linkedByIdentity}`
			);
			if (summary.skippedLegacyGoogleEvents) {
				console.info(`Skipped ${summary.skippedLegacyGoogleEvents} legacy mirror event(s) on Google.`);
			}
			if (opts.dryRun) {
				console.info("Dry run only: no writes were performed and state file was not updated.");
			}
		});

	program
		.command("clear")
		.description("Delete mirrored events from Google Calendar (dry-run unless --yes)")
		.option("--google-credentials <path>", "Google OAuth credentials JSON (overrides config)")
		.option("--google-token <path>", "Google token JSON path (overrides config)")
		.option("--calendar <idOrName>", "Destination Google calendar (id or name; overrides config)")
		.option("--calendar-name <name>", "Deprecated: use --calendar")
		.option("--no-require-marker", "Skip description marker checks (less safe)")
		.option("--yes", "Actually delete events")
		.option("--no-log-events", "Disable per-event logging")
		.action(async (opts) => {
			const cfgPath = program.opts().config;
			/** @type {import('./config.js').MirrorConfig | null} */
			let cfg = null;
			try {
				cfg = await loadConfig(cfgPath);
			} catch {
				// allow running without config if flags are provided
			}

			const credentialsPath = normalizePath(opts.googleCredentials ?? cfg?.google?.credentialsPath) ?? null;
			const tokenPath = normalizePath(opts.googleToken ?? cfg?.google?.tokenPath) ?? DEFAULT_TOKEN_PATH;
			const calendarRef = String(
				opts.calendar ?? opts.calendarName ?? cfg?.google?.calendarId ?? cfg?.google?.calendarName ?? "Outlook Mirror"
			);
			const requireMarker = opts.requireMarker !== false;
			const logEvents = opts.logEvents !== false;

			if (!credentialsPath) {
				throw new UserError(
					"Missing Google credentials. Run 'setup' or pass --google-credentials /path/to/credentials.json"
				);
			}

			const { calendar } = await getGoogleCalendarClient({
				credentialsPath: /** @type {string} */ (credentialsPath),
				tokenPath,
			});

			const calendarId = await findCalendarId({ calendar, calendarRef });
			if (!calendarId) {
				throw new UserError(`Calendar not found: ${calendarRef}`);
			}

			const mirrorEvents = await listMirrorEventsAll({
				calendar,
				calendarId,
				query: requireMarker ? "Mirrored from Outlook" : undefined,
			});
			const candidates = mirrorEvents.filter((ev) => looksLikeMirroredEvent(ev, { requireMarker }));

			console.info(
				`Found ${mirrorEvents.length} event(s) with source keys; ${candidates.length} matched mirror signatures.`
			);

			if (!candidates.length) return;

			if (!opts.yes) {
				if (logEvents) {
					for (const ev of candidates) {
						console.info(`DRY RUN: would delete: ${ev.summary ?? "(no title)"} (${ev.id ?? "no-id"})`);
					}
				}
				console.info("Dry run only. Re-run with --yes to delete.");
				return;
			}

			let deleted = 0;
			for (const ev of candidates) {
				if (!ev?.id) continue;
				await calendar.events.delete({
					calendarId,
					eventId: ev.id,
					sendUpdates: "none",
				});
				deleted += 1;
				if (logEvents) {
					console.info(`DELETED: ${ev.summary ?? "(no title)"} (${ev.id})`);
				}
			}

			console.info(`Deleted ${deleted} mirrored event(s).`);
		});

	return program;
}

async function main() {
	const program = buildProgram();

	try {
		await program.parseAsync(process.argv);
	} catch (err) {
		const msg = errorMessage(err);
		if (err instanceof UserError) {
			console.error(`Error: ${msg}`);
			process.exitCode = 1;
			return;
		}
		console.error(err);
		console.error(`Error: ${msg}`);
		process.exitCode = 1;
	}
}

const invokedPath = process.argv[1];
if (invokedPath && import.meta.url === pathToFileURL(invokedPath).href) {
	await main();
}
