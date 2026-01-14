#!/usr/bin/env node

import { Command } from "commander";
import process from "node:process";

import { connectOverCdp } from "./cdp/index.js";
import { UserError } from "./errors.js";
import {
	DEFAULT_CONFIG_PATH,
	DEFAULT_TOKEN_PATH,
	loadConfig,
	saveConfig,
} from "./config.js";
import { validateAbsoluteUrl, validateCdpPort, errorMessage } from "./utils.js";
import { discoverOwaCandidates, suggestTemplate } from "./owa/discovery.js";
import { captureOwaEvents } from "./owa/capture.js";
import { shouldSyncEvent } from "./sync/filters.js";
import {
	getGoogleSyncContext,
	listMirrorEvents,
	markMirroredCancelled,
	upsertMirroredEvent,
	PRIVATE_SOURCE_KEY,
} from "./sync/google.js";

const DEFAULT_TARGET_URL = "https://outlook.office.com/calendar/view/week";

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

/**
 * @param {unknown} value
 * @returns {"playwright" | "puppeteer"}
 */
function parseEngine(value) {
	const v = String(value ?? "").trim().toLowerCase();
	if (v === "playwright" || v === "puppeteer") return /** @type {any} */ (v);
	throw new UserError("--engine must be 'playwright' or 'puppeteer'");
}

/**
 * @param {any} browser
 */
async function disconnectBestEffort(browser) {
	try {
		if (browser && typeof browser.disconnect === "function") {
			browser.disconnect();
			return;
		}
		if (browser && typeof browser.close === "function") {
			await browser.close();
		}
	} catch {
		// ignore
	}
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

function buildProgram() {
	const program = new Command();
	program
		.name("outlook-gcal-mirror")
		.description("Mirror Outlook Web calendar details into a dedicated Google Calendar")
		.option("--config <path>", "Config path", DEFAULT_CONFIG_PATH)
		.showHelpAfterError(true);

	program
		.command("setup")
		.description("Write a config file (non-interactive)")
		.option("-p, --cdp-port <port>", "CDP port (where browser-keepalive exposed CDP)")
		.option("--engine <engine>", "playwright|puppeteer", "playwright")
		.option("--target-url <url>", "OWA calendar URL", DEFAULT_TARGET_URL)
		.option("--google-credentials <path>", "Google OAuth credentials JSON (Installed app)")
		.option("--google-token <path>", "Google token JSON path", DEFAULT_TOKEN_PATH)
		.option("--calendar-name <name>", "Destination Google calendar", "Outlook Mirror")
		.option("--window-days <n>", "Days ahead to mirror", "14")
		.option("--mark-cancelled", "Mark missing mirrored events as CANCELLED (unsafe unless full window captured)")
		.action(async (opts) => {
			const cfgPath = program.opts().config;

			const cdpPort = validateCdpPort(opts.cdpPort);
			const engine = parseEngine(opts.engine);
			const targetUrl = validateAbsoluteUrl(opts.targetUrl);

			const credentialsPath = normalizePath(opts.googleCredentials);
			const tokenPath = normalizePath(opts.googleToken) ?? DEFAULT_TOKEN_PATH;
			const calendarName = String(opts.calendarName ?? "Outlook Mirror");
			const windowDays = parsePositiveInt(opts.windowDays, "--window-days") ?? 14;

			/** @type {import('./config.js').MirrorConfig} */
			const cfg = {
				outlook: {
					cdpPort,
					engine,
					targetUrl,
				},
				google: {
					credentialsPath: credentialsPath ?? "",
					tokenPath,
					calendarName,
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
		.command("discover-owa")
		.description("Observe OWA network responses and print candidate JSON endpoints")
		.option("-p, --cdp-port <port>", "CDP port", "9222")
		.option("--engine <engine>", "playwright|puppeteer", "playwright")
		.option("--target-url <url>", "If no suitable tab exists, open this URL", DEFAULT_TARGET_URL)
		.option("--duration-ms <ms>", "How long to observe network", "60000")
		.option("--min-score <n>", "Minimum JSON key score", "3")
		.action(async (opts) => {
			const cdpPort = validateCdpPort(opts.cdpPort);
			const engine = parseEngine(opts.engine);
			const targetUrl = validateAbsoluteUrl(opts.targetUrl);
			const durationMs = parsePositiveInt(opts.durationMs, "--duration-ms") ?? 60000;
			const minScore = parsePositiveInt(opts.minScore, "--min-score") ?? 3;

			console.info("Connects via CDP. While this runs: click calendar items / navigate weeks to generate JSON responses.");

			const conn = await connectOverCdp({ engine, port: cdpPort, targetUrl });
			try {
				const candidates = await discoverOwaCandidates({
					page: conn.page,
					durationMs,
					minScore,
				});

				if (!candidates.length) {
					console.info("No candidates found. Try increasing --duration-ms and interact with the calendar UI.");
					return;
				}

				console.info(`Found ${candidates.length} candidate endpoint(s):`);
				for (const c of candidates) {
					console.info(`\n- ${c.method} ${c.url}`);
					console.info(`  interestingKeys: ${c.interestingKeys.join(", ")}`);
					console.info("  suggestedTemplate:");
					console.info(JSON.stringify(suggestTemplate(c), null, 2));
				}
			} finally {
				await disconnectBestEffort(conn.browser);
			}
		});

	program
		.command("capture-owa")
		.description("Capture OWA JSON responses for a short window and extract events")
		.option("-p, --cdp-port <port>", "CDP port", "9222")
		.option("--engine <engine>", "playwright|puppeteer", "playwright")
		.option("--target-url <url>", "If no suitable tab exists, open this URL", DEFAULT_TARGET_URL)
		.option("--duration-ms <ms>", "Capture duration", "15000")
		.option("--json", "Print events as JSON")
		.action(async (opts) => {
			const cdpPort = validateCdpPort(opts.cdpPort);
			const engine = parseEngine(opts.engine);
			const targetUrl = validateAbsoluteUrl(opts.targetUrl);
			const durationMs = parsePositiveInt(opts.durationMs, "--duration-ms") ?? 15000;

			console.info("Capturing JSON responses. Tip: open week view and click events while capturing.");

			const conn = await connectOverCdp({ engine, port: cdpPort, targetUrl });
			try {
				const events = await captureOwaEvents({ page: conn.page, durationMs });
				if (opts.json) {
					console.info(JSON.stringify(events, null, 2));
					return;
				}

				console.info(`Extracted ${events.length} event(s).`);
				for (const ev of events.slice(0, 15)) {
					console.info(`- ${ev.subject} (${ev.start.dateTime ?? ev.start.date} → ${ev.end.dateTime ?? ev.end.date})`);
				}
				if (events.length > 15) console.info(`…and ${events.length - 15} more`);
			} finally {
				await disconnectBestEffort(conn.browser);
			}
		});

	program
		.command("sync")
		.description("Capture events from OWA and mirror them to Google Calendar")
		.option("-p, --cdp-port <port>", "CDP port (overrides config)")
		.option("--engine <engine>", "playwright|puppeteer (overrides config)")
		.option("--target-url <url>", "OWA calendar URL (overrides config)")
		.option("--capture-ms <ms>", "How long to capture OWA JSON", "15000")
		.option("--google-credentials <path>", "Google OAuth credentials JSON (overrides config)")
		.option("--google-token <path>", "Google token JSON path (overrides config)")
		.option("--calendar-name <name>", "Destination Google calendar (overrides config)")
		.option("--window-days <n>", "Days ahead to mirror (overrides config)")
		.option("--lookback-days <n>", "Days back to include when listing mirror events", "1")
		.option("--mark-cancelled", "Mark missing mirrored events as CANCELLED (unsafe unless full window captured)")
		.option("--dry-run", "Do not write to Google; print what would happen")
		.action(async (opts) => {
			const cfgPath = program.opts().config;
			/** @type {import('./config.js').MirrorConfig | null} */
			let cfg = null;
			try {
				cfg = await loadConfig(cfgPath);
			} catch {
				// allow running without config if flags are provided
			}

			const cdpPort = validateCdpPort(opts.cdpPort ?? cfg?.outlook?.cdpPort ?? 9222);
			const engine = parseEngine(opts.engine ?? cfg?.outlook?.engine ?? "playwright");
			const targetUrl = validateAbsoluteUrl(opts.targetUrl ?? cfg?.outlook?.targetUrl ?? DEFAULT_TARGET_URL);

			const captureMs = parsePositiveInt(opts.captureMs, "--capture-ms") ?? 15000;
			const windowDays =
				parsePositiveInt(opts.windowDays ?? cfg?.sync?.windowDays, "--window-days") ?? 14;
			const lookbackDays = parsePositiveInt(opts.lookbackDays, "--lookback-days") ?? 1;

			const credentialsPath =
				normalizePath(opts.googleCredentials ?? cfg?.google?.credentialsPath) ?? null;
			const tokenPath =
				normalizePath(opts.googleToken ?? cfg?.google?.tokenPath) ?? DEFAULT_TOKEN_PATH;
			const calendarName = String(opts.calendarName ?? cfg?.google?.calendarName ?? "Outlook Mirror");

			const markCancelled =
				opts.markCancelled !== undefined ? !!opts.markCancelled : !!cfg?.sync?.markCancelled;

			if (!opts.dryRun && !credentialsPath) {
				throw new UserError(
					"Missing Google credentials. Run 'setup' or pass --google-credentials /path/to/credentials.json"
				);
			}

			console.info("Capturing events from Outlook Web.");
			console.info("Tip: while capturing, click events / navigate weeks so OWA loads JSON.");

			const conn = await connectOverCdp({ engine, port: cdpPort, targetUrl });
			try {
				let events = await captureOwaEvents({ page: conn.page, durationMs: captureMs });
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

				if (opts.dryRun) {
					for (const ev of events.slice(0, 30)) {
						console.info(`DRY RUN: would mirror: ${ev.subject} [${ev.sourceKey}]`);
					}
					if (events.length > 30) console.info(`…and ${events.length - 30} more`);
					return;
				}

				const { calendar, calendarId } = await getGoogleSyncContext({
					credentialsPath: /** @type {string} */ (credentialsPath),
					tokenPath,
					calendarName,
				});

				let created = 0;
				let updated = 0;
				for (const ev of events) {
					const res = await upsertMirroredEvent({ calendar, calendarId, ev });
					if (res.action === "created") created += 1;
					if (res.action === "updated") updated += 1;
				}

				console.info(`Upsert complete. created=${created} updated=${updated}`);

				if (!markCancelled) return;

				console.info("mark-cancelled enabled: listing mirror events and cancelling missing source keys.");
				console.info(
					"Warning: only enable this if you're confident the capture covered the full time window."
				);

				const start = new Date();
				start.setDate(start.getDate() - lookbackDays);
				const end = new Date();
				end.setDate(end.getDate() + windowDays);
				const { timeMin, timeMax } = toTimeRangeIso({ start, end });

				const mirrorEvents = await listMirrorEvents({
					calendar,
					calendarId,
					timeMin,
					timeMax,
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
				console.info(`Cancelled ${cancelled} mirror event(s) not present in capture.`);
			} finally {
				await disconnectBestEffort(conn.browser);
			}
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

await main();
