#!/usr/bin/env node

import { Command } from "commander";
import process from "node:process";
import { spawn } from "node:child_process";

import { connectOverCdp } from "./cdp/index.js";
import { UserError } from "./errors.js";
import {
	DEFAULT_CONFIG_PATH,
	DEFAULT_TEMPLATES_PATH,
	DEFAULT_TOKEN_PATH,
	loadConfig,
	saveConfig,
} from "./config.js";
import { validateAbsoluteUrl, validateCdpPort, errorMessage } from "./utils.js";
import { normalizePort } from "./keepalive/engines.js";
import { resolveUserDataDir, runKeepalive } from "./keepalive/keepalive.js";
import { getGoogleCalendarClient } from "./google/client.js";
import {
	discoverOwaCandidates,
	discoverOwaCandidatesFromLog,
	suggestTemplate,
} from "./owa/discovery.js";
import { captureOwaEvents } from "./owa/capture.js";
import { fetchOwaEventsByTemplate, fetchOwaEventsByTemplates } from "./owa/events.js";
import { loadTemplateFromFile, saveTemplatesFile } from "./owa/templates.js";
import { shouldSyncEvent } from "./sync/filters.js";
import {
	findCalendarId,
	getGoogleSyncContext,
	listMirrorEvents,
	listMirrorEventsAll,
	markMirroredCancelled,
	upsertMirroredEvent,
	PRIVATE_SOURCE_KEY,
} from "./sync/google.js";

const DEFAULT_TARGET_URLS = [
	"https://outlook.office.com/calendar/view/week",
	"https://outlook.cloud.microsoft/calendar/view/week",
];
const DEFAULT_TARGET_URL = DEFAULT_TARGET_URLS[0];
const TARGET_HOST_SUFFIXES = ["office.com", "cloud.microsoft"];

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

function parsePositiveNumber(value, label) {
	if (value === undefined || value === null || value === "") return null;
	const n = Number(value);
	if (!Number.isFinite(n) || n <= 0) throw new UserError(`${label} must be a positive number`);
	return n;
}

function collectList(value, previous) {
	if (!value) return previous ?? [];
	return [...(previous ?? []), value];
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
 * @param {unknown} value
 * @returns {"capture" | "template"}
 */
function parseSource(value) {
	const v = String(value ?? "").trim().toLowerCase();
	if (v === "capture" || v === "template") return /** @type {any} */ (v);
	throw new UserError("--source must be 'capture' or 'template'");
}

const MIRROR_MARKER = "Mirrored from Outlook (read-only)";

function sleep(ms) {
	return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForCdpReady({ port, timeoutMs }) {
	const startedAt = Date.now();
	const url = `http://127.0.0.1:${port}/json/version`;
	let lastErr;

	while (Date.now() - startedAt < timeoutMs) {
		try {
			const res = await fetch(url, { headers: { accept: "application/json" } });
			if (res.ok) return true;
			lastErr = new Error(`HTTP ${res.status}`);
		} catch (err) {
			lastErr = err;
		}
		await sleep(500);
	}

	throw lastErr ?? new Error(`Timed out waiting for CDP at ${url}`);
}

function normalizeTargetUrlForMatch(url) {
	if (!url) return null;
	try {
		const parsed = new URL(url);
		parsed.search = "";
		parsed.hash = "";
		return parsed.toString();
	} catch {
		return null;
	}
}

function isAllowedTargetHost(host, targetUrl) {
	if (!host) return false;
	if (targetUrl) {
		try {
			if (new URL(targetUrl).host === host) return true;
		} catch {
			// ignore
		}
	}
	return TARGET_HOST_SUFFIXES.some((suffix) => host === suffix || host.endsWith(`.${suffix}`));
}

function shouldWaitForLogin(url, targetUrl) {
	if (!url) return true;
	const lower = String(url).toLowerCase();
	if (
		lower.includes("login") ||
		lower.includes("signin") ||
		lower.includes("microsoftonline") ||
		lower.includes("oauth") ||
		lower.includes("account")
	) {
		return true;
	}
	try {
		const pageHost = new URL(url).host;
		if (!isAllowedTargetHost(pageHost, targetUrl)) return true;
	} catch {
		// ignore
	}
	return false;
}

async function waitForOutlookLogin({ engine, port, targetUrl, timeoutMs }) {
	const startedAt = Date.now();
	let lastUrl = null;
	while (Date.now() - startedAt < timeoutMs) {
		try {
			const conn = await connectOverCdp({ engine, port, targetUrl });
			const pageUrl = typeof conn.page.url === "function" ? conn.page.url() : conn.page.url;
			lastUrl = pageUrl ?? null;
			await disconnectBestEffort(conn.browser);
			if (!shouldWaitForLogin(pageUrl, targetUrl)) {
				return (
					normalizeTargetUrlForMatch(pageUrl) ??
					normalizeTargetUrlForMatch(targetUrl)
				);
			}
		} catch {
			// ignore; we'll retry until timeout
		}
		console.info("[ensure-cdp] waiting for Outlook login...");
		await sleep(2000);
	}

	throw new UserError(`Timed out waiting for Outlook login. Last URL: ${lastUrl ?? "unknown"}`);
}

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
 * @param {import('./config.js').MirrorConfig | null | undefined} cfg
 */
function templateMatchesAction(template, action) {
	if (!template || typeof template !== "object") return false;
	const url = String(template.url ?? "");
	return url.includes(`action=${action}`);
}

async function resolveOwaTemplates(cfg) {
	const templatesPath = cfg?.outlook?.owaTemplatesPath ?? DEFAULT_TEMPLATES_PATH;

	let viewTemplate = cfg?.outlook?.owaRequestTemplate ?? null;
	let eventTemplate = cfg?.outlook?.owaEventRequestTemplate ?? null;

	if (!viewTemplate) {
		viewTemplate = await loadTemplateFromFile(
			templatesPath,
			(template) => templateMatchesAction(template, "GetCalendarView")
		);
		if (!viewTemplate) {
			viewTemplate = await loadTemplateFromFile(templatesPath);
		}
	}

	if (!eventTemplate) {
		eventTemplate = await loadTemplateFromFile(
			templatesPath,
			(template) => templateMatchesAction(template, "GetCalendarEvent")
		);
	}

	return { viewTemplate, eventTemplate };
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

function buildProgram() {
	const program = new Command();
	program
		.name("outlook-gcal-mirror")
		.description("Mirror Outlook Web calendar details into a dedicated Google Calendar")
		.option("--config <path>", "Config path", DEFAULT_CONFIG_PATH)
		.showHelpAfterError(true);

	program
		.command("keepalive")
		.description("Launch a browser, load a URL, and periodically refresh it to keep it alive")
		.option("--target-url <url>", "OWA calendar URL", DEFAULT_TARGET_URL)
		.option("-i, --interval <seconds>", "Refresh interval in seconds", "60")
		.option("--cache-bust", "Add cache-busting query param on each refresh (default: true)")
		.option("--no-cache-bust", "Disable cache-busting query param")
		.option("--always-reset", "Always navigate to the original URL instead of refreshing current page")
		.option("--engine <engine>", "playwright|puppeteer", "playwright")
		.option("--headless", "Run browser without visible window")
		.option(
			"--user-data-dir <dir>",
			"Persist browser profile/cookies in this directory (defaults to ~/.config/outlook-gcal-mirror/chrome, falling back to ~/.browser-keepalive/chrome)"
		)
		.option("-p, --cdp-port <port>", "Enable Chrome DevTools Protocol on this port")
		.option("--only-if-idle", "Only refresh when browser has been idle for the full interval")
		.option("--record-network <path>", "Write NDJSON network log to this path")
		.option(
			"--record-include <substr>",
			"Only record responses whose URL includes this substring (repeatable)",
			collectList,
			[]
		)
		.option("--record-max-bytes <bytes>", "Max response body bytes to store per entry", "1000000")
		.option("--no-record-body", "Do not include response bodies in network log")
		.action(async (opts) => {
			const targetUrl = validateAbsoluteUrl(opts.targetUrl ?? DEFAULT_TARGET_URL);
			const intervalSeconds = parsePositiveNumber(opts.interval, "--interval") ?? 60;
			const engine = parseEngine(opts.engine);
			const cdpPort = normalizePort(opts.cdpPort);
			const userDataDir = resolveUserDataDir(opts.userDataDir);
			const recordMaxBytes = parsePositiveInt(opts.recordMaxBytes, "--record-max-bytes") ?? 1000000;

			const recordNetworkPath = opts.recordNetwork ? String(opts.recordNetwork).trim() : null;
			const recordIncludes = Array.isArray(opts.recordInclude)
				? opts.recordInclude.map((value) => String(value)).filter(Boolean)
				: [];

			await runKeepalive({
				url: targetUrl,
				intervalSeconds,
				cacheBust: opts.cacheBust !== false,
				alwaysReset: !!opts.alwaysReset,
				engine,
				headless: !!opts.headless,
				userDataDir,
				cdpPort,
				onlyIfIdle: !!opts.onlyIfIdle,
				recordNetworkPath,
				recordIncludes,
				recordMaxBytes,
				recordBody: opts.recordBody !== false,
			});
		});

	program
		.command("setup")
		.description("Write a config file (non-interactive)")
		.option("-p, --cdp-port <port>", "CDP port (where keepalive exposes CDP)")
		.option("--engine <engine>", "playwright|puppeteer", "playwright")
		.option("--target-url <url>", "OWA calendar URL", DEFAULT_TARGET_URL)
		.option("--google-credentials <path>", "Google OAuth credentials JSON (Installed app)")
		.option("--google-token <path>", "Google token JSON path", DEFAULT_TOKEN_PATH)
		.option("--calendar <idOrName>", "Destination Google calendar (id or name)")
		.option("--calendar-name <name>", "Deprecated: use --calendar")
		.option("--window-days <n>", "Days ahead to mirror", "14")
		.option("--mark-cancelled", "Mark missing mirrored events as CANCELLED (unsafe unless full window captured)")
		.action(async (opts) => {
			const cfgPath = program.opts().config;

			const cdpPort = validateCdpPort(opts.cdpPort);
			const engine = parseEngine(opts.engine);
			const targetUrl = validateAbsoluteUrl(opts.targetUrl);

			const credentialsPath = normalizePath(opts.googleCredentials);
			const tokenPath = normalizePath(opts.googleToken) ?? DEFAULT_TOKEN_PATH;
			const calendarRef = String(opts.calendar ?? opts.calendarName ?? "Outlook Mirror");
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
		.command("discover-owa")
		.description("Observe OWA network responses and print candidate JSON endpoints")
		.option("-p, --cdp-port <port>", "CDP port", "9222")
		.option("--engine <engine>", "playwright|puppeteer", "playwright")
		.option("--target-url <url>", "If no suitable tab exists, open this URL", DEFAULT_TARGET_URL)
		.option("--duration-ms <ms>", "How long to observe network", "60000")
		.option("--min-score <n>", "Minimum JSON key score", "3")
		.option(
			"--url-includes <substr>",
			"Only consider responses whose URL contains this substring",
			"outlook.office.com"
		)
		.option("--no-url-filter", "Do not filter responses by URL")
		.action(async (opts) => {
			const cdpPort = validateCdpPort(opts.cdpPort);
			const engine = parseEngine(opts.engine);
			const targetUrl = validateAbsoluteUrl(opts.targetUrl);
			const durationMs = parsePositiveInt(opts.durationMs, "--duration-ms") ?? 60000;
			const minScore = parsePositiveInt(opts.minScore, "--min-score") ?? 3;
			const urlIncludes = opts.urlFilter === false ? null : String(opts.urlIncludes ?? "outlook.office.com");

			console.info("Connects via CDP. While this runs: click calendar items / navigate weeks to generate JSON responses.");

			const conn = await connectOverCdp({ engine, port: cdpPort, targetUrl });
			try {
				const pageUrl = typeof conn.page.url === "function" ? conn.page.url() : conn.page.url;
				console.info(`Using tab: ${pageUrl}`);

				const candidates = await discoverOwaCandidates({
					page: conn.page,
					durationMs,
					minScore,
					urlIncludes,
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
		.command("discover-owa-log")
		.description("Scan an NDJSON network log and print candidate JSON endpoints")
		.option("--log <path>", "Path to NDJSON network log")
		.option("--min-score <n>", "Minimum JSON key score", "3")
		.option(
			"--url-includes <substr>",
			"Only consider responses whose URL contains this substring",
			"outlook.office.com"
		)
		.option("--no-url-filter", "Do not filter responses by URL")
		.option("--save-templates", "Save candidate templates to disk")
		.option(
			"--templates-path <path>",
			"Path to templates JSON (default: ~/.config/outlook-gcal-mirror/templates.json)"
		)
		.action(async (opts) => {
			const logPath = normalizePath(opts.log);
			if (!logPath) {
				throw new UserError("--log is required");
			}
			const minScore = parsePositiveInt(opts.minScore, "--min-score") ?? 3;
			const urlIncludes = opts.urlFilter === false ? null : String(opts.urlIncludes ?? "outlook.office.com");

			const candidates = await discoverOwaCandidatesFromLog({
				filePath: logPath,
				minScore,
				urlIncludes,
			});

			if (!candidates.length) {
				console.info("No candidates found. Capture more traffic or relax --min-score/--url-includes.");
				return;
			}

			const enriched = candidates.map((c) => ({
				...c,
				suggestedTemplate: suggestTemplate(c),
			}));

			if (opts.saveTemplates) {
				const templatesPath = normalizePath(opts.templatesPath) ?? DEFAULT_TEMPLATES_PATH;
				await saveTemplatesFile(templatesPath, {
					generatedAt: new Date().toISOString(),
					candidates: enriched,
				});
				console.info(`Saved templates: ${templatesPath}`);
			}

			console.info(`Found ${candidates.length} candidate endpoint(s):`);
			for (const c of enriched) {
				console.info(`\n- ${c.method} ${c.url}`);
				console.info(`  interestingKeys: ${c.interestingKeys.join(", ")}`);
				console.info("  suggestedTemplate:");
				console.info(JSON.stringify(c.suggestedTemplate, null, 2));
			}
		});

	program
		.command("capture-owa")
		.description("Capture OWA JSON responses for a short window and extract events")
		.option("-p, --cdp-port <port>", "CDP port", "9222")
		.option("--engine <engine>", "playwright|puppeteer", "playwright")
		.option("--target-url <url>", "If no suitable tab exists, open this URL", DEFAULT_TARGET_URL)
		.option("--duration-ms <ms>", "Capture duration", "15000")
		.option(
			"--url-includes <substr>",
			"Only consider responses whose URL contains this substring",
			"outlook.office.com"
		)
		.option("--no-url-filter", "Do not filter responses by URL")
		.option("--json", "Print events as JSON")
		.action(async (opts) => {
			const cdpPort = validateCdpPort(opts.cdpPort);
			const engine = parseEngine(opts.engine);
			const targetUrl = validateAbsoluteUrl(opts.targetUrl);
			const durationMs = parsePositiveInt(opts.durationMs, "--duration-ms") ?? 15000;
			const urlIncludes = opts.urlFilter === false ? null : String(opts.urlIncludes ?? "outlook.office.com");

			console.info("Capturing JSON responses. Tip: open week view and click events while capturing.");

			const conn = await connectOverCdp({ engine, port: cdpPort, targetUrl });
			try {
				const pageUrl = typeof conn.page.url === "function" ? conn.page.url() : conn.page.url;
				console.info(`Using tab: ${pageUrl}`);

				const events = await captureOwaEvents({ page: conn.page, durationMs, urlIncludes });
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
		.command("fetch-owa")
		.description("Fetch OWA JSON via a configured request template and extract events")
		.option("-p, --cdp-port <port>", "CDP port (overrides config)")
		.option("--engine <engine>", "playwright|puppeteer (overrides config)")
		.option("--target-url <url>", "OWA calendar URL (overrides config)")
		.option("--window-days <n>", "Days ahead to fetch (overrides config)")
		.option("--lookback-days <n>", "Days back to include in the fetch window", "1")
		.option("--json", "Print extracted events as JSON")
		.action(async (opts) => {
			const cfgPath = program.opts().config;
			const cfg = await loadConfig(cfgPath);

			const cdpPort = validateCdpPort(opts.cdpPort ?? cfg?.outlook?.cdpPort ?? 9222);
			const engine = parseEngine(opts.engine ?? cfg?.outlook?.engine ?? "playwright");
			const targetUrl = validateAbsoluteUrl(opts.targetUrl ?? cfg?.outlook?.targetUrl ?? DEFAULT_TARGET_URL);

			const windowDays =
				parsePositiveInt(opts.windowDays ?? cfg?.sync?.windowDays, "--window-days") ?? 14;
			const lookbackDays = parsePositiveInt(opts.lookbackDays, "--lookback-days") ?? 1;
			const range = buildMirrorWindowRange({ lookbackDays, windowDays });

			const { viewTemplate, eventTemplate } = await resolveOwaTemplates(cfg);
			if (!viewTemplate) {
				throw new UserError(
					"Missing Outlook request template. Run 'discover-owa' or 'discover-owa-log --save-templates' and set outlook.owaRequestTemplate (or outlook.owaTemplatesPath)."
				);
			}

			if (!eventTemplate) {
				console.info("Note: event details template not found; attendee names may be missing.");
			}

			console.info("Fetching events from Outlook Web using request templates.");
			const conn = await connectOverCdp({ engine, port: cdpPort, targetUrl });
			try {
				const pageUrl = typeof conn.page.url === "function" ? conn.page.url() : conn.page.url;
				console.info(`Using tab: ${pageUrl}`);

				const events = await fetchOwaEventsByTemplates({
					page: conn.page,
					viewTemplate,
					eventTemplate,
					range,
					templateVars: cfg?.outlook?.owaTemplateVars,
				});

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
		.description("Read events from OWA and mirror them to Google Calendar")
		.option("-p, --cdp-port <port>", "CDP port (overrides config)")
		.option("--engine <engine>", "playwright|puppeteer (overrides config)")
		.option("--target-url <url>", "OWA calendar URL (overrides config)")
		.option("--source <source>", "Event source: capture|template", "capture")
		.option("--capture-ms <ms>", "How long to capture OWA JSON (capture source only)", "15000")
		.option(
			"--url-includes <substr>",
			"Only consider responses whose URL contains this substring (capture source only)",
			"outlook.office.com"
		)
		.option("--no-url-filter", "Do not filter responses by URL (capture source only)")
		.option("--google-credentials <path>", "Google OAuth credentials JSON (overrides config)")
		.option("--google-token <path>", "Google token JSON path (overrides config)")
		.option("--calendar <idOrName>", "Destination Google calendar (id or name; overrides config)")
		.option("--calendar-name <name>", "Deprecated: use --calendar")
		.option("--window-days <n>", "Days ahead to mirror (overrides config)")
		.option("--lookback-days <n>", "Days back to include when listing mirror events", "1")
		.option("--mark-cancelled", "Mark missing mirrored events as CANCELLED")
		.option("--ensure-cdp", "Start keepalive if CDP is unavailable and wait for login")
		.option("--ensure-cdp-timeout <ms>", "How long to wait for CDP/login when --ensure-cdp is set", "300000")
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

			const cdpPort = validateCdpPort(opts.cdpPort ?? cfg?.outlook?.cdpPort ?? 9222);
			const engine = parseEngine(opts.engine ?? cfg?.outlook?.engine ?? "playwright");
			const targetUrl = validateAbsoluteUrl(opts.targetUrl ?? cfg?.outlook?.targetUrl ?? DEFAULT_TARGET_URL);
			let effectiveTargetUrl = targetUrl;
			const source = parseSource(opts.source);

			const captureMs = parsePositiveInt(opts.captureMs, "--capture-ms") ?? 15000;
			const urlIncludes = opts.urlFilter === false ? null : String(opts.urlIncludes ?? "outlook.office.com");
			const windowDays =
				parsePositiveInt(opts.windowDays ?? cfg?.sync?.windowDays, "--window-days") ?? 14;
			const lookbackDays = parsePositiveInt(opts.lookbackDays, "--lookback-days") ?? 1;
			const range = buildMirrorWindowRange({ lookbackDays, windowDays });

			const credentialsPath = normalizePath(opts.googleCredentials ?? cfg?.google?.credentialsPath) ?? null;
			const tokenPath = normalizePath(opts.googleToken ?? cfg?.google?.tokenPath) ?? DEFAULT_TOKEN_PATH;
			const calendarRef = String(
				opts.calendar ?? opts.calendarName ?? cfg?.google?.calendarId ?? cfg?.google?.calendarName ?? "Outlook Mirror"
			);

			const markCancelled = opts.markCancelled !== undefined ? !!opts.markCancelled : !!cfg?.sync?.markCancelled;
			const ensureCdp = !!opts.ensureCdp;
			const ensureCdpTimeoutMs = parsePositiveInt(opts.ensureCdpTimeout, "--ensure-cdp-timeout") ?? 300000;
			const logEvents = opts.logEvents !== false;

			if (!opts.dryRun && !credentialsPath) {
				throw new UserError(
					"Missing Google credentials. Run 'setup' or pass --google-credentials /path/to/credentials.json"
				);
			}

			let viewTemplate = null;
			let eventTemplate = null;
			if (source === "capture") {
				console.info("Capturing events from Outlook Web.");
				console.info("Tip: while capturing, click events / navigate weeks so OWA loads JSON.");
			} else {
				console.info("Fetching events from Outlook Web using request templates.");
				({ viewTemplate, eventTemplate } = await resolveOwaTemplates(cfg));
				if (!viewTemplate) {
					throw new UserError(
						"Missing Outlook request template. Run 'discover-owa' or 'discover-owa-log --save-templates' and set outlook.owaRequestTemplate (or outlook.owaTemplatesPath)."
					);
				}
				if (!eventTemplate) {
					console.info("Note: event details template not found; attendee names may be missing.");
				}
			}

			let keepaliveProc = null;
			let startedKeepalive = false;

			if (ensureCdp) {
				try {
					await waitForCdpReady({ port: cdpPort, timeoutMs: 1000 });
				} catch {
					console.info("[ensure-cdp] CDP unavailable; starting keepalive...");
					const args = [
						process.argv[1],
						"keepalive",
						"--target-url",
						targetUrl,
						"--engine",
						engine,
						"--cdp-port",
						String(cdpPort),
					];
					keepaliveProc = spawn(process.execPath, args, { stdio: "inherit" });
					startedKeepalive = true;
					await waitForCdpReady({ port: cdpPort, timeoutMs: ensureCdpTimeoutMs });
				}

				const resolvedTargetUrl = await waitForOutlookLogin({
					engine,
					port: cdpPort,
					targetUrl,
					timeoutMs: ensureCdpTimeoutMs,
				});
				if (resolvedTargetUrl) {
					effectiveTargetUrl = resolvedTargetUrl;
				}
			}

			let conn = null;
			try {
				conn = await connectOverCdp({ engine, port: cdpPort, targetUrl: effectiveTargetUrl });
				try {
					const pageUrl = typeof conn.page.url === "function" ? conn.page.url() : conn.page.url;
					console.info(`Using tab: ${pageUrl}`);

					let events;
					if (source === "capture") {
						events = await captureOwaEvents({ page: conn.page, durationMs: captureMs, urlIncludes });
					} else {
						events = await fetchOwaEventsByTemplates({
							page: conn.page,
							viewTemplate: viewTemplate,
							eventTemplate,
							range,
							templateVars: cfg?.outlook?.owaTemplateVars,
						});
					}

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
					if (source === "capture") {
						console.info(
							"Warning: capture mode is partial by nature; only enable this if you're confident the capture covered the full time window."
						);
					}

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
				} finally {
					if (conn) {
						await disconnectBestEffort(conn.browser);
					}
				}
			} finally {
				if (startedKeepalive && keepaliveProc) {
					try {
						keepaliveProc.kill("SIGTERM");
					} catch {
						// ignore
					}
				}
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

await main();
