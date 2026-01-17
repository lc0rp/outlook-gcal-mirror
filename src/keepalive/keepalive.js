import { createWriteStream, mkdirSync, existsSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

import { launchEngine } from "./engines.js";
import { sleep, stripQueryParam, withCacheBuster } from "./utils.js";

const LEGACY_USER_DATA_DIR = join(homedir(), ".browser-keepalive", "chrome");
const DEFAULT_USER_DATA_DIR = join(homedir(), ".config", "outlook-gcal-mirror", "chrome");

export function resolveUserDataDir(explicitDir) {
	if (explicitDir && String(explicitDir).trim()) return String(explicitDir).trim();
	if (existsSync(LEGACY_USER_DATA_DIR)) return LEGACY_USER_DATA_DIR;
	return DEFAULT_USER_DATA_DIR;
}

function ensureDir(dir) {
	if (!dir) return;
	try {
		mkdirSync(dir, { recursive: true });
	} catch (err) {
		throw new Error(`Failed to create user data dir '${dir}': ${err.message || err}`);
	}
}

function shouldRecordBody(contentType) {
	if (!contentType) return true;
	const ct = contentType.toLowerCase();
	return (
		ct.includes("json") ||
		ct.includes("text") ||
		ct.includes("javascript") ||
		ct.includes("xml") ||
		ct.includes("html") ||
		ct.includes("form")
	);
}

function getHeader(headers, name) {
	if (!headers) return "";
	const key = name.toLowerCase();
	return headers[key] ?? headers[name] ?? "";
}

function startNetworkRecorder(page, options) {
	const recordPath = options.recordNetworkPath;
	if (!recordPath) return null;

	ensureDir(dirname(recordPath));

	const includes = options.recordIncludes ?? [];
	const recordBody = options.recordBody !== false;
	const maxBytes = options.recordMaxBytes ?? 1000000;
	const stream = createWriteStream(recordPath, { flags: "a" });

	const shouldIncludeUrl = (url) =>
		!includes.length || includes.some((needle) => needle && url.includes(needle));

	const writeEntry = (entry) => {
		try {
			stream.write(`${JSON.stringify(entry)}\n`);
		} catch {
			// ignore
		}
	};

	const onResponse = async (response) => {
		try {
			const url = typeof response.url === "function" ? response.url() : response.url;
			if (!url) return;
			const urlString = String(url);
			if (!shouldIncludeUrl(urlString)) return;

			const request = typeof response.request === "function" ? response.request() : null;
			const method = request && typeof request.method === "function" ? request.method() : request?.method;
			const status = typeof response.status === "function" ? response.status() : response.status;

			const responseHeaders =
				typeof response.headers === "function" ? response.headers() : response.headers;
			const contentType = String(getHeader(responseHeaders, "content-type") ?? "");

			let body = null;
			let bodyTruncated = false;
			let bodyError = null;
			if (recordBody && shouldRecordBody(contentType)) {
				try {
					let text = await response.text();
					if (typeof text !== "string") text = text ? String(text) : "";
					if (maxBytes && text.length > maxBytes) {
						bodyTruncated = true;
						text = text.slice(0, maxBytes);
					}
					body = text;
				} catch (err) {
					bodyError = err?.message ?? String(err ?? "");
				}
			}

			let requestPostData = null;
			if (request && typeof request.postData === "function") {
				try {
					requestPostData = request.postData();
				} catch {
					requestPostData = null;
				}
			} else if (request?.postData) {
				requestPostData = request.postData;
			}
			if (typeof requestPostData === "string" && maxBytes && requestPostData.length > maxBytes) {
				requestPostData = requestPostData.slice(0, maxBytes);
			}

			writeEntry({
				ts: new Date().toISOString(),
				url: urlString,
				method: method ?? "GET",
				status,
				contentType,
				body,
				bodyTruncated,
				bodyError,
				requestPostData,
			});
		} catch (err) {
			writeEntry({
				ts: new Date().toISOString(),
				error: err?.message ?? String(err ?? ""),
			});
		}
	};

	page.on("response", onResponse);

	return {
		stop() {
			try {
				page.off("response", onResponse);
			} catch {
				// ignore
			}
			try {
				stream.end();
			} catch {
				// ignore
			}
		},
	};
}

function registerActivityTracking(page, markActivity) {
	const events = [
		"domcontentloaded",
		"load",
		"framenavigated",
		"request",
		"requestfinished",
		"requestfailed",
		"response",
	];

	for (const evt of events) {
		try {
			page.on(evt, () => markActivity(evt));
		} catch {
			// best-effort
		}
	}
}

async function waitForIdle({ intervalMs, getLastActivityAt, stoppedRef }) {
	while (!stoppedRef.stopped) {
		const now = Date.now();
		const idleForMs = now - getLastActivityAt();
		if (idleForMs >= intervalMs) {
			return;
		}

		const remainingMs = intervalMs - idleForMs;
		const sleepMs = Math.min(remainingMs, 5000);
		console.info(`[keepalive] waiting for idle (~${Math.ceil(remainingMs / 1000)}s remaining)`);
		await sleep(sleepMs);
	}
}

async function waitForJson(urlString, timeoutMs) {
	const startedAt = Date.now();
	let lastErr;

	while (Date.now() - startedAt < timeoutMs) {
		try {
			if (typeof fetch !== "function") {
				throw new Error("global fetch() is not available (Node 18+ required)");
			}
			const res = await fetch(urlString, { headers: { accept: "application/json" } });
			if (!res.ok) {
				throw new Error(`HTTP ${res.status}`);
			}
			return await res.json();
		} catch (err) {
			lastErr = err;
			await sleep(250);
		}
	}

	throw lastErr ?? new Error(`Timed out fetching ${urlString}`);
}

async function printCdpEndpoints(cdpPort) {
	const base = `http://127.0.0.1:${cdpPort}`;
	console.info(`[keepalive] CDP enabled: ${base}`);

	try {
		const version = await waitForJson(`${base}/json/version`, 10000);
		if (version?.webSocketDebuggerUrl) {
			console.info(`[keepalive] CDP websocket: ${version.webSocketDebuggerUrl}`);
		}
	} catch (err) {
		console.warn("[keepalive] CDP: could not read /json/version:", err.message || err);
	}
}

export async function runKeepalive(config) {
	const baseUrl = config.cacheBust ? stripQueryParam(config.url, "_cb") : config.url;
	const firstUrl = config.cacheBust ? withCacheBuster(baseUrl) : baseUrl;

	ensureDir(config.userDataDir);
	const session = await launchEngine(config.engine, {
		headless: config.headless,
		userDataDir: config.userDataDir,
		cdpPort: config.cdpPort,
	});

	if (config.cdpPort) {
		await printCdpEndpoints(config.cdpPort);
	}

	let stopped = false;
	const stoppedRef = {
		get stopped() {
			return stopped;
		},
	};

	let lastActivityAt = Date.now();
	const markActivity = () => {
		lastActivityAt = Date.now();
	};
	registerActivityTracking(session.page, markActivity);

	const recordIncludes = Array.isArray(config.recordIncludes) ? config.recordIncludes : [];
	const recorder = startNetworkRecorder(session.page, { ...config, recordIncludes });
	if (recorder) {
		const includeLabel = recordIncludes.length ? recordIncludes.join(",") : "all";
		console.info(
			`[keepalive] network log: ${config.recordNetworkPath} (include=${includeLabel} body=${config.recordBody} maxBytes=${config.recordMaxBytes})`
		);
	}

	const stop = async (reason) => {
		if (stopped) return;
		stopped = true;
		console.info(`[keepalive] stopping (${reason})...`);
		if (recorder) recorder.stop();
		await session.close();
		process.exit(0);
	};

	process.on("SIGINT", () => void stop("SIGINT"));
	process.on("SIGTERM", () => void stop("SIGTERM"));

	console.info(
		`[keepalive] engine=${session.engine} interval=${config.intervalSeconds}s cacheBust=${config.cacheBust} alwaysReset=${config.alwaysReset} headless=${config.headless} userDataDir=${config.userDataDir || "(none)"} cdp=${config.cdpPort ?? "off"} onlyIfIdle=${config.onlyIfIdle}`
	);
	console.info(`[keepalive] loading: ${firstUrl}`);

	await session.goto(firstUrl, { waitUntil: "domcontentloaded" });
	markActivity();

	const intervalMs = config.intervalSeconds * 1000;

	while (!stopped) {
		await sleep(intervalMs);
		if (stopped) break;

		if (config.onlyIfIdle) {
			await waitForIdle({
				intervalMs,
				getLastActivityAt: () => lastActivityAt,
				stoppedRef,
			});
			if (stopped) break;
		}

		try {
			if (config.alwaysReset) {
				const nextUrl = config.cacheBust ? withCacheBuster(baseUrl) : baseUrl;
				console.info(`[keepalive] goto: ${nextUrl}`);
				await session.goto(nextUrl, { waitUntil: "domcontentloaded" });
				continue;
			}

			if (config.cacheBust) {
				const current = await session.currentUrl();
				const currentBase = current && current !== "about:blank" ? stripQueryParam(current, "_cb") : baseUrl;
				const nextUrl = withCacheBuster(currentBase);
				console.info(`[keepalive] goto: ${nextUrl}`);
				await session.goto(nextUrl, { waitUntil: "domcontentloaded" });
				continue;
			}

			console.info("[keepalive] reload");
			await session.reload({ waitUntil: "domcontentloaded" });
		} catch (err) {
			console.error("[keepalive] refresh failed:", err);
		}
	}
}
