#!/usr/bin/env node

import { spawn } from "node:child_process";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

import { DEFAULT_CONFIG_PATH, loadConfig } from "../src/config.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "..");

function envString(name, fallback) {
	const raw = process.env[name];
	if (raw && raw.trim()) return raw.trim();
	return fallback;
}

function envNumber(name, fallback) {
	const raw = envString(name, "");
	if (!raw) return fallback;
	const value = Number(raw);
	if (!Number.isFinite(value)) {
		throw new Error(`${name} must be a number (got: ${raw})`);
	}
	return value;
}

function envBool(name, fallback = false) {
	const raw = envString(name, "");
	if (!raw) return fallback;
	return ["1", "true", "yes", "on"].includes(raw.toLowerCase());
}

function splitList(raw) {
	if (!raw) return [];
	return String(raw)
		.split(",")
		.map((value) => value.trim())
		.filter(Boolean);
}

function delay(ms) {
	return new Promise((resolve) => setTimeout(resolve, ms));
}

function optionalArgs(flag, value, args) {
	if (value === undefined || value === null || value === "") return;
	args.push(flag, String(value));
}

async function run(cmd, args, options = {}) {
	await new Promise((resolve, reject) => {
		const child = spawn(cmd, args, { stdio: "inherit", ...options });
		child.on("error", reject);
		child.on("exit", (code) => {
			if (code === 0) {
				resolve();
				return;
			}
			reject(new Error(`${cmd} exited with code ${code}`));
		});
	});
}

async function runCli(args) {
	await run(process.execPath, [path.join(repoRoot, "src", "cli.js"), ...args], {
		cwd: repoRoot,
	});
}

const configPath = envString("OGM_CONFIG", DEFAULT_CONFIG_PATH);
let cfg = null;
try {
	cfg = await loadConfig(configPath);
} catch {
	cfg = null;
}

const engine = envString("OGM_ENGINE", cfg?.outlook?.engine ?? "playwright");
const cdpPort = envNumber("OGM_CDP_PORT", cfg?.outlook?.cdpPort ?? 9222);
const targetUrl = envString(
	"OGM_TARGET_URL",
	cfg?.outlook?.targetUrl ?? "https://outlook.office.com/calendar/view/week"
);

const discoverMs = envNumber("OGM_DISCOVER_MS", 60000);
const verifyCaptureMs = envNumber("OGM_VERIFY_CAPTURE_MS", 15000);
const syncCaptureMs = envNumber("OGM_SYNC_CAPTURE_MS", 30000);
const idleMinutes = envNumber("BK_IDLE_MINUTES", 150);
const waitMs = envNumber("BK_WAIT_MS", 2000);
const skipBk = envBool("BK_SKIP", false);
const onlyIfIdle = envBool("BK_ONLY_IF_IDLE", true);
const hasTemplate = Boolean(cfg?.outlook?.owaRequestTemplate);

const requestedSource = envString("OGM_SOURCE", "").toLowerCase();
const source =
	requestedSource === "capture" || requestedSource === "template"
		? requestedSource
		: hasTemplate
			? "template"
			: "capture";

if (source === "template" && !hasTemplate) {
	console.error(
		"Missing outlook.owaRequestTemplate in config. Run discover-owa and set it, or run with OGM_SOURCE=capture."
	);
	process.exit(1);
}

const configArgs = configPath ? ["--config", configPath] : [];
const commonArgs = ["--cdp-port", String(cdpPort), "--engine", engine, "--target-url", targetUrl];
const windowArgs = [];
optionalArgs("--window-days", envString("OGM_WINDOW_DAYS", ""), windowArgs);
optionalArgs("--lookback-days", envString("OGM_LOOKBACK_DAYS", ""), windowArgs);

const captureFilterArgs = [];
if (envBool("OGM_NO_URL_FILTER", false)) {
	captureFilterArgs.push("--no-url-filter");
} else {
	optionalArgs("--url-includes", envString("OGM_URL_INCLUDES", ""), captureFilterArgs);
}

const syncExtraArgs = [];
optionalArgs("--google-credentials", envString("OGM_GOOGLE_CREDS", ""), syncExtraArgs);
optionalArgs("--google-token", envString("OGM_GOOGLE_TOKEN", ""), syncExtraArgs);
optionalArgs("--calendar-name", envString("OGM_CALENDAR_NAME", ""), syncExtraArgs);
if (envBool("OGM_MARK_CANCELLED", false)) syncExtraArgs.push("--mark-cancelled");
if (envBool("OGM_DRY_RUN", false)) syncExtraArgs.push("--dry-run");

let bkProc = null;
let exiting = false;
const cleanup = () => {
	if (exiting) return;
	exiting = true;
	if (bkProc && !bkProc.killed) {
		try {
			bkProc.kill("SIGTERM");
		} catch {
			// ignore
		}
	}
};

process.on("SIGINT", () => {
	cleanup();
	process.exit(130);
});
process.on("SIGTERM", () => {
	cleanup();
	process.exit(143);
});
process.on("exit", cleanup);

try {
	if (!skipBk) {
		const bkArgs = [
			path.join(repoRoot, "src", "cli.js"),
			"keepalive",
			"--target-url",
			targetUrl,
			"--engine",
			engine,
		];
		if (onlyIfIdle) bkArgs.push("--only-if-idle");
		bkArgs.push("-i", String(idleMinutes), "-p", String(cdpPort));

		const userDataDir = envString("BK_USER_DATA_DIR", "");
		if (userDataDir) {
			bkArgs.push("--user-data-dir", userDataDir);
		}

		const recordNetwork = envString("BK_RECORD_NETWORK", "");
		if (recordNetwork) {
			bkArgs.push("--record-network", recordNetwork);
		}

		const recordIncludes = splitList(envString("BK_RECORD_INCLUDE", ""));
		for (const include of recordIncludes) {
			bkArgs.push("--record-include", include);
		}

		if (envBool("BK_HEADLESS", false)) {
			bkArgs.push("--headless");
		}

		bkProc = spawn(process.execPath, bkArgs, { stdio: "inherit", cwd: repoRoot });
		await delay(waitMs);
	}

	await runCli([...configArgs, "discover-owa", ...commonArgs, "--duration-ms", String(discoverMs)]);

	if (source === "template") {
		await runCli([...configArgs, "fetch-owa", ...commonArgs, ...windowArgs, "--json"]);
		await runCli([
			...configArgs,
			"sync",
			...commonArgs,
			...windowArgs,
			"--source",
			"template",
			...syncExtraArgs,
		]);
	} else {
		await runCli([
			...configArgs,
			"capture-owa",
			...commonArgs,
			"--duration-ms",
			String(verifyCaptureMs),
			...captureFilterArgs,
			"--json",
		]);
		await runCli([
			...configArgs,
			"sync",
			...commonArgs,
			"--source",
			"capture",
			"--capture-ms",
			String(syncCaptureMs),
			...captureFilterArgs,
			...syncExtraArgs,
		]);
	}
} finally {
	cleanup();
}
