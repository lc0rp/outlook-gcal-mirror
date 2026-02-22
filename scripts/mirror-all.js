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

const configArgs = configPath ? ["--config", configPath] : [];
const syncArgs = ["sync"];

optionalArgs("--cli365-bin", envString("OGM_CLI365_BIN", ""), syncArgs);
optionalArgs("--cli365-config", envString("OGM_CLI365_CONFIG", cfg?.bidir?.cli365ConfigPath ?? ""), syncArgs);
optionalArgs("--cli365-cdp-port", envNumber("OGM_CLI365_CDP_PORT", cfg?.bidir?.cli365CdpPort ?? ""), syncArgs);
optionalArgs("--cli365-folder", envString("OGM_CLI365_FOLDER", ""), syncArgs);
if (envBool("OGM_CLI365_ENSURE_CDP", cfg?.bidir?.cli365EnsureCdp ?? false)) {
	syncArgs.push("--cli365-ensure-cdp");
}
optionalArgs(
	"--cli365-ensure-cdp-timeout",
	envString("OGM_CLI365_ENSURE_CDP_TIMEOUT", cfg?.bidir?.cli365EnsureCdpTimeout ?? ""),
	syncArgs
);

optionalArgs("--window-days", envString("OGM_WINDOW_DAYS", ""), syncArgs);
optionalArgs("--lookback-days", envString("OGM_LOOKBACK_DAYS", ""), syncArgs);
optionalArgs("--google-credentials", envString("OGM_GOOGLE_CREDS", ""), syncArgs);
optionalArgs("--google-token", envString("OGM_GOOGLE_TOKEN", ""), syncArgs);
optionalArgs("--calendar", envString("OGM_CALENDAR_NAME", ""), syncArgs);

if (envBool("OGM_MARK_CANCELLED", false)) syncArgs.push("--mark-cancelled");
if (envBool("OGM_DRY_RUN", false)) syncArgs.push("--dry-run");
if (envBool("OGM_NO_LOG_EVENTS", false)) syncArgs.push("--no-log-events");

await runCli([...configArgs, ...syncArgs]);
