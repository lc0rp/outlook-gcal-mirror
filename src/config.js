import os from "node:os";
import path from "node:path";
import { promises as fs } from "node:fs";

import { UserError } from "./errors.js";

export const DEFAULT_CONFIG_DIR = path.join(os.homedir(), ".config", "outlook-gcal-mirror");
export const DEFAULT_CONFIG_PATH = path.join(DEFAULT_CONFIG_DIR, "config.json");
export const DEFAULT_TOKEN_PATH = path.join(DEFAULT_CONFIG_DIR, "google-token.json");

/**
 * @typedef {object} MirrorConfig
 * @property {{
 *   cdpPort: number,
 *   engine: "playwright" | "puppeteer",
 *   targetUrl: string,
 *   includeCalendars?: string[],
 *   skipCalendars?: string[],
 *   includeOwnerEmails?: string[],
 *   skipOwnerEmails?: string[],
 * }} outlook
 * @property {{
 *   credentialsPath: string,
 *   tokenPath: string,
 *   calendarName: string,
 *   calendarId?: string,
 * }} google
 * @property {{
 *   windowDays: number,
 *   markCancelled: boolean,
 * }} sync
 */

/**
 * @param {string} filePath
 * @returns {Promise<MirrorConfig>}
 */
export async function loadConfig(filePath) {
	let raw;
	try {
		raw = await fs.readFile(filePath, "utf-8");
	} catch {
		throw new UserError(
			`Config not found: ${filePath}. Run 'outlook-gcal-mirror setup' first (or pass --config).`
		);
	}

	/** @type {MirrorConfig} */
	const cfg = JSON.parse(raw);
	return cfg;
}

/**
 * @param {string} filePath
 * @param {MirrorConfig} config
 */
export async function saveConfig(filePath, config) {
	await fs.mkdir(path.dirname(filePath), { recursive: true });
	await fs.writeFile(filePath, JSON.stringify(config, null, 2), "utf-8");
}
