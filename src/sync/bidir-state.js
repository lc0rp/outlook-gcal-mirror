import { promises as fs } from "node:fs";
import path from "node:path";

import { UserError } from "../errors.js";

export const BIDIR_STATE_VERSION = 1;

/**
 * @typedef {{
 *  outlookId: string,
 *  googleId: string,
 *  lastFingerprint?: string,
 *  updatedAt?: string,
 * }} BidirLink
 */

/**
 * @typedef {{
 *  version: number,
 *  links: BidirLink[],
 * }} BidirState
 */

/** @returns {BidirState} */
export function createEmptyBidirState() {
	return {
		version: BIDIR_STATE_VERSION,
		links: [],
	};
}

function normalizeLink(link) {
	if (!link || typeof link !== "object") return null;
	const outlookId = String(link.outlookId ?? "").trim();
	const googleId = String(link.googleId ?? "").trim();
	if (!outlookId || !googleId) return null;
	const out = { outlookId, googleId };
	if (link.lastFingerprint && String(link.lastFingerprint).trim()) {
		out.lastFingerprint = String(link.lastFingerprint).trim();
	}
	if (link.updatedAt && String(link.updatedAt).trim()) {
		out.updatedAt = String(link.updatedAt).trim();
	}
	return out;
}

/**
 * @param {unknown} raw
 * @returns {BidirState}
 */
export function normalizeBidirState(raw) {
	if (!raw || typeof raw !== "object") return createEmptyBidirState();
	const version = Number(raw.version ?? BIDIR_STATE_VERSION);
	if (version !== BIDIR_STATE_VERSION) {
		throw new UserError(`Unsupported bidir state version: ${version}`);
	}
	const linksRaw = Array.isArray(raw.links) ? raw.links : [];
	const links = linksRaw.map(normalizeLink).filter(Boolean);
	return {
		version: BIDIR_STATE_VERSION,
		links,
	};
}

/**
 * @param {string} filePath
 * @returns {Promise<BidirState>}
 */
export async function loadBidirState(filePath) {
	try {
		const raw = await fs.readFile(filePath, "utf-8");
		return normalizeBidirState(JSON.parse(raw));
	} catch (err) {
		if (err && typeof err === "object" && "code" in err && err.code === "ENOENT") {
			return createEmptyBidirState();
		}
		if (err instanceof UserError) throw err;
		throw new UserError(`Failed to load bidir state: ${filePath}`);
	}
}

/**
 * @param {string} filePath
 * @param {BidirState} state
 */
export async function saveBidirState(filePath, state) {
	const normalized = normalizeBidirState(state);
	await fs.mkdir(path.dirname(filePath), { recursive: true });
	await fs.writeFile(filePath, JSON.stringify(normalized, null, 2), "utf-8");
}
