import path from "node:path";
import { promises as fs } from "node:fs";

import { DEFAULT_TEMPLATES_PATH } from "../config.js";

/**
 * @param {string} filePath
 */
export async function loadTemplatesFile(filePath = DEFAULT_TEMPLATES_PATH) {
	try {
		const raw = await fs.readFile(filePath, "utf-8");
		return JSON.parse(raw);
	} catch {
		return null;
	}
}

/**
 * @param {any} data
 */
export function selectTemplateFromTemplates(data, predicate) {
	if (!data) return null;

	const list = Array.isArray(data)
		? data
		: data.candidates ?? data.templates ?? data.entries ?? [];

	for (const item of list) {
		if (!item) continue;
		const template = item.suggestedTemplate
			? item.suggestedTemplate
			: item.template
				? item.template
				: item.url && item.method
					? { url: item.url, method: item.method, headers: item.headers, body: item.body }
					: null;
		if (!template) continue;
		if (predicate && !predicate(template, item)) continue;
		return template;
	}

	return null;
}

/**
 * @param {string} filePath
 */
export async function loadTemplateFromFile(filePath = DEFAULT_TEMPLATES_PATH, predicate) {
	const data = await loadTemplatesFile(filePath);
	return selectTemplateFromTemplates(data, predicate);
}

/**
 * @param {string} filePath
 * @param {any} data
 */
export async function saveTemplatesFile(filePath = DEFAULT_TEMPLATES_PATH, data) {
	await fs.mkdir(path.dirname(filePath), { recursive: true });
	await fs.writeFile(filePath, JSON.stringify(data, null, 2), "utf-8");
}
