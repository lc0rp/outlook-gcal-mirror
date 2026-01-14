import { importOptional } from "./imports.js";

/**
 * @typedef {object} CdpConnection
 * @property {"playwright"} engine
 * @property {any} browser
 * @property {any} page
 */

/**
 * @param {{ port: number, targetUrl?: string }} opts
 * @returns {Promise<CdpConnection>}
 */
export async function connectPlaywrightOverCdp({ port, targetUrl }) {
	const mod = await importOptional(["playwright", "playwright-core"]);
	const chromium = mod.chromium;
	if (!chromium?.connectOverCDP) {
		throw new Error("Playwright is installed but chromium.connectOverCDP() is missing.");
	}

	const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`);
	const context = browser.contexts()[0] ?? (await browser.newContext());

	let page = null;
	for (const p of context.pages()) {
		if (targetUrl && p.url()?.startsWith(targetUrl)) {
			page = p;
			break;
		}
		if (!page && p.url() && p.url() !== "about:blank") {
			page = p;
		}
	}

	if (!page) {
		page = await context.newPage();
		if (targetUrl) {
			await page.goto(targetUrl, { waitUntil: "domcontentloaded" });
		}
	}

	return { engine: "playwright", browser, page };
}
