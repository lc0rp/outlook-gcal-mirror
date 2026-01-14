import { importOptional } from "./imports.js";

/**
 * @typedef {object} CdpConnection
 * @property {"puppeteer"} engine
 * @property {any} browser
 * @property {any} page
 */

/**
 * @param {{ port: number, targetUrl?: string }} opts
 * @returns {Promise<CdpConnection>}
 */
export async function connectPuppeteerOverCdp({ port, targetUrl }) {
	const mod = await importOptional(["puppeteer", "puppeteer-core"]);
	const puppeteer = mod.default ?? mod;
	if (!puppeteer?.connect) {
		throw new Error("Puppeteer is installed but connect() is missing.");
	}

	const browser = await puppeteer.connect({ browserURL: `http://127.0.0.1:${port}` });

	let page = null;
	for (const p of await browser.pages()) {
		if (targetUrl && p.url()?.startsWith(targetUrl)) {
			page = p;
			break;
		}
		if (!page && p.url() && p.url() !== "about:blank") {
			page = p;
		}
	}

	if (!page) {
		page = await browser.newPage();
		if (targetUrl) {
			await page.goto(targetUrl, { waitUntil: "domcontentloaded" });
		}
	}

	return { engine: "puppeteer", browser, page };
}
