import { connectPlaywrightOverCdp } from "./playwright.js";
import { connectPuppeteerOverCdp } from "./puppeteer.js";

/**
 * @param {{ engine: "playwright" | "puppeteer", port: number, targetUrl?: string }} opts
 */
export async function connectOverCdp(opts) {
	if (opts.engine === "playwright") {
		return await connectPlaywrightOverCdp({ port: opts.port, targetUrl: opts.targetUrl });
	}
	if (opts.engine === "puppeteer") {
		return await connectPuppeteerOverCdp({ port: opts.port, targetUrl: opts.targetUrl });
	}
	throw new Error(`Unknown engine: ${opts.engine}`);
}
