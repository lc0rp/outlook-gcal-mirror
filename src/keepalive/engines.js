/**
 * Browser engine abstraction layer.
 * Supports Playwright and Puppeteer.
 */

function isPuppeteerCouldNotFindChromeError(err) {
	const message = err instanceof Error ? err.message : String(err ?? "");
	return message.toLowerCase().includes("could not find chrome");
}

function defaultPuppeteerChannel() {
	return "chrome";
}

function isPlaywrightMissingSystemChannelError(err, channel) {
	if (!channel) return false;
	const message = err instanceof Error ? err.message : String(err ?? "");
	const m = message.toLowerCase();
	return (
		m.includes(`chromium distribution '${channel}' is not found`) ||
		m.includes(`chromium distribution \"${channel}\" is not found`) ||
		m.includes("executable doesn't exist") ||
		m.includes("executable does not exist") ||
		(m.includes("channel") && m.includes("not supported"))
	);
}

function preferredPlaywrightChannels() {
	return ["chrome", "msedge"];
}

/**
 * Normalize and validate CDP port.
 * @param {number | string | null | undefined} value
 * @returns {number | null}
 */
export function normalizePort(value) {
	if (value === undefined || value === null || value === "") {
		return null;
	}
	const n = Number(value);
	if (!Number.isInteger(n) || n <= 0 || n > 65535) {
		throw new Error("CDP port must be an integer between 1 and 65535");
	}
	return n;
}

/**
 * Build Chromium launch arguments.
 * @param {{ cdpPort: number | null }} options
 * @returns {string[]}
 */
export function buildChromiumArgs({ cdpPort }) {
	const args = [];
	if (cdpPort) {
		args.push(`--remote-debugging-port=${cdpPort}`);
		args.push("--remote-debugging-address=127.0.0.1");
	}
	return args;
}

/**
 * Format error message with cause.
 * @param {string} message
 * @param {unknown} err
 * @returns {string}
 */
export function withCause(message, err) {
	if (err instanceof Error && err.message) {
		return `${message}\nCause: ${err.message}`;
	}
	return message;
}

/**
 * Create a session wrapper from a browser page.
 * @param {string} engine
 * @param {object} page
 * @param {object} browser
 * @param {number | null} cdpPort
 * @returns {object}
 */
export function createSession(engine, page, browser, cdpPort) {
	return {
		engine,
		page,
		cdpPort,
		async close() {
			await browser.close();
		},
		async goto(url, options) {
			return await page.goto(url, options);
		},
		async reload(options) {
			return await page.reload(options);
		},
		async currentUrl() {
			return page.url();
		},
	};
}

/**
 * Import a module dynamically with error handling.
 * Exported for testing purposes.
 * @param {string} moduleName
 * @returns {Promise<object>}
 */
export async function importModule(moduleName) {
	return await import(moduleName);
}

/**
 * Launch Playwright browser.
 * @param {{ headless: boolean, cdpPort: number | null, userDataDir?: string | null, _import?: function }} options
 * @returns {Promise<object>}
 */
export async function launchPlaywright({ headless, cdpPort, userDataDir = null, _import = importModule }) {
	let mod;
	try {
		mod = await _import("playwright");
	} catch (err) {
		const message =
			"Failed to import 'playwright'. Install it in this project (e.g. `npm i playwright`).";
		throw new Error(withCause(message, err));
	}

	const chromium = mod.chromium;
	if (!chromium) {
		throw new Error("'playwright' was imported but `chromium` export was not found.");
	}

	const args = buildChromiumArgs({ cdpPort });
	if (!headless) {
		args.push("--start-maximized");
	}

	const contextOptions = !headless ? { viewport: null } : {};

	const launch = async (channel) => {
		const opts = { headless, args };
		if (channel) {
			opts.channel = channel;
		}
		return await chromium.launch(opts);
	};

	const launchPersistent = async (channel) => {
		if (!chromium.launchPersistentContext) {
			throw new Error("'playwright' was imported but `chromium.launchPersistentContext()` was not found.");
		}
		const opts = { headless, args, ...contextOptions };
		if (channel) {
			opts.channel = channel;
		}
		return await chromium.launchPersistentContext(userDataDir, opts);
	};

	const doLaunch = userDataDir ? launchPersistent : launch;

	let browser;
	for (const channel of preferredPlaywrightChannels()) {
		try {
			browser = await doLaunch(channel);
			break;
		} catch (err) {
			if (isPlaywrightMissingSystemChannelError(err, channel)) {
				continue;
			}
			throw err;
		}
	}

	if (!browser) {
		browser = await doLaunch(null);
	}

	if (userDataDir) {
		const context = browser;
		const existing = typeof context.pages === "function" ? context.pages() : [];
		const page = existing[0] || (await context.newPage());
		return createSession("playwright", page, context, cdpPort);
	}

	const context = await browser.newContext(contextOptions);
	const page = await context.newPage();

	return createSession("playwright", page, browser, cdpPort);
}

/**
 * Launch Puppeteer browser.
 * @param {{ headless: boolean, cdpPort: number | null, userDataDir?: string | null, _import?: function }} options
 * @returns {Promise<object>}
 */
export async function launchPuppeteer({ headless, cdpPort, userDataDir = null, _import = importModule }) {
	let mod;
	try {
		mod = await _import("puppeteer");
	} catch (err) {
		const message =
			"Failed to import 'puppeteer'. Install it in this project (e.g. `npm i puppeteer`).";
		throw new Error(withCause(message, err));
	}

	const puppeteer = mod.default ?? mod;
	if (!puppeteer?.launch) {
		throw new Error("'puppeteer' was imported but no `launch()` function was found.");
	}

	const args = buildChromiumArgs({ cdpPort });
	if (!headless) {
		args.push("--start-maximized");
	}

	const launch = async (channel) => {
		const opts = { headless, args };
		if (userDataDir) {
			opts.userDataDir = userDataDir;
		}
		if (!headless) {
			opts.defaultViewport = null;
		}
		if (channel) {
			opts.channel = channel;
		}
		return await puppeteer.launch(opts);
	};

	const preferredChannel = defaultPuppeteerChannel();
	let browser;
	try {
		browser = await launch(preferredChannel);
	} catch (err) {
		if (preferredChannel && isPuppeteerCouldNotFindChromeError(err)) {
			browser = await launch(null);
		} else {
			throw err;
		}
	}

	const page = await browser.newPage();

	return createSession("puppeteer", page, browser, cdpPort);
}

/**
 * Launch a browser using the specified engine.
 * @param {"playwright" | "puppeteer"} engine
 * @param {{ headless?: boolean, cdpPort?: number | null, userDataDir?: string | null, _import?: function }} options
 * @returns {Promise<object>}
 */
export async function launchEngine(engine, options = {}) {
	const headless = options.headless === true;
	const cdpPort = normalizePort(options.cdpPort);
	const userDataDir = options.userDataDir ?? null;
	const _import = options._import || importModule;

	if (engine === "playwright") {
		return await launchPlaywright({ headless, cdpPort, userDataDir, _import });
	}

	if (engine === "puppeteer") {
		return await launchPuppeteer({ headless, cdpPort, userDataDir, _import });
	}

	throw new Error(`Unknown engine: ${engine}`);
}
