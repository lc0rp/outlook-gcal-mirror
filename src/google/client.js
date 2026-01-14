import { spawnSync } from "node:child_process";
import { promises as fs } from "node:fs";
import http from "node:http";
import path from "node:path";

import { UserError } from "../errors.js";

/**
 * @param {string} moduleName
 */
async function importGoogleApis(moduleName) {
	try {
		return await import(moduleName);
	} catch {
		throw new UserError(
			"Missing dependency 'googleapis'. Install it in this project (e.g. `pnpm add googleapis`)."
		);
	}
}

/**
 * @param {string} url
 */
function openUrlBestEffort(url) {
	const platform = process.platform;

	if (platform === "darwin") {
		spawnSync("open", [url], { stdio: "ignore" });
		return;
	}

	if (platform === "win32") {
		spawnSync("cmd", ["/c", "start", "", url], { stdio: "ignore" });
		return;
	}

	spawnSync("xdg-open", [url], { stdio: "ignore" });
}

/**
 * OAuth scopes needed for Calendar read/write.
 */
export const GOOGLE_CALENDAR_SCOPES = ["https://www.googleapis.com/auth/calendar"];

/**
 * @param {{ credentialsPath: string, tokenPath: string, scopes?: string[] }} opts
 */
export async function getGoogleCalendarClient({ credentialsPath, tokenPath, scopes = GOOGLE_CALENDAR_SCOPES }) {
	const { google } = await importGoogleApis("googleapis");

	const credentialsRaw = await fs.readFile(credentialsPath, "utf-8");
	const credentials = JSON.parse(credentialsRaw);
	const oauth = credentials.installed || credentials.web;
	if (!oauth?.client_id || !oauth?.client_secret) {
		throw new UserError("Google credentials JSON must contain an installed OAuth client.");
	}

	const oAuth2Client = new google.auth.OAuth2(oauth.client_id, oauth.client_secret);

	// Load existing token if present.
	try {
		const tokenRaw = await fs.readFile(tokenPath, "utf-8");
		oAuth2Client.setCredentials(JSON.parse(tokenRaw));
	} catch {
		// ignore
	}

	const hasToken = !!(
		oAuth2Client.credentials &&
		(oAuth2Client.credentials.refresh_token || oAuth2Client.credentials.access_token)
	);

	if (!hasToken) {
		await new Promise((resolve, reject) => {
			/** @type {string | null} */
			let redirectUri = null;

			const server = http.createServer(async (req, res) => {
				try {
					const requestUrl = new URL(req.url ?? "/", "http://127.0.0.1");
					if (requestUrl.pathname !== "/oauth2callback") {
						res.statusCode = 404;
						res.end();
						return;
					}

					const code = requestUrl.searchParams.get("code");
					if (!code) {
						res.statusCode = 400;
						res.end("Missing code");
						return;
					}

					if (!redirectUri) {
						res.statusCode = 500;
						res.end("OAuth callback not ready");
						return;
					}

					res.end("Authentication complete. You can close this window.");

					const { tokens } = await oAuth2Client.getToken({ code, redirect_uri: redirectUri });
					oAuth2Client.setCredentials(tokens);

					await fs.mkdir(path.dirname(tokenPath), { recursive: true });
					await fs.writeFile(tokenPath, JSON.stringify(tokens, null, 2), "utf-8");
					server.close();
					resolve();
				} catch (e) {
					server.close();
					reject(e);
				}
			});

			server.listen(0, "127.0.0.1", () => {
				const address = server.address();
				if (!address || typeof address === "string") {
					reject(new Error("Failed to start OAuth callback server"));
					return;
				}

				redirectUri = `http://127.0.0.1:${address.port}/oauth2callback`;
				const authUrl = oAuth2Client.generateAuthUrl({
					access_type: "offline",
					scope: scopes,
					prompt: "consent",
					redirect_uri: redirectUri,
				});

				openUrlBestEffort(authUrl);
				console.info("Open this URL if your browser did not open automatically:\n" + authUrl);
			});
		});
	}

	const calendar = google.calendar({ version: "v3", auth: oAuth2Client });
	return { calendar, auth: oAuth2Client };
}
