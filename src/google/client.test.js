import { describe, expect, it, vi, beforeEach } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";

import { GOOGLE_CALENDAR_SCOPES, getGoogleCalendarClient } from "./client.js";

// Mock googleapis with hoisted variables
const mockCalendar = { events: { list: vi.fn() } };

class MockOAuth2 {
	constructor() {
		this.credentials = {};
	}
	setCredentials(creds) {
		this.credentials = creds;
	}
	generateAuthUrl() {
		return "https://accounts.google.com/auth";
	}
	getToken() {
		return Promise.resolve({ tokens: { access_token: "tok", refresh_token: "ref" } });
	}
}

vi.mock("googleapis", () => ({
	google: {
		auth: {
			OAuth2: MockOAuth2,
		},
		calendar: vi.fn(() => mockCalendar),
	},
}));

describe("google/client", () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	it("exports GOOGLE_CALENDAR_SCOPES", () => {
		expect(GOOGLE_CALENDAR_SCOPES).toContain("https://www.googleapis.com/auth/calendar");
	});

	it("loads existing token and returns client", async () => {
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-google-"));
		const credPath = path.join(dir, "credentials.json");
		const tokenPath = path.join(dir, "token.json");

		// Write test credentials
		await fs.writeFile(
			credPath,
			JSON.stringify({ installed: { client_id: "test-id", client_secret: "test-secret" } })
		);
		await fs.writeFile(tokenPath, JSON.stringify({ refresh_token: "existing-token" }));

		const { calendar, auth } = await getGoogleCalendarClient({ credentialsPath: credPath, tokenPath });
		expect(calendar).toBe(mockCalendar);
		expect(auth).toBeInstanceOf(MockOAuth2);
		expect(auth.credentials.refresh_token).toBe("existing-token");
	});

	it("throws UserError for invalid credentials", async () => {
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-google-"));
		const credPath = path.join(dir, "credentials.json");
		const tokenPath = path.join(dir, "token.json");

		await fs.writeFile(credPath, JSON.stringify({ web: {} })); // missing client_id/secret

		await expect(getGoogleCalendarClient({ credentialsPath: credPath, tokenPath })).rejects.toThrow(
			/must contain an installed OAuth client/
		);
	});

	it("works with web credentials format", async () => {
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-google-"));
		const credPath = path.join(dir, "credentials.json");
		const tokenPath = path.join(dir, "token.json");

		await fs.writeFile(
			credPath,
			JSON.stringify({ web: { client_id: "web-id", client_secret: "web-secret" } })
		);
		await fs.writeFile(tokenPath, JSON.stringify({ access_token: "existing" }));

		const { calendar, auth } = await getGoogleCalendarClient({ credentialsPath: credPath, tokenPath });
		expect(calendar).toBe(mockCalendar);
	});
});
