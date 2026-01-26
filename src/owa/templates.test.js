import { describe, expect, it } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";

import {
	loadTemplatesFile,
	selectTemplateFromTemplates,
	loadTemplateFromFile,
	saveTemplatesFile,
} from "./templates.js";

describe("owa/templates", () => {
	it("loadTemplatesFile returns null for missing file", async () => {
		const result = await loadTemplatesFile("/tmp/nonexistent-templates.json");
		expect(result).toBeNull();
	});

	it("saveTemplatesFile and loadTemplatesFile round-trip", async () => {
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-templates-"));
		const filePath = path.join(dir, "templates.json");
		const data = { candidates: [{ url: "https://example.com", method: "GET" }] };

		await saveTemplatesFile(filePath, data);
		const loaded = await loadTemplatesFile(filePath);
		expect(loaded).toEqual(data);
	});

	describe("selectTemplateFromTemplates", () => {
		it("returns null for null data", () => {
			expect(selectTemplateFromTemplates(null)).toBeNull();
		});

		it("selects from array format", () => {
			const data = [{ url: "https://a.com", method: "POST" }];
			const template = selectTemplateFromTemplates(data);
			expect(template).toEqual({ url: "https://a.com", method: "POST", headers: undefined, body: undefined });
		});

		it("selects from candidates property", () => {
			const data = { candidates: [{ suggestedTemplate: { url: "https://b.com", method: "GET" } }] };
			const template = selectTemplateFromTemplates(data);
			expect(template).toEqual({ url: "https://b.com", method: "GET" });
		});

		it("selects from templates property", () => {
			const data = { templates: [{ template: { url: "https://c.com", method: "PUT" } }] };
			const template = selectTemplateFromTemplates(data);
			expect(template).toEqual({ url: "https://c.com", method: "PUT" });
		});

		it("applies predicate filter", () => {
			const data = [
				{ url: "https://a.com", method: "GET" },
				{ url: "https://b.com", method: "POST" },
			];
			const template = selectTemplateFromTemplates(data, (t) => t.method === "POST");
			expect(template?.url).toBe("https://b.com");
		});
	});

	it("loadTemplateFromFile returns template", async () => {
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-templates-"));
		const filePath = path.join(dir, "templates.json");
		const data = [{ url: "https://test.com", method: "GET" }];
		await saveTemplatesFile(filePath, data);

		const template = await loadTemplateFromFile(filePath);
		expect(template?.url).toBe("https://test.com");
	});
});
