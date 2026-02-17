import { describe, expect, it } from "vitest";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";

import {
	BIDIR_STATE_VERSION,
	createEmptyBidirState,
	normalizeBidirState,
	loadBidirState,
	saveBidirState,
} from "./bidir-state.js";
import { UserError } from "../errors.js";

describe("sync/bidir-state", () => {
	it("creates empty state", () => {
		expect(createEmptyBidirState()).toEqual({
			version: BIDIR_STATE_VERSION,
			links: [],
		});
	});

	it("normalizes links and removes invalid ones", () => {
		const state = normalizeBidirState({
			version: BIDIR_STATE_VERSION,
			links: [
				{ outlookId: "o-1", googleId: "g-1", lastFingerprint: "x" },
				{ outlookId: "", googleId: "bad" },
			],
		});
		expect(state.links).toHaveLength(1);
		expect(state.links[0].outlookId).toBe("o-1");
	});

	it("rejects unsupported version", () => {
		expect(() => normalizeBidirState({ version: 99, links: [] })).toThrow(UserError);
	});

	it("load returns empty state when file missing", async () => {
		const state = await loadBidirState("/tmp/no-such-state.json");
		expect(state.links).toEqual([]);
	});

	it("save/load roundtrip", async () => {
		const dir = await fs.mkdtemp(path.join(os.tmpdir(), "ogm-bidir-state-"));
		const filePath = path.join(dir, "state.json");
		const input = {
			version: BIDIR_STATE_VERSION,
			links: [{ outlookId: "o-1", googleId: "g-1", lastFingerprint: "fp" }],
		};
		await saveBidirState(filePath, input);
		const loaded = await loadBidirState(filePath);
		expect(loaded).toEqual(input);
	});
});
