import { defineConfig } from "vitest/config";

export default defineConfig({
	test: {
		coverage: {
			provider: "v8",
			reporter: ["text", "json", "html", "lcov"],
			include: ["src/**/*.js"],
			exclude: [
				"src/cli.js",
				"src/keepalive/keepalive.js", // Main loop with signal handlers, tested manually
				"src/google/client.js", // OAuth flow requires interactive browser
			],
		},
	},
});
