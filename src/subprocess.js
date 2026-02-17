import { execFile } from "node:child_process";
import { promisify } from "node:util";

import { UserError } from "./errors.js";
import { errorMessage } from "./utils.js";

const execFileAsync = promisify(execFile);

function tail(text, maxChars = 1200) {
	const s = String(text ?? "");
	if (s.length <= maxChars) return s;
	return s.slice(s.length - maxChars);
}

function formatCommand(command, args) {
	const parts = [command, ...(args ?? [])]
		.filter((v) => v !== undefined && v !== null)
		.map((v) => String(v));
	return parts.join(" ");
}

/**
 * @param {{ command: string, args?: string[], cwd?: string, env?: Record<string,string>, timeoutMs?: number, maxBuffer?: number, runner?: (command: string, args: string[], options: any) => Promise<{stdout?: string, stderr?: string}> }} opts
 */
export async function runCommand(opts) {
	const command = String(opts.command ?? "").trim();
	const args = (opts.args ?? []).map((v) => String(v));
	const cwd = opts.cwd;
	const env = opts.env;
	const timeoutMs = opts.timeoutMs ?? 120000;
	const maxBuffer = opts.maxBuffer ?? 10 * 1024 * 1024;
	const runner = opts.runner ?? execFileAsync;

	if (!command) {
		throw new UserError("Missing subprocess command");
	}

	try {
		const res = await runner(command, args, {
			cwd,
			env,
			timeout: timeoutMs,
			maxBuffer,
		});
		return {
			stdout: String(res?.stdout ?? ""),
			stderr: String(res?.stderr ?? ""),
		};
	} catch (err) {
		const message = [
			`Command failed: ${formatCommand(command, args)}`,
			`Reason: ${errorMessage(err)}`,
		];
		const stderr = String(err?.stderr ?? "").trim();
		if (stderr) message.push(`stderr:\n${tail(stderr)}`);
		const stdout = String(err?.stdout ?? "").trim();
		if (stdout) message.push(`stdout:\n${tail(stdout)}`);
		throw new UserError(message.join("\n"));
	}
}

function parseJsonOutput(text) {
	const trimmed = String(text ?? "").trim();
	if (!trimmed) return null;
	if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
		return JSON.parse(trimmed);
	}

	const lines = trimmed
		.split(/\r?\n/)
		.map((line) => line.trim())
		.filter(Boolean);

	for (let i = lines.length - 1; i >= 0; i -= 1) {
		const line = lines[i];
		if (!line.startsWith("{") && !line.startsWith("[")) continue;
		try {
			return JSON.parse(line);
		} catch {
			// continue
		}
	}

	throw new Error("No JSON object found in command output");
}

/**
 * @param {{ command: string, args?: string[], cwd?: string, env?: Record<string,string>, timeoutMs?: number, maxBuffer?: number, runner?: (command: string, args: string[], options: any) => Promise<{stdout?: string, stderr?: string}>, label?: string }} opts
 */
export async function runJsonCommand(opts) {
	const label = String(opts.label ?? "subprocess");
	const { stdout } = await runCommand(opts);
	try {
		const parsed = parseJsonOutput(stdout);
		if (parsed === null) {
			throw new Error("Empty output");
		}
		return parsed;
	} catch (err) {
		throw new UserError(
			`${label} returned invalid JSON. ${errorMessage(err)}\nstdout:\n${tail(stdout)}`
		);
	}
}

export const _internal = { parseJsonOutput, formatCommand, tail };
