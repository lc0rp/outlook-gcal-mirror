import { UserError } from "../errors.js";

/**
 * Import one of the provided module names.
 * Useful when projects want to support both the full package and the "-core" variant.
 *
 * @param {string | string[]} nameOrNames
 * @returns {Promise<any>}
 */
export async function importOptional(nameOrNames) {
	const names = Array.isArray(nameOrNames) ? nameOrNames : [nameOrNames];

	for (const name of names) {
		try {
			return await import(name);
		} catch {
			// keep trying
		}
	}

	const hint =
		names.length === 1
			? `'${names[0]}'`
			: `one of: ${names.map((n) => `'${n}'`).join(", ")}`;

	throw new UserError(
		`Missing dependency ${hint}. Install it in this project (e.g. 'npm i ${names[0]}' or 'pnpm add ${names[0]}').`
	);
}
