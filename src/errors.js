export class UserError extends Error {
	/** @param {string} message */
	constructor(message) {
		super(message);
		this.name = "UserError";
	}
}
