#!/usr/bin/env bash
set -euo pipefail

if [[ $# -eq 0 ]]; then
	echo "Usage: scripts/with-gog-keyring.sh <command> [args...]" >&2
	exit 2
fi

load_keyring_password_from_env_file() {
	local env_file="${OGM_ENV_FILE:-/home/user/.config/.env}"
	[[ -r "$env_file" ]] || return 1

	local value
	value="$(awk '
		/^[[:space:]]*(export[[:space:]]+)?GOG_KEYRING_PASSWORD[[:space:]]*=/ {
			line = $0
			sub(/^[[:space:]]*(export[[:space:]]+)?GOG_KEYRING_PASSWORD[[:space:]]*=[[:space:]]*/, "", line)
			print line
			exit
		}
	' "$env_file")"
	value="${value%$'\r'}"
	[[ -n "$value" ]] || return 1

	if [[ "$value" == \"*\" && "$value" == *\" ]]; then
		value="${value:1:${#value}-2}"
	elif [[ "$value" == \'*\' && "$value" == *\' ]]; then
		value="${value:1:${#value}-2}"
	fi

	[[ -n "$value" ]] || return 1
	export GOG_KEYRING_PASSWORD="$value"
	return 0
}

load_keyring_password_from_pass() {
	command -v pass >/dev/null 2>&1 || return 1

	local pass_output keyring_password
	pass_output="$(pass show openclaw/gog_keyring_password 2>/dev/null)" || return 1
	keyring_password="$(printf '%s\n' "$pass_output" | head -n1 | tr -d '\r')"
	[[ -n "$keyring_password" ]] || return 1

	export GOG_KEYRING_PASSWORD="$keyring_password"
	return 0
}

if [[ -z "${GOG_KEYRING_PASSWORD:-}" ]]; then
	if ! load_keyring_password_from_env_file && ! load_keyring_password_from_pass; then
		echo "failed to load GOG_KEYRING_PASSWORD from /home/user/.config/.env or pass" >&2
		echo "set GOG_KEYRING_PASSWORD explicitly if needed" >&2
		exit 1
	fi
fi

exec "$@"
