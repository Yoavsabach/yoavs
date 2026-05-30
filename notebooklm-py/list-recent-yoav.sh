#!/usr/bin/env bash
# List the 3 most recent NotebookLM notebooks for yoav8288@gmail.com.
#
# RUN THIS ON YOUR LOCAL MACHINE — NOT in a remote container.
# Prerequisites:
#   - notebooklm-py installed:  uv pip install -e .
#   - You are signed in to yoav8288@gmail.com in Chrome or Firefox on this machine.
#
# What it does:
#   1. Checks whether the "yoav8288" notebooklm profile already has live auth.
#   2. If not, tries to extract cookies from your local browser into that profile.
#   3. Falls back to interactive `notebooklm login` if cookie extraction can't find the account.
#   4. Lists the 3 most recent notebooks.

set -euo pipefail

EMAIL="yoav8288@gmail.com"
PROFILE="yoav8288"

command -v notebooklm >/dev/null 2>&1 || {
  echo "Error: notebooklm CLI not found." >&2
  echo "Install from the vendored source:" >&2
  echo "  cd $(dirname "$0") && uv sync --frozen --extra browser --extra dev && source .venv/bin/activate" >&2
  exit 1
}

auth_ok() {
  notebooklm -p "$PROFILE" auth check --test --json 2>/dev/null \
    | grep -q '"token_fetch": *true'
}

if auth_ok; then
  echo "Auth for profile '$PROFILE' (${EMAIL}) is live."
else
  echo "No live auth for profile '$PROFILE'. Looking for ${EMAIL} in local browsers..."
  for browser in chrome firefox edge brave; do
    if notebooklm auth inspect --browser "$browser" 2>/dev/null \
        | grep -q "$EMAIL"; then
      echo "Found ${EMAIL} signed in to $browser. Importing into profile '$PROFILE'..."
      notebooklm -p "$PROFILE" login --browser-cookies "$browser" && break
    fi
  done

  if ! auth_ok; then
    cat >&2 <<EOF

Could not import auth for ${EMAIL} from any local browser.

Run an interactive login (opens a browser window):

  notebooklm -p ${PROFILE} login

Then sign in as ${EMAIL} and re-run this script.
EOF
    exit 1
  fi
fi

echo
echo "3 most recent notebooks for ${EMAIL}:"
echo
notebooklm -p "$PROFILE" list --limit 3
