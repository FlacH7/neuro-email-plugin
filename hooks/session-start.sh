#!/usr/bin/env bash
# SessionStart hook: auto-connect Microsoft 365 email via OAuth2
# This runs automatically at the beginning of each new session.

set -e

PLUGIN_DIR="$(cd "$(dirname "$0")/.." && pwd)"
AUTH_DIR="$HOME/.kimi-email"

# Ensure auth directory exists
mkdir -p "$AUTH_DIR"

# Run the email connection script if cache exists
if [ -f "$AUTH_DIR/msal_cache.json" ] && [ -s "$AUTH_DIR/msal_cache.json" ]; then
    echo "[neuro-email] Auto-connecting to Microsoft 365 email..."
    python3 "$PLUGIN_DIR/connect_email.py" 2>/dev/null || true
else
    echo "[neuro-email] No cached credentials found. Run 'python3 ~/.kimi-email/connect_email.py' manually to authenticate."
fi
