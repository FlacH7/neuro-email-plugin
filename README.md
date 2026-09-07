# Neuro Email

A Kimi plugin for **Microsoft 365 email with OAuth2**. Read your inbox, list folders, and send email from Kimi Work using IMAP/SMTP with XOAUTH2 authentication — no passwords stored, tokens auto-refresh via MSAL.

> ⚠️ **Development status:** early development (`v0.2.0-dev`). The authentication flow is implemented but not yet completed end-to-end (see [Development status](#development-status) and [docs/DEVELOPMENT-STATUS.md](docs/DEVELOPMENT-STATUS.md)).

---

## Features

- 🔐 **OAuth2 device-flow authentication** — sign in via https://login.microsoft.com/device, no password ever touches the plugin
- 🔄 **Persistent MSAL token cache with silent refresh** — approve once, then tokens refresh automatically across sessions
- ⚡ **SessionStart auto-connect hook** — the plugin re-authenticates silently every time a Kimi session starts
- 📥 **IMAP read** — connect to `outlook.office365.com:993` and list inbox messages
- 📤 **SMTP send** — send mail through `smtp.office365.com:587` (STARTTLS + XOAUTH2)
- 🧩 **Kimi plugin packaging** — manifest, skill instructions, and session hook ready for the personal marketplace

---

## Development status

Current state as of **2026-09-07** (full detail in [docs/DEVELOPMENT-STATUS.md](docs/DEVELOPMENT-STATUS.md)):

| Area | Status | Notes |
| --- | --- | --- |
| OAuth2 device flow (initiate) | ✅ Implemented | `scripts/generate_msal_code.py` |
| OAuth2 device flow (complete + cache) | ✅ Implemented | `scripts/complete_msal_auth.py` |
| Silent token refresh | ✅ Implemented | `scripts/connect_email.py` (`TokenManager`) |
| SessionStart auto-connect hook | ✅ Implemented | `hooks/session-start.sh` |
| IMAP connection + inbox listing | ✅ Implemented | Tested logic against O365 endpoints |
| SMTP auth + send | ✅ Implemented | XOAUTH2 via `AUTH XOAUTH2` |
| **End-to-end approved token on file** | ❌ Not yet | Last device codes expired before approval (AADSTS70020); cache is empty — first task on a new machine is a fresh approval |
| Unified CLI (`list/search/read/send`) | 🚧 Planned | See [Roadmap](#roadmap--next-steps) |
| Configurable account (env vars) | 🚧 Planned | Mailbox is currently hardcoded |
| Encrypted token storage | 🚧 Planned | Cache is plaintext JSON today |

### Known issues

- **Redundant script generations.** `scripts/` holds the current MSAL-cache flow; `legacy/` contains the first generation (plain `token_cache.json`) and ad-hoc test scripts. Consolidation is planned.
- **Hardcoded mailbox** (`pedro.valdes@neuroinformatics-collaboratory.org`) in every script.
- **Hardcoded Azure AD app registration** (`CLIENT_ID` / `TENANT_ID`) — acceptable for a public-client MSAL app, but should move to config/env.
- **Plaintext token cache** at `~/.kimi-email/msal_cache.json`.
- **IT consent**: the `SMTP.Send` scope may require tenant admin approval before sending works.

---

## Architecture

```
neuro-email-plugin/
├── kimi.plugin.json           # Plugin manifest (name, version, hooks, skills)
├── hooks/
│   └── session-start.sh       # SessionStart hook: silent auto-connect
├── scripts/                   # Current generation (canonical)
│   ├── generate_msal_code.py  # Start OAuth2 device flow, save msal_flow.json
│   ├── complete_msal_auth.py  # Complete flow after browser approval, save cache
│   └── connect_email.py       # TokenManager: silent auth + IMAP/SMTP tests
├── skills/
│   └── neuro-email/
│       └── SKILL.md           # Agent-facing instructions loaded by Kimi
├── legacy/                    # Earlier generations — kept for reference
│   ├── connect_email_flow_variant.py
│   ├── v1-token-cache/        # First generation (device_flow.json + token_cache.json)
│   └── tests/                 # Ad-hoc IMAP/SMTP/inbox test scripts
├── docs/
│   └── DEVELOPMENT-STATUS.md  # Detailed status + next steps
└── README.md
```

## How authentication works

```
┌──────────────┐     ┌─────────────────────────┐     ┌────────────────────┐
│  First run   │────▶│ generate_msal_code.py   │────▶│ Device code shown  │
└──────────────┘     └─────────────────────────┘     └─────────┬──────────┘
                                                               │
                            User opens https://login.microsoft.com/device
                            and approves the code ◀────────────┘
                                                               │
┌──────────────┐     ┌─────────────────────────┐     ┌─────────▼──────────┐
│ Token cached │◀────│ complete_msal_auth.py   │◀────│ Microsoft issues   │
│ + IMAP/SMTP  │     │ (saves msal_cache.json) │     │ tokens             │
│   tested     │     └─────────────────────────┘     └────────────────────┘
└──────┬───────┘
       │ Subsequent sessions
       ▼
┌─────────────────────────┐     ┌──────────────────────────────┐
│ hooks/session-start.sh  │────▶│ connect_email.py             │
│ (runs on SessionStart)  │     │ acquire_token_silent()       │
└─────────────────────────┘     │ → IMAP + SMTP stay alive     │
                                └──────────────────────────────┘
```

Provider endpoints:

| Protocol | Host | Port | Security |
| --- | --- | --- | --- |
| IMAP | `outlook.office365.com` | 993 | SSL |
| SMTP | `smtp.office365.com` | 587 | STARTTLS |
| Auth | OAuth2 via MSAL device flow | — | Scopes: `IMAP.AccessAsUser.All`, `SMTP.Send` |

## Getting started (development)

Requirements: Python 3.9+ and `msal`:

```bash
pip install msal
```

### First-time authentication

```bash
# 1. Start the device flow — it prints a code
python3 scripts/generate_msal_code.py

# 2. Open https://login.microsoft.com/device, enter the code,
#    sign in with your Microsoft 365 account and approve.

# 3. Complete the flow — this saves the token cache and tests IMAP + SMTP
python3 scripts/complete_msal_auth.py

# 4. From now on, silent refresh keeps you connected
python3 scripts/connect_email.py
```

The token cache lives at `~/.kimi-email/msal_cache.json` and is **never** committed to git (see `.gitignore`).

## Installing as a Kimi plugin

### Path A — register via Kimi conversation

In a Kimi conversation, ask:

> "Register this local plugin in my personal market: `<absolute path to this repo>`"

Then open Kimi's Plugins page → **「个人」 (Personal)** tab → click **＋** on "Neuro Email". No restart needed.

### Path B — register via CLI

**Windows (PowerShell):**

```powershell
$share = "$env:APPDATA\kimi-desktop\daimon-share"
& "$env:APPDATA\kimi-desktop\daimon-bundle\bin\kimi-daimon.cmd" kimi-plugin register-personal "C:\full\path\to\neuro-email-plugin" --share-dir "$share"
```

**macOS / Linux:**

```bash
kimi-daimon kimi-plugin register-personal ~/path/to/neuro-email-plugin \
  --share-dir ~/Library/Application\ Support/kimi-desktop/daimon-share   # macOS
# or
kimi-daimon kimi-plugin register-personal ~/path/to/neuro-email-plugin \
  --share-dir ~/.kimi/daimon-share                                        # Linux
```

## Usage

Once a valid token exists, ask Kimi in natural language, e.g.:

| Request | Action |
| --- | --- |
| _"Check my inbox"_ | List recent messages via IMAP |
| _"Send an email to …"_ | Send via SMTP with XOAUTH2 |
| _"How many unread emails do I have?"_ | IMAP `SEARCH UNSEEN` |

## Configuration

| Setting | Where | Default |
| --- | --- | --- |
| Mailbox | Constant `MAILBOX` in each script | `pedro.valdes@neuroinformatics-collaboratory.org` |
| Azure AD `CLIENT_ID` | Constant in scripts | app registration for this plugin |
| Azure AD `TENANT_ID` | Constant in scripts | organization tenant |
| Token cache | `~/.kimi-email/msal_cache.json` | auto-created |

> Planned: move mailbox and app IDs to environment variables (`NEURO_EMAIL_MAILBOX`, `NEURO_EMAIL_CLIENT_ID`, `NEURO_EMAIL_TENANT_ID`) — see the roadmap.

## Security notes

- **Tokens are secrets.** `msal_cache.json`, `msal_flow.json`, `device_flow.json`, `token_cache.json`, `full_result.json`, and `auth_output.txt` are gitignored. Never commit them.
- `CLIENT_ID` / `TENANT_ID` belong to a **public client** MSAL application; they identify the app but cannot be used alone to access any mailbox. Still, treat app-registration changes as sensitive.
- The token cache is **plaintext** at rest in your home directory. Roadmap item: OS keyring integration.
- Test scripts under `legacy/tests/` contain personal addresses used during development — replace before reuse.

## Roadmap / next steps

Prioritized for continuing development on a second machine:

1. **Complete a fresh device-flow approval** and confirm IMAP + SMTP green end-to-end (`generate_msal_code.py` → approve → `complete_msal_auth.py`).
2. **Unify the CLI**: one `scripts/neuro_email.py` with subcommands — `auth login|status|logout`, `mail list [--folder --unread --limit]`, `mail read <uid>`, `mail search <query>`, `mail send --to --subject --body [--attach]`, `mail move`, `mail mark-read`. Absorb `TokenManager` from `connect_email.py`.
3. **Configuration via environment variables** for mailbox, client ID, tenant ID (keep constants as fallback).
4. **Real mail operations**: folder listing, `SEARCH` (unseen/from/subject/date), body fetch (plain/HTML + attachments), move/mark-read/delete with `--confirm-write`.
5. **Token storage hardening**: encrypt the MSAL cache or delegate to the OS keyring.
6. **Skill instructions (`SKILL.md`) expansion**: document every CLI subcommand with copy-paste examples for the agent.
7. **Housekeeping**: delete or formally deprecate `legacy/`, add `requirements.txt`, add smoke tests + CI (GitHub Actions), GitHub releases.
8. **Optional**: Graph API migration (`Mail.Read`/`Mail.Send`) as an alternative to IMAP/SMTP for richer operations.

## Troubleshooting

### `expired_token: AADSTS70020`

The device code expired before approval (15 min limit). Re-run `generate_msal_code.py` and approve promptly.

### `Authentication failed` on IMAP/SMTP

- Token expired → re-run the device flow (silent refresh only works with a non-empty cache).
- The Azure AD app registration must include scopes `IMAP.AccessAsUser.All` and `SMTP.Send`.
- Tenant admins can disable IMAP/SMTP AUTH or require admin consent for `SMTP.Send`.

### SessionStart hook prints "No cached credentials found"

`~/.kimi-email/msal_cache.json` is missing or empty — run the first-time authentication steps.

### `msal` not found

```bash
pip install msal
```

## Changelog

### v0.2.0-dev (2026-09-07)

- Repository structure aligned with `kimi-zotero-plugin` conventions
- Consolidated canonical scripts under `scripts/`; moved earlier generations to `legacy/`
- Added `.gitignore`, MIT `LICENSE`, and `docs/DEVELOPMENT-STATUS.md`

### v0.1.0 (2026-09-02)

- Initial OAuth2 device-flow authentication with MSAL
- Persistent token cache + silent refresh (`TokenManager`)
- SessionStart auto-connect hook
- IMAP inbox read and SMTP send via XOAUTH2

## License

MIT — see [LICENSE](LICENSE).
