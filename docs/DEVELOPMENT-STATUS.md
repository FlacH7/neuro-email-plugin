# Development status

**Date:** 2026-09-07
**Version:** `0.2.0-dev`
**Primary dev machine:** Windows (Git Bash), Python 3 with `msal`

This document records the exact state of the codebase when it was packaged for
GitHub, so development can continue seamlessly from another PC.

---

## 1. What exists today

### Canonical plugin (`scripts/`, `skills/`, `hooks/`)

| File | Role | State |
| --- | --- | --- |
| `kimi.plugin.json` | Plugin manifest | ✅ Registered locally via `register-personal`; name `neuro-email`, MIT, SessionStart hook configured |
| `scripts/generate_msal_code.py` | Starts the OAuth2 device flow | ✅ Saves `~/.kimi-email/msal_flow.json`; prints the user code |
| `scripts/complete_msal_auth.py` | Completes the flow after browser approval | ✅ Saves `~/.kimi-email/msal_cache.json` (MSAL serializable cache) and tests IMAP + SMTP |
| `scripts/connect_email.py` | `TokenManager` class | ✅ Silent auth via `acquire_token_silent()`; falls back to interactive device flow; tests IMAP + SMTP |
| `hooks/session-start.sh` | SessionStart hook | ✅ Runs `connect_email.py` silently when a non-empty cache exists |
| `skills/neuro-email/SKILL.md` | Agent instructions | ✅ Basic: auth flow, reading, sending, provider info, troubleshooting |

### Legacy material (`legacy/` — reference only)

- **v1 generation** (`legacy/v1-token-cache/`): first auth implementation using
  `device_flow.json` + plain `token_cache.json`/`full_result.json` (no MSAL cache).
- **Tests** (`legacy/tests/`): manual scripts — `check_inbox.py` (list last 5
  messages), `test_imap_oauth.py` (folder listing), `test_smtp.py` /
  `send_test_email.py` (SMTP sends), `test_silent_auth.py` (silent refresh check).
- `legacy/connect_email_flow_variant.py`: alternate `connect_email.py` that
  resumes a saved flow file instead of using `TokenManager`.

## 2. Verified behavior (from runtime evidence on the dev machine)

- Device flow initiation works and prints a valid user code.
- The completion script, silent refresh logic, and IMAP/SMTP XOAUTH2 code paths
  are implemented and exercised up to the approval step.
- **No completed approval is on file:** the attempts captured in
  `~/.kimi-email/auth_output.txt` ended with `expired_token: AADSTS70020`
  (device code expired before the user approved), and `msal_cache.json` is
  empty (`{}`). ⇒ **First task on the new machine: run a fresh
  `generate_msal_code.py` → approve → `complete_msal_auth.py` and confirm both
  IMAP and SMTP report success.**

## 3. Known issues and technical debt

1. **Two script generations coexist** (`scripts/` vs `legacy/`). Pick the MSAL
   cache generation and delete the rest after the unified CLI exists.
2. **Hardcoded mailbox** `pedro.valdes@neuroinformatics-collaboratory.org` in
   every script.
3. **Hardcoded Azure AD app registration** (`CLIENT_ID`, `TENANT_ID`) — fine
   for a public-client app, but should be overridable via env vars.
4. **Plaintext token cache** at `~/.kimi-email/msal_cache.json`.
5. **No real mail operations yet** — only connect/test/list-last-5. No search,
   no body fetch, no move/mark/delete, no attachment handling.
6. **No packaging hygiene**: no `requirements.txt`, no tests, no CI, no
   automated versioning of `kimi.plugin.json`.
7. **Possible tenant consent blocker**: `SMTP.Send` may need admin approval in
   the `neuroinformatics-collaboratory.org` tenant.
8. Test scripts contain personal addresses (`clopez@…`, `charlis2992@yahoo.com`).

## 4. Next steps (ordered)

1. **Authenticate end-to-end** on the new PC and save a working cache; record
   the result (IMAP/SMTP ✅) in this file.
2. **Unified CLI** `scripts/neuro_email.py`:
   `auth login|status|logout`, `mail list [--folder --unread --limit]`,
   `mail read <uid>`, `mail search <query>`, `mail send --to --subject --body`,
   `mail move <uid> <folder>`, `mail mark-read <uid>` — write operations gated
   behind `--confirm-write`.
3. **Env-var configuration**: `NEURO_EMAIL_MAILBOX`, `NEURO_EMAIL_CLIENT_ID`,
   `NEURO_EMAIL_TENANT_ID` (constants as fallback).
4. **Token cache hardening**: OS keyring or encrypted cache file.
5. **Expand `SKILL.md`** with every CLI subcommand and example output.
6. **Repo hygiene**: `requirements.txt`, delete `legacy/`, GitHub Actions smoke
   test, tagged releases, keep `kimi.plugin.json` version in sync.
7. **Evaluate Microsoft Graph API** (`Mail.Read` / `Mail.Send`) as a
   richer-operations alternative to raw IMAP/SMTP.

## 5. Environment quick reference

- Python 3.9+; only external dependency: `msal`
- IMAP `outlook.office365.com:993` (SSL) · SMTP `smtp.office365.com:587` (STARTTLS)
- Scopes: `https://outlook.office.com/IMAP.AccessAsUser.All`,
  `https://outlook.office.com/SMTP.Send`
- Token/cache dir: `~/.kimi-email/` (created automatically; gitignored)
