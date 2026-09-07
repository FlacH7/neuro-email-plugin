# Legacy scripts

These files are **kept for reference only**. They belong to earlier development
generations and will be removed once the unified CLI (see README roadmap) lands.

- `v1-token-cache/` — first authentication generation: `device_flow.json` +
  `token_cache.json` + `full_result.json` instead of the MSAL serializable cache.
- `tests/` — ad-hoc manual test scripts for IMAP connect, SMTP send, inbox
  listing, and silent auth. **They contain personal email addresses from
  development; replace them before reuse.**
- `connect_email_flow_variant.py` — alternative `connect_email.py` that resumes
  a saved `msal_flow.json` instead of using the `TokenManager` class.

Do not build new features here.
