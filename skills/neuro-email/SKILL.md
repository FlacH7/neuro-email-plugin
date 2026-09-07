---
name: neuro-email
description: Microsoft 365 email integration with OAuth2 for Kimi Work. Supports reading inbox, listing folders, sending emails, and searching messages.
---

# Neuro Email

Use this skill when the user wants to read, send, search, or manage their Microsoft 365 email.

## Prerequisites

This plugin requires OAuth2 authentication with a Microsoft 365 account. The first time you use it, you must authenticate via device flow.

## Authentication

### First-time setup

Run the device flow to authenticate:

```bash
python3 ~/.kimi-email/generate_msal_code.py
```

This will print a device code. The user must:
1. Open https://login.microsoft.com/device in their browser
2. Enter the code shown
3. Sign in with their Microsoft 365 account
4. Approve the access

Then complete the flow:

```bash
python3 ~/.kimi-email/complete_msal_auth.py
```

The token is saved to `~/.kimi-email/msal_cache.json` and will auto-refresh on subsequent sessions.

### Auto-connect on session start

A SessionStart hook runs automatically at the beginning of each new session to refresh the token silently if a cached token exists.

## Reading Email

To read the inbox, use:

```python
python3 ~/.kimi-email/connect_email.py
```

Then the token is available for IMAP operations.

## Sending Email

SMTP is authenticated automatically when the token is available. Use standard Python smtplib with XOAUTH2.

## Account Configuration

Default account: pedro.valdes@neuroinformatics-collaboratory.org

To change the account, edit the MAILBOX variable in `~/.kimi-email/connect_email.py`.

## Provider Info

- Provider: Microsoft 365 (Outlook)
- IMAP: outlook.office365.com:993 (SSL)
- SMTP: smtp.office365.com:587 (STARTTLS)
- Auth: OAuth2 via MSAL device flow

## Troubleshooting

- If authentication fails, the token may have expired. Re-run the device flow.
- Ensure the Azure AD app registration includes scopes: IMAP.AccessAsUser.All and SMTP.Send
- The IT admin may need to approve the device flow for SMTP.Send scope
