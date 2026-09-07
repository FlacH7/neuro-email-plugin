import os
import msal
import imaplib
import smtplib
import base64

CLIENT_ID = "e2469caf-4b8d-4c34-8af4-02a09b6ed6d3"
TENANT_ID = "b0d60629-de22-4251-96f3-1d43de1f3201"
MAILBOX = "pedro.valdes@neuroinformatics-collaboratory.org"
SCOPES = [
    "https://outlook.office.com/IMAP.AccessAsUser.All",
    "https://outlook.office.com/SMTP.Send",
]
AUTH_DIR = os.path.expanduser("~/.kimi-email")
CACHE_FILE = os.path.join(AUTH_DIR, "msal_cache.json")

# Load persistent token cache
cache = msal.SerializableTokenCache()
with open(CACHE_FILE, "r") as f:
    cache.deserialize(f.read())

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    token_cache=cache,
)

# Try silent authentication
accounts = app.get_accounts()
print(f"Accounts in cache: {len(accounts)}")
for acc in accounts:
    print(f"  - {acc.get('username', 'unknown')}")

if not accounts:
    print("No accounts in cache. Need device flow.")
    exit(1)

result = app.acquire_token_silent(SCOPES, account=accounts[0])

if result and "access_token" in result:
    print(f"✅ Silent auth successful! Token: ...{result['access_token'][-20:]}")
    
    # Test IMAP
    auth_string = f"user={MAILBOX}\x01auth=Bearer {result['access_token']}\x01\x01"
    imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
    imap.authenticate("XOAUTH2", lambda _: auth_string.encode("utf-8"))
    _, count = imap.select("INBOX")
    print(f"📥 INBOX: {count[0].decode()} messages")
    imap.logout()
    print("✅ Auto-refresh works! No manual approval needed.")
else:
    print("❌ Silent auth failed. Token may need manual refresh.")
    print(f"Error: {result}")
