import json
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
FLOW_FILE = os.path.join(AUTH_DIR, "msal_flow.json")
CACHE_FILE = os.path.join(AUTH_DIR, "msal_cache.json")

# Load or create token cache
cache = msal.SerializableTokenCache()
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, "r") as f:
        cache.deserialize(f.read())

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    token_cache=cache,
)

# Load saved flow
with open(FLOW_FILE, "r") as f:
    flow = json.load(f)

print("Completing device flow...")
result = app.acquire_token_by_device_flow(flow)

# Save cache immediately
cache_data = cache.serialize()
with open(CACHE_FILE, "w") as f:
    f.write(cache_data)

if "access_token" not in result:
    raise RuntimeError(f"{result.get('error')}: {result.get('error_description')}")

print(f"✅ Token acquired! Account: {result.get('id_token_claims', {}).get('preferred_username', 'unknown')}")
print(f"   Token will be auto-refreshed on next run via acquire_token_silent()")

# Clean up flow file
os.remove(FLOW_FILE)

# Test connections
token = result["access_token"]
auth_string = f"user={MAILBOX}\x01auth=Bearer {token}\x01\x01"

print("\n📡 Testing IMAP...")
try:
    imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
    imap.authenticate("XOAUTH2", lambda _: auth_string.encode("utf-8"))
    _, count = imap.select("INBOX")
    print(f"   ✅ IMAP: {count[0].decode()} messages")
    imap.logout()
except Exception as e:
    print(f"   ❌ IMAP failed: {e}")

print("📡 Testing SMTP...")
try:
    auth_b64 = base64.b64encode(auth_string.encode("utf-8")).decode("ascii")
    server = smtplib.SMTP("smtp.office365.com", 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    code, _ = server.docmd("AUTH", f"XOAUTH2 {auth_b64}")
    server.quit()
    if code == 235:
        print("   ✅ SMTP: Authenticated")
    else:
        print(f"   ❌ SMTP failed: {code}")
except Exception as e:
    print(f"   ❌ SMTP failed: {e}")
