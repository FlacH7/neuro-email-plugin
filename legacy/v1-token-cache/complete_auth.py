import json
import msal
import imaplib
import os

CLIENT_ID = "e2469caf-4b8d-4c34-8af4-02a09b6ed6d3"
TENANT_ID = "b0d60629-de22-4251-96f3-1d43de1f3201"
MAILBOX = "pedro.valdes@neuroinformatics-collaboratory.org"
AUTH_DIR = os.path.expanduser("~/.kimi-email")

with open(os.path.join(AUTH_DIR, "device_flow.json"), "r") as f:
    flow = json.load(f)

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
)

print("Completing device flow...")
result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    raise RuntimeError(f"{result.get('error')}: {result.get('error_description')}")

# Save FULL result including refresh_token if available
with open(os.path.join(AUTH_DIR, "full_result.json"), "w") as f:
    json.dump(result, f)

# Also save just access_token for backward compatibility
with open(os.path.join(AUTH_DIR, "token_cache.json"), "w") as f:
    json.dump({"access_token": result["access_token"]}, f)

print("Token acquired and saved!")
print(f"Keys in result: {list(result.keys())}")

# Test IMAP
auth = f"user={MAILBOX}\x01auth=Bearer {result['access_token']}\x01\x01"
print(f"\nTesting IMAP connection for {MAILBOX}...")
try:
    imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
    imap.authenticate("XOAUTH2", lambda _: auth.encode("utf-8"))
    status, folders = imap.list()
    if status == "OK":
        print(f"Folders found: {len(folders)}")
    status, count = imap.select("INBOX")
    if status == "OK":
        print(f"INBOX messages: {count[0].decode()}")
    imap.logout()
    print("\n✅ IMAP connection successful!")
except Exception as e:
    print(f"\n❌ IMAP connection failed: {e}")
