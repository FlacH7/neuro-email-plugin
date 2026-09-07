"""
Email Quick Connect for Kimi Work
Usage: python3 ~/.kimi-email/connect_email.py
This automatically uses cached tokens or guides you through device flow.
"""
import os
import sys
import msal
import imaplib
import smtplib
import base64
import json

CLIENT_ID = "e2469caf-4b8d-4c34-8af4-02a09b6ed6d3"
TENANT_ID = "b0d60629-de22-4251-96f3-1d43de1f3201"
MAILBOX = "pedro.valdes@neuroinformatics-collaboratory.org"
SCOPES = [
    "https://outlook.office.com/IMAP.AccessAsUser.All",
    "https://outlook.office.com/SMTP.Send",
]
AUTH_DIR = os.path.expanduser("~/.kimi-email")
CACHE_FILE = os.path.join(AUTH_DIR, "msal_cache.json")
FLOW_FILE = os.path.join(AUTH_DIR, "msal_flow.json")

def ensure_auth_dir():
    os.makedirs(AUTH_DIR, exist_ok=True)

def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache

def save_cache(cache):
    with open(CACHE_FILE, "w") as f:
        f.write(cache.serialize())

def get_token():
    ensure_auth_dir()
    cache = load_cache()
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )

    # Try silent auth first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            save_cache(cache)
            return result["access_token"]

    # If there's a saved flow, try completing it
    if os.path.exists(FLOW_FILE):
        print("Found pending device flow. Completing...")
        with open(FLOW_FILE, "r") as f:
            flow = json.load(f)
        result = app.acquire_token_by_device_flow(flow)
        save_cache(cache)
        os.remove(FLOW_FILE)
        if "access_token" in result:
            return result["access_token"]
        raise RuntimeError(f"Flow failed: {result}")

    # Need new device flow
    print("No cached token. Starting device flow...")
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(flow)

    with open(FLOW_FILE, "w") as f:
        json.dump(flow, f)

    print("\n" + "=" * 60)
    print(flow["message"])
    print("=" * 60)
    print("\nApprove the code above, then run this script again.")
    return None

def test_connections(token):
    auth_string = f"user={MAILBOX}\x01auth=Bearer {token}\x01\x01"
    imap_ok = False
    smtp_ok = False

    print("\n📡 Testing IMAP...")
    try:
        imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
        imap.authenticate("XOAUTH2", lambda _: auth_string.encode("utf-8"))
        _, count = imap.select("INBOX")
        print(f"   ✅ IMAP: {count[0].decode()} messages")
        imap.logout()
        imap_ok = True
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
            smtp_ok = True
        else:
            print(f"   ❌ SMTP failed: {code}")
    except Exception as e:
        print(f"   ❌ SMTP failed: {e}")

    return imap_ok, smtp_ok

def main():
    token = get_token()
    if token is None:
        return  # User needs to approve first

    print(f"✅ Authenticated as: {MAILBOX}")
    imap_ok, smtp_ok = test_connections(token)

    if imap_ok and smtp_ok:
        print("\n🎉 Email ready! IMAP and SMTP are working.")
    else:
        print("\n⚠️  Connection issues. Check errors above.")

if __name__ == "__main__":
    main()
