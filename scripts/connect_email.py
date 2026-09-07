import json
import os
import msal
import imaplib
import smtplib
import base64

# Configuration
CLIENT_ID = "e2469caf-4b8d-4c34-8af4-02a09b6ed6d3"
TENANT_ID = "b0d60629-de22-4251-96f3-1d43de1f3201"
MAILBOX = "pedro.valdes@neuroinformatics-collaboratory.org"
SCOPES = [
    "https://outlook.office.com/IMAP.AccessAsUser.All",
    "https://outlook.office.com/SMTP.Send",
]
AUTH_DIR = os.path.expanduser("~/.kimi-email")
CACHE_FILE = os.path.join(AUTH_DIR, "msal_cache.json")

os.makedirs(AUTH_DIR, exist_ok=True)


class TokenManager:
    """Manages OAuth tokens using MSAL's persistent token cache."""

    def __init__(self):
        self.cache = msal.SerializableTokenCache()
        self._load_cache()
        self.app = msal.PublicClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            token_cache=self.cache,
        )

    def _load_cache(self):
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, "r") as f:
                self.cache.deserialize(f.read())

    def _save_cache(self):
        with open(CACHE_FILE, "w") as f:
            f.write(self.cache.serialize())

    def get_token(self):
        """Get a valid access token, using silent auth if possible."""
        accounts = self.app.get_accounts()

        if accounts:
            # Try silent authentication first
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self._save_cache()
                return result["access_token"]

        # No valid token in cache — need device flow
        print("No valid token in cache. Starting device flow...")
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(flow)

        print("=" * 60)
        print(flow["message"])
        print("=" * 60)
        print("\nApprove the code above, then this script will continue automatically.")

        result = self.app.acquire_token_by_device_flow(flow)
        self._save_cache()

        if "access_token" not in result:
            raise RuntimeError(f"{result.get('error')}: {result.get('error_description')}")

        return result["access_token"]

    def test_connections(self):
        """Test both IMAP and SMTP connections."""
        token = self.get_token()
        auth_string = f"user={MAILBOX}\x01auth=Bearer {token}\x01\x01"

        # Test IMAP
        print("\n📡 Testing IMAP...")
        try:
            imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
            imap.authenticate("XOAUTH2", lambda _: auth_string.encode("utf-8"))
            _, count = imap.select("INBOX")
            print(f"   ✅ IMAP: {count[0].decode()} messages in INBOX")
            imap.logout()
            imap_ok = True
        except Exception as e:
            print(f"   ❌ IMAP failed: {e}")
            imap_ok = False

        # Test SMTP
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
                print(f"   ✅ SMTP: Authenticated")
                smtp_ok = True
            else:
                print(f"   ❌ SMTP auth failed: {code}")
                smtp_ok = False
        except Exception as e:
            print(f"   ❌ SMTP failed: {e}")
            smtp_ok = False

        return imap_ok, smtp_ok


if __name__ == "__main__":
    tm = TokenManager()
    imap_ok, smtp_ok = tm.test_connections()

    if imap_ok and smtp_ok:
        print("\n🎉 Email fully configured and ready!")
    else:
        print("\n⚠️  Something went wrong. Check the errors above.")
