"""
Email Auth Quick Setup for Kimi Work
Run this in any new session to load email credentials and test connection.
"""
import json
import imaplib
import smtplib
import base64
import os
import shutil

AUTH_DIR = os.path.expanduser("~/.kimi-email")
MAILBOX = "pedro.valdes@neuroinformatics-collaboratory.org"

def copy_auth_files():
    """Copy auth files from persistent storage to current directory."""
    files = ["token_cache.json", "full_result.json", "device_flow.json"]
    for f in files:
        src = os.path.join(AUTH_DIR, f)
        if os.path.exists(src):
            shutil.copy2(src, f)
            print(f"✅ Copied {f}")
        else:
            print(f"⚠️  {f} not found in {AUTH_DIR}")

def test_imap():
    with open("token_cache.json", "r") as f:
        token = json.load(f)["access_token"]
    auth = (
        f"user={MAILBOX}\x01"
        f"auth=Bearer {token}\x01\x01"
    )
    try:
        imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
        imap.authenticate("XOAUTH2", lambda _: auth.encode("utf-8"))
        _, count = imap.select("INBOX")
        print(f"📥 INBOX: {count[0].decode()} messages")
        imap.logout()
        return True
    except Exception as e:
        print(f"❌ IMAP failed: {e}")
        return False

def test_smtp():
    with open("token_cache.json", "r") as f:
        token = json.load(f)["access_token"]
    auth_string = f"user={MAILBOX}\x01auth=Bearer {token}\x01\x01"
    auth_b64 = base64.b64encode(auth_string.encode("utf-8")).decode("ascii")
    try:
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        code, _ = server.docmd("AUTH", f"XOAUTH2 {auth_b64}")
        server.quit()
        if code == 235:
            print("📤 SMTP: Authenticated")
            return True
        else:
            print(f"❌ SMTP auth failed: {code}")
            return False
    except Exception as e:
        print(f"❌ SMTP failed: {e}")
        return False

if __name__ == "__main__":
    print("Setting up email auth...")
    copy_auth_files()
    print()
    imap_ok = test_imap()
    smtp_ok = test_smtp()
    print()
    if imap_ok and smtp_ok:
        print("🎉 Email is ready! You can now send/receive emails.")
    else:
        print("⚠️  Token may be expired. Run generate_device_code.py and approve again.")
