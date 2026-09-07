import json
import imaplib

# Read saved token
with open("token_cache.json", "r") as f:
    token_data = json.load(f)
access_token = token_data["access_token"]

MAILBOX = "pedro.valdes@neuroinformatics-collaboratory.org"

auth = (
    f"user={MAILBOX}\x01"
    f"auth=Bearer {access_token}\x01\x01"
)

try:
    print(f"Connecting to outlook.office365.com:993 ...")
    imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
    print("Connected. Authenticating with OAuth2...")
    imap.authenticate("XOAUTH2", lambda _: auth.encode("utf-8"))
    print("Authenticated!")
    
    status, folders = imap.list()
    if status == "OK":
        print(f"Folders found: {len(folders)}")
        for f in folders[:10]:
            print(f"  {f.decode()}")
    
    status, count = imap.select("INBOX")
    if status == "OK":
        print(f"INBOX messages: {count[0].decode()}")
    
    imap.logout()
    print("\n✅ IMAP connection successful!")
except Exception as e:
    print(f"\n❌ IMAP connection failed: {e}")
