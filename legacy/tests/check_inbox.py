import json
import imaplib
import email
from email.header import decode_header

# Read saved token
with open("token_cache.json", "r") as f:
    token_data = json.load(f)
access_token = token_data["access_token"]

MAILBOX = "pedro.valdes@neuroinformatics-collaboratory.org"

auth = (
    f"user={MAILBOX}\x01"
    f"auth=Bearer {access_token}\x01\x01"
)

def decode_str(s):
    if s is None:
        return ""
    decoded, charset = decode_header(s)[0]
    if isinstance(decoded, bytes):
        try:
            return decoded.decode(charset or "utf-8", errors="replace")
        except:
            return decoded.decode("utf-8", errors="replace")
    return str(decoded)

imap = imaplib.IMAP4_SSL("outlook.office365.com", 993)
imap.authenticate("XOAUTH2", lambda _: auth.encode("utf-8"))

# Select INBOX
imap.select("INBOX")

# Get all UIDs
status, data = imap.uid("SEARCH", None, "ALL")
if status != "OK":
    print("Failed to search emails")
    imap.logout()
    exit(1)

uids = data[0].split()
total = len(uids)
print(f"Total emails in INBOX: {total}\n")

# Get last 5 emails (most recent)
last_5 = uids[-5:]

print("=" * 70)
print(f"{'#':<3} {'From':<35} {'Date':<20} Subject")
print("=" * 70)

for i, uid in enumerate(last_5, 1):
    status, msg_data = imap.uid("FETCH", uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE)])")
    if status == "OK" and msg_data[0]:
        raw_msg = msg_data[0][1]
        msg = email.message_from_bytes(raw_msg)
        
        from_addr = decode_str(msg.get("From", "Unknown"))
        subject = decode_str(msg.get("Subject", "(No Subject)"))
        date = msg.get("Date", "Unknown")
        
        # Truncate for display
        from_short = from_addr[:34]
        date_short = date[:19]
        subject_short = subject[:60] if len(subject) <= 60 else subject[:57] + "..."
        
        print(f"{i:<3} {from_short:<35} {date_short:<20} {subject_short}")

print("=" * 70)

# Also show more detail for each
print("\n📧 DETAIL VIEW:")
print("=" * 70)

for i, uid in enumerate(last_5, 1):
    status, msg_data = imap.uid("FETCH", uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE TO CC)])")
    if status == "OK" and msg_data[0]:
        raw_msg = msg_data[0][1]
        msg = email.message_from_bytes(raw_msg)
        
        print(f"\n[{i}] UID: {uid.decode()}")
        print(f"    From:    {decode_str(msg.get('From', 'N/A'))}")
        print(f"    To:      {decode_str(msg.get('To', 'N/A'))}")
        print(f"    Date:    {decode_str(msg.get('Date', 'N/A'))}")
        print(f"    Subject: {decode_str(msg.get('Subject', 'N/A'))}")

imap.logout()
print("\n✅ Done!")
