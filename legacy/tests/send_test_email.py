import json
import smtplib
import ssl
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Read saved token
with open("token_cache.json", "r") as f:
    token_data = json.load(f)
access_token = token_data["access_token"]

MAILBOX = "pedro.valdes@neuroinformatics-collaboratory.org"
RECIPIENT = "clopez@neuroinformatics-collaboratory.org"

# Create message
msg = MIMEMultipart()
msg["From"] = MAILBOX
msg["To"] = RECIPIENT
msg["Subject"] = "Test email from Kimi Work - OAuth2 SMTP"

body = """Hello,

This is a test email sent from Kimi Work using OAuth2 authentication.

If you received this, the SMTP connection is working correctly!

Best regards,
Kimi Work
"""

msg.attach(MIMEText(body, "plain"))

# Build OAuth2 string for SMTP and base64 encode it
auth_string = f"user={MAILBOX}\x01auth=Bearer {access_token}\x01\x01"
auth_b64 = base64.b64encode(auth_string.encode("utf-8")).decode("ascii")

try:
    print(f"Connecting to smtp.office365.com:587 ...")
    context = ssl.create_default_context()
    server = smtplib.SMTP("smtp.office365.com", 587)
    server.ehlo()
    server.starttls(context=context)
    server.ehlo()
    print("STARTTLS established. Authenticating with OAuth2...")
    
    code, response = server.docmd("AUTH", f"XOAUTH2 {auth_b64}")
    print(f"AUTH response: {code} {response}")
    
    if code == 235:
        print("Authenticated! Sending email...")
        server.sendmail(MAILBOX, RECIPIENT, msg.as_string())
        server.quit()
        print(f"\n✅ Email sent successfully to {RECIPIENT}!")
    else:
        print(f"\n❌ Authentication failed: {code} {response}")
        server.quit()
except Exception as e:
    print(f"\n❌ Failed to send email: {e}")
