import json
import os
import msal

CLIENT_ID = "e2469caf-4b8d-4c34-8af4-02a09b6ed6d3"
TENANT_ID = "b0d60629-de22-4251-96f3-1d43de1f3201"
SCOPES = [
    "https://outlook.office.com/IMAP.AccessAsUser.All",
    "https://outlook.office.com/SMTP.Send",
]
AUTH_DIR = os.path.expanduser("~/.kimi-email")
FLOW_FILE = os.path.join(AUTH_DIR, "msal_flow.json")
CODE_FILE = os.path.join(AUTH_DIR, "msal_code.txt")

os.makedirs(AUTH_DIR, exist_ok=True)

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
)

flow = app.initiate_device_flow(scopes=SCOPES)
if "user_code" not in flow:
    raise RuntimeError(flow)

# Save flow for later completion
with open(FLOW_FILE, "w") as f:
    json.dump(flow, f)

with open(CODE_FILE, "w") as f:
    f.write(flow["message"])

print("Device code generated and saved!")
print("=" * 60)
print(flow["message"])
print("=" * 60)
