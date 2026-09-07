import json
import msal
import os

CLIENT_ID = "e2469caf-4b8d-4c34-8af4-02a09b6ed6d3"
TENANT_ID = "b0d60629-de22-4251-96f3-1d43de1f3201"
SCOPES = [
    "https://outlook.office.com/IMAP.AccessAsUser.All",
    "https://outlook.office.com/SMTP.Send",
    "offline_access",  # Needed for refresh_token
]

AUTH_DIR = os.path.expanduser("~/.kimi-email")
os.makedirs(AUTH_DIR, exist_ok=True)

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
)

flow = app.initiate_device_flow(scopes=SCOPES)
if "user_code" not in flow:
    raise RuntimeError(flow)

with open(os.path.join(AUTH_DIR, "device_flow.json"), "w") as f:
    json.dump(flow, f)

print("=" * 60)
print(flow["message"])
print("=" * 60)
print(f"\nFlow saved to {AUTH_DIR}/device_flow.json")
print("After approving in the browser, run complete_auth.py to finish.")
