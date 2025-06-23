import msal
import requests
import configparser

# Load config
config = configparser.ConfigParser()
config.read("config.cfg")

CLIENT_ID = config["azure"]["clientId"]
TENANT_ID = config["azure"]["tenantId"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Read"]

# Token cache in memory
app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

# Acquire token interactively (opens browser once)
result = None
accounts = app.get_accounts()
if accounts:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
if not result:
    result = app.acquire_token_interactive(scopes=SCOPES)

if "access_token" in result:
    access_token = result["access_token"]
    print("Access token acquired!")

    # Example: Read messages from the Inbox
    headers = {"Authorization": f"Bearer {access_token}"}
    folder_name = "Inbox"  # or your custom folder
    endpoint = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_name}/messages?$top=5"

    response = requests.get(endpoint, headers=headers)
    if response.ok:
        data = response.json()
        for msg in data.get("value", []):
            print(f"From: {msg['from']['emailAddress']['address']}")
            print(f"Subject: {msg['subject']}")
            print(f"Body preview: {msg['bodyPreview'][:100]}")
            print("=" * 40)
    else:
        print("Failed to fetch messages:", response.text)
else:
    print("Error getting token:", result.get("error_description"))

# === Find Folder by Name ===
def get_folder_id(folder_name):
    url = "https://graph.microsoft.com/v1.0/me/mailFolders"
    res = requests.get(url, headers=headers)
    folders = res.json().get("value", [])
    for folder in folders:
        if folder["displayName"].lower() == folder_name.lower():
            return folder["id"]
    return None

# === Read Emails from Folder ===
def read_emails_from_folder(folder_name):
    folder_id = get_folder_id(folder_name)
    if not folder_id:
        print(f"Folder '{folder_name}' not found.")
        return

    print(f"Reading emails from folder: {folder_name}")
    url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages?$top=5"
    res = requests.get(url, headers=headers)
    if not res.ok:
        print("Error fetching messages:", res.text)
        return

    messages = res.json().get("value", [])
    for msg in messages:
        print("From:", msg["from"]["emailAddress"]["address"])
        print("Subject:", msg["subject"])
        print("Body Preview:", msg["bodyPreview"])
        print("-" * 40)

# === RUN ===
read_emails_from_folder("Clive Forms")  # Replace