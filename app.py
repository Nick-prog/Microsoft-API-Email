import msal
import requests
import configparser
from typing import Any
from bs4 import BeautifulSoup

class MicosoftEntraApp:

    def __init__(self):
        # Load config
        config = configparser.ConfigParser()
        config.read("config.cfg")

        CLIENT_ID = config["azure"]["clientId"]
        TENANT_ID = config["azure"]["tenantId"]
        AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
        self.SCOPES = ["Mail.Read"]

        # Token cache in memory
        self.app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

    def acquire_token(self) -> dict[str, Any] | dict[str, str] | dict | Any | dict[Any | str, Any] | dict[str | Any, Any | str] | None:
        accounts = self.app.get_accounts()
        result = self.app.acquire_token_silent(self.SCOPES, account=accounts[0]) if accounts else None
        if not result:
            result = self.app.acquire_token_interactive(scopes=self.SCOPES)
        return result.get("access_token") if result and "access_token" in result else None
    
    def get_nested_folder_id(self, folder_path: str, headers: dict[str, str]) -> Any | None:
        """
        Takes a path like 'Clive Forms/Upload Documents' and returns the folder ID of the final folder.
        """
        folder_names = folder_path.split("/")
        current_url = "https://graph.microsoft.com/v1.0/me/mailFolders"
        current_id = None

        for name in folder_names:
            res = requests.get(current_url, headers=headers)
            if not res.ok:
                print(f"Failed to retrieve folders at {current_url}: {res.text}")
                return None
            folders = res.json().get("value", [])
            match = next((f for f in folders if f["displayName"].lower() == name.lower()), None)
            if not match:
                print(f"Folder '{name}' not found.")
                return None
            current_id = match["id"]
            current_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{current_id}/childFolders"

        return current_id
    
    def read_emails_from_nested_folder(self, folder_path: str, headers: dict[str, str], output_file: str, received_after: str) -> None:
        folder_id = self.get_nested_folder_id(folder_path, headers)
        if not folder_id:
            print(f"Folder path '{folder_path}' not found.")
            return

        base_url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages"
        
        # Add filter if date provided
        if received_after:
            url = f"{base_url}?$top=25&$filter=receivedDateTime ge {received_after}"
        else:
            url = f"{base_url}?$top=25"

        res = requests.get(url, headers=headers)
        if not res.ok:
            print("Error fetching messages:", res.text)
            return
        
        data = res.json()
        messages = data.get("value", [])
        submissions = []

        for msg in messages:
            msg_id = msg["id"]
            full_msg_url = f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}"
            full_res = requests.get(full_msg_url, headers=headers)
            full_body = ""
            if full_res.ok:
                full_body = full_res.json().get("body", {}).get("content", "")
                soup = BeautifulSoup(full_body, 'html.parser')
                # print(soup.prettify())
                soup_text = soup.get_text("|").split("|")
                soup_text.append(soup.a["href"])

            url = data.get("@odata.nextLink", None)
            submissions.append(soup_text)

        print(len(submissions))
        return
        
        # with open(output_file, "w", encoding="utf-8") as f:
        #     f.write(f"Emails from folder: {folder_path}\n")
        #     f.write("=" * 50 + "\n")

        #     while url:
        #         res = requests.get(url, headers=headers)
        #         if not res.ok:
        #             print("Error fetching messages:", res.text)
        #             break

        #         data = res.json()
        #         messages = data.get("value", [])

        #         for msg in messages:
        #             msg_id = msg["id"]
        #             full_msg_url = f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}"
        #             full_res = requests.get(full_msg_url, headers=headers)
        #             full_body = ""
        #             if full_res.ok:
        #                 full_body = full_res.json().get("body", {}).get("content", "")


        #             f.write(f"From: {msg['from']['emailAddress']['address']}\n")
        #             f.write(f"Subject: {msg['subject']}\n")
        #             f.write(f"Received: {msg['receivedDateTime']}\n")
        #             f.write(f"Body:\n{full_body.strip()}\n")
        #             f.write("-" * 40 + "\n")


        #         url = data.get("@odata.nextLink", None)

        # print(f"Filtered emails written to '{output_file}'")

if __name__ == "__main__":
    mea = MicosoftEntraApp()

    access_token = mea.acquire_token()
    if access_token:
        headers = {"Authorization": f"Bearer {access_token}"}
        mea.read_emails_from_nested_folder(
            "Clive Forms/Upload Documents",
            headers,
            output_file="filtered_emails.txt",
            received_after="2025-07-11T00:00:00Z"
        )
    else:
        print("Failed to acquire access token.")