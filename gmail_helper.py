import os, base64, json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
DATA_DIR = "/data" if os.path.exists("/data") else os.path.dirname(__file__)
TOKEN_PATH = os.path.join(DATA_DIR, "gmail_token.json")

BUNGE_SENDERS = ["mirian.gdansky@bunge.com", "bar.exeterra@bunge.com"]
PE_SENDER_DOMAIN = "estudionecochea.com.ar"

def get_flow(client_id, client_secret, redirect_uri):
    return Flow.from_client_config(
        {
            "web": {
                "client_id": client_id,
                "client_secret": client_secret,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": [redirect_uri]
            }
        },
        scopes=SCOPES,
        redirect_uri=redirect_uri
    )

def get_credentials():
    if not os.path.exists(TOKEN_PATH):
        return None
    creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        save_credentials(creds)
    return creds if creds and creds.valid else None

def save_credentials(creds):
    with open(TOKEN_PATH, "w") as f:
        f.write(creds.to_json())

def get_gmail_service():
    creds = get_credentials()
    if not creds:
        return None
    return build("gmail", "v1", credentials=creds)

def get_pdf_attachments_from_message(service, msg_id):
    """Returns list of (filename, pdf_bytes) from a Gmail message."""
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
    attachments = []
    
    def extract_parts(parts):
        for part in parts:
            if part.get("parts"):
                extract_parts(part["parts"])
            mime = part.get("mimeType", "")
            filename = part.get("filename", "")
            if mime == "application/pdf" or filename.lower().endswith(".pdf"):
                att_id = part.get("body", {}).get("attachmentId")
                if att_id:
                    att = service.users().messages().attachments().get(
                        userId="me", messageId=msg_id, id=att_id
                    ).execute()
                    data = base64.urlsafe_b64decode(att["data"])
                    attachments.append((filename, data))
    
    payload = msg.get("payload", {})
    if payload.get("parts"):
        extract_parts(payload["parts"])
    return attachments

def fetch_new_emails(service, sender_filter, max_results=10, last_history_id=None):
    """Fetch unread emails from specific senders with PDF attachments."""
    query = f"from:({' OR '.join(sender_filter)}) has:attachment filename:pdf is:unread"
    results = service.users().messages().list(
        userId="me", q=query, maxResults=max_results
    ).execute()
    return results.get("messages", [])

def mark_as_read(service, msg_id):
    service.users().messages().modify(
        userId="me", id=msg_id,
        body={"removeLabelIds": ["UNREAD"]}
    ).execute()

def get_message_sender(service, msg_id):
    msg = service.users().messages().get(userId="me", id=msg_id, format="metadata",
        metadataHeaders=["From", "Subject", "Date"]).execute()
    headers = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
    return headers.get("From", ""), headers.get("Subject", ""), headers.get("Date", "")
