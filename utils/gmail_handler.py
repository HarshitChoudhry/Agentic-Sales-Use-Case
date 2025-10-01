import os
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

class GmailHandler:
    """
    Utility to send emails using Gmail API with or without attachments.
    """

    SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

    def __init__(self):
        self.service = self.authenticate()

    def authenticate(self):
        """Authenticate with Gmail API using refresh token in .env."""
        creds = Credentials(
            token=None,
            refresh_token=os.getenv("GOOGLE_REFRESH_TOKEN"),
            token_uri="https://oauth2.googleapis.com/token",
            client_id=os.getenv("GOOGLE_CLIENT_ID"),
            client_secret=os.getenv("GOOGLE_CLIENT_SECRET"),
            scopes=self.SCOPES,
        )
        return build("gmail", "v1", credentials=creds)

    def send_email(self, subject: str, body: str, to_email: str = None, attachment_path: str = None):
        """Send an email via Gmail API. Optionally include a PDF/text attachment."""
        from_email = os.getenv("GMAIL_USER_EMAIL")
        to_email = to_email or os.getenv("NOTIFICATION_EMAIL")

        # Create base email
        message = MIMEMultipart()
        message["to"] = to_email
        message["from"] = from_email
        message["subject"] = subject

        # Add email body
        message.attach(MIMEText(body, "plain"))

        # Add attachment if provided
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as f:
                mime_base = MIMEBase("application", "octet-stream")
                mime_base.set_payload(f.read())
            encoders.encode_base64(mime_base)
            mime_base.add_header(
                "Content-Disposition",
                f'attachment; filename="{os.path.basename(attachment_path)}"',
            )
            message.attach(mime_base)

        # Encode message
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
        send_message = {"raw": raw_message}

        try:
            sent = self.service.users().messages().send(userId="me", body=send_message).execute()
            print(f"✅ Email sent to {to_email} (ID: {sent['id']})")
            return sent
        except Exception as e:
            print(f"⚠️ Failed to send email: {e}")
            return None
