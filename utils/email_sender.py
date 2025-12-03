"""Email sending utilities with Google Drive fallback"""
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import logging
from pathlib import Path
from dotenv import load_dotenv
import resend
import base64

logger = logging.getLogger(__name__)

from dotenv import load_dotenv
import os

BASE_DIR = Path(__file__).resolve().parent.parent  # ../backend
ENV_PATH = BASE_DIR / ".env"

load_dotenv(ENV_PATH)

# Get configuration from environment
ADMIN_EMAIL = os.environ.get('ADMIN_EMAIL', 'achmadkabir003@gmail.com')
ADMIN_WHATSAPP = os.environ.get('ADMIN_WHATSAPP', '085810746193')
SMTP_EMAIL = os.environ.get('SMTP_EMAIL')
SMTP_PASSWORD = os.environ.get('SMTP_PASSWORD')
SMTP_HOST = os.environ.get('SMTP_HOST', 'smtp.gmail.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '465'))
SMTP_USE_SSL = os.environ.get('SMTP_USE_SSL', 'true').lower() == 'true'


def send_letter_to_admin(patient_data: dict, pdf_path: str) -> dict:
    # Try Telegram first (lebih stabil)
    tg = send_to_telegram(patient_data, pdf_path)
    if tg["success"]:
        return tg

    # Coba email (kalau kamu masih mau)
    email = _send_via_email(patient_data, pdf_path)
    if email["success"]:
        return email

    return tg


import requests

def send_to_telegram(patient_data, pdf_path):
    try:
        token = os.getenv("TELEGRAM_TOKEN")
        chat_id = os.getenv("TELEGRAM_CHAT_ID")
        url = f"https://api.telegram.org/bot{token}/sendDocument"

        with open(pdf_path, "rb") as file:
            files = {"document": file}
            data = {
                "chat_id": chat_id,
                "caption": (
                    f"New Medical Letter Request\n"
                    f"Name: {patient_data.get('name')}\n"
                    f"Diagnosis: {patient_data.get('notes')}\n"
                    f"Duration: {patient_data.get('duration')}"
                )
            }

            response = requests.post(url, data=data, files=files)
        
        if response.status_code == 200:
            return {"success": True, "method": "telegram"}
        else:
            return {"success": False, "method": "telegram", "error": response.text}

    except Exception as e:
        return {"success": False, "method": "telegram", "error": str(e)}




def _send_via_email(patient_data, pdf_path):
    try:
        resend.api_key = os.getenv("RESEND_API_KEY")

        # baca PDF
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        # BODY EMAIL
        body_html = f"""
        <p><b>New Medical Letter Request</b></p>
        <p>Name: {patient_data.get('name')}</p>
        <p>Diagnosis: {patient_data.get('notes')}</p>
        <p>Duration: {patient_data.get('duration')}</p>
        <p>PDF terlampir.</p>
        """

        response = resend.Emails.send({
            "from": "no-reply@jasasuratsakit.com",
            "to": os.getenv("ADMIN_EMAIL"),
            "subject": f"New Medical Letter - {patient_data.get('name')}",
            "html": body_html,
            "attachments": [
                {
                    "filename": "surat_sakit.pdf",
                    "content": base64.b64encode(pdf_bytes).decode(),
                    "type": "application/pdf"
                }
            ]
        })

        return {
            "success": True,
            "method": "email",
            "error": None
        }

    except Exception as e:
        return {
            "success": False,
            "method": "email",
            "error": str(e)
        }


def _upload_to_drive(patient_data: dict, pdf_path: str) -> dict:
    """
    Upload PDF to Google Drive as fallback
    """
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        import pickle
        
        folder_id = os.environ.get('GOOGLE_DRIVE_FOLDER_ID')
        if not folder_id:
            return {
                'success': False,
                'method': 'drive',
                'error': 'Google Drive folder ID not configured'
            }
        
        # For now, return as not configured since we need service account
        # This requires setting up Google Cloud service account
        logger.warning("Google Drive upload not fully configured (requires service account)")
        
        return {
            'success': False,
            'method': 'drive',
            'error': 'Google Drive service account not configured'
        }
        
    except Exception as e:
        logger.error(f"❌ Google Drive upload failed: {str(e)}")
        return {
            'success': False,
            'method': 'drive',
            'error': str(e)
        }


def get_admin_contact_info() -> dict:
    """
    Get admin contact information (WhatsApp only - email is internal)
    
    Returns:
        Dictionary with admin WhatsApp only
    """
    return {
        'whatsapp': ADMIN_WHATSAPP
    }
