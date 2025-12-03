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
    """
    Send generated medical letter to admin via email
    Falls back to Google Drive upload if email fails
    
    Args:
        patient_data: Dictionary containing patient information
        pdf_path: Path to the generated PDF file
    
    Returns:
        Dict with success status and method used
    """
    result = {
        'success': False,
        'method': None,
        'error': None
    }
    
    # Try email first
    email_result = _send_via_email(patient_data, pdf_path)
    if email_result['success']:
        return email_result
    
    # If email fails, try Google Drive
    logger.warning(f"Email sending failed: {email_result['error']}. Trying Google Drive...")
    drive_result = _upload_to_drive(patient_data, pdf_path)
    
    return drive_result if drive_result['success'] else email_result




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
            "from": os.getenv("SMTP_EMAIL"),
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
