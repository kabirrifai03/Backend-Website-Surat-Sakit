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


def _send_via_email(patient_data: dict, pdf_path: str) -> dict:
    """
    Send email with PDF attachment to admin
    """
    try:
        if not SMTP_EMAIL or not SMTP_PASSWORD:
            return {
                'success': False,
                'method': 'email',
                'error': 'SMTP credentials not configured'
            }
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = SMTP_EMAIL
        msg['To'] = ADMIN_EMAIL
        msg['Subject'] = f"New Medical Letter Request - {patient_data.get('name', 'Unknown')}"
        
        # Email body
        body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; padding: 20px;">
            <div style="background: linear-gradient(135deg, #10b981 0%, #14b8a6 100%); padding: 20px; border-radius: 10px; margin-bottom: 20px;">
                <h2 style="color: white; margin: 0;">🏥 New Medical Letter Request</h2>
            </div>
            
            <div style="background: #f9fafb; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
                <h3 style="color: #1f2937; margin-top: 0;">👤 Patient Information</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 5px 0;"><strong>Name:</strong></td><td>{patient_data.get('name', 'N/A')}</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>Gender:</strong></td><td>{patient_data.get('gender', 'N/A')}</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>Age:</strong></td><td>{patient_data.get('age', 'N/A')} years</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>Occupation:</strong></td><td>{patient_data.get('occupation', 'N/A')}</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>Address:</strong></td><td>{patient_data.get('address', 'N/A')}</td></tr>
                </table>
            </div>
            
            <div style="background: #fef3c7; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
                <h3 style="color: #92400e; margin-top: 0;">🏥 Medical Information</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 5px 0;"><strong>Duration:</strong></td><td>{patient_data.get('duration', 'N/A')}</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>From:</strong></td><td>{patient_data.get('from_date', 'N/A')}</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>To:</strong></td><td>{patient_data.get('to_date', 'N/A')}</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>Diagnosis:</strong></td><td>{patient_data.get('notes', 'N/A')}</td></tr>
                </table>
            </div>
            
            <div style="background: #dbeafe; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
                <h3 style="color: #1e40af; margin-top: 0;">🏢 Clinic Information</h3>
                <table style="width: 100%; border-collapse: collapse;">
                    <tr><td style="padding: 5px 0;"><strong>Clinic:</strong></td><td>{patient_data.get('clinic_name', 'N/A')} - {patient_data.get('clinic_type', 'N/A')}</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>Address:</strong></td><td>{patient_data.get('clinic_address', 'N/A')}</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>Doctor:</strong></td><td>{patient_data.get('doctor_name', 'N/A')}</td></tr>
                    <tr><td style="padding: 5px 0;"><strong>Doctor NIP:</strong></td><td>{patient_data.get('doctor_nip', 'N/A')}</td></tr>
                </table>
            </div>
            
            <div style="background: #fee2e2; padding: 15px; border-radius: 8px; border-left: 4px solid #dc2626;">
                <h3 style="color: #991b1b; margin-top: 0;">⚠️ Action Required</h3>
                <ol style="color: #7f1d1d; margin: 0; padding-left: 20px;">
                    <li>Verify payment has been received</li>
                    <li>Review the attached medical letter PDF</li>
                    <li>Contact patient via WhatsApp: <strong>{ADMIN_WHATSAPP}</strong></li>
                    <li>Send final document to patient</li>
                </ol>
            </div>
            
            <div style="margin-top: 20px; padding: 15px; background: #f3f4f6; border-radius: 8px; text-align: center;">
                <p style="margin: 0; color: #6b7280; font-size: 14px;">
                    📎 <strong>Attachment:</strong> Generated medical letter is attached to this email
                </p>
            </div>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(body, 'html'))
        
        # Attach PDF file
        if os.path.exists(pdf_path):
            filename = Path(pdf_path).name
            with open(pdf_path, 'rb') as attachment:
                part = MIMEBase('application', 'pdf')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{filename}"',
            )
            msg.attach(part)
        else:
            logger.error(f"PDF file not found: {pdf_path}")
            return {
                'success': False,
                'method': 'email',
                'error': 'PDF file not found'
            }
        
        # Send email via SMTP
        if SMTP_USE_SSL:
            # Use SSL (port 465)
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as server:
                server.login(SMTP_EMAIL, SMTP_PASSWORD)
                server.send_message(msg)
        else:
            # Use STARTTLS (port 587)
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
                server.starttls()
                server.login(SMTP_EMAIL, SMTP_PASSWORD)
                server.send_message(msg)
        
        logger.info(f"✅ Email sent successfully to {ADMIN_EMAIL}")
        logger.info(f"Patient: {patient_data.get('name')} - Letter: {patient_data.get('letter_number')}")
        
        return {
            'success': True,
            'method': 'email',
            'error': None
        }
        
    except Exception as e:
        logger.error(f"❌ Email sending failed: {str(e)}")
        return {
            'success': False,
            'method': 'email',
            'error': str(e)
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
