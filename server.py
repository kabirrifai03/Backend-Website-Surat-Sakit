from fastapi import FastAPI, APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from dotenv import load_dotenv
from starlette.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
import os
import logging
import json
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict
from typing import List, Optional
import uuid
from datetime import datetime, timezone
import shutil

# Import utility modules
from utils.docx_renderer import render_docx_with_docxtpl, create_docx_from_scratch
from utils.pdf_converter import convert_docx_to_pdf
from utils.excel_exporter import export_records_to_excel
from utils.email_sender import send_letter_to_admin, get_admin_contact_info


ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

# MongoDB connection
mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

# Create the main app without a prefix
app = FastAPI()

# Create a router with the /api prefix
api_router = APIRouter(prefix="/api")

# Ensure directories exist
TEMPS_DIR = ROOT_DIR / 'temps'
UPLOADS_DIR = ROOT_DIR / 'uploads'
TEMPLATES_DIR = ROOT_DIR / 'templates'

TEMPS_DIR.mkdir(exist_ok=True)
UPLOADS_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)


# Define Models
class SickLetterRecord(BaseModel):
    model_config = ConfigDict(extra="ignore")
    
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    name: str
    gender: str
    age: str
    occupation: str
    address: str
    duration: str
    from_date: str
    to_date: str
    height: str
    weight: str
    notes: str
    clinic_address: str
    clinic_name: str
    clinic_type: str
    doctor_name: Optional[str] = "dr. [Nama Dokter]"
    doctor_nip: Optional[str] = "[NIP Dokter]"
    letter_number: Optional[str] = None
    date_issued: Optional[str] = None
    location: Optional[str] = "Jakarta"
    payment_confirmed: Optional[bool] = False
    created_at: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))


class PaymentConfirmation(BaseModel):
    record_id: str
    confirmed: bool = True


# Add your routes to the router instead of directly to app
@api_router.get("/")
async def root():
    return {"message": "Medical Absence Letter Generator API", "version": "1.0.0"}


@api_router.get("/admin-contact")
async def get_admin_contact():
    """
    Get admin contact information for user notifications
    """
    return get_admin_contact_info()


def format_indonesian_date(date_str):
    bulan = {
        "01": "Januari", "02": "Februari", "03": "Maret", "04": "April",
        "05": "Mei", "06": "Juni", "07": "Juli", "08": "Agustus",
        "09": "September", "10": "Oktober", "11": "November", "12": "Desember",
    }

    # Input pasti format YYYY-MM-DD dari HTML <input type="date">
    try:
        tahun = date_str[:4]
        bulan_num = date_str[5:7]
        hari = str(int(date_str[8:10]))  # Buang leading zero
        return f"{hari} {bulan[bulan_num]} {tahun}"
    except:
        return date_str



@api_router.post("/generate")
async def generate_sick_letter(
    data: str = Form(...),
    logo: Optional[UploadFile] = File(None),
    paper_size: str = Form("A4")
):
    try:
        # Parse JSON data
        patient_data = json.loads(data)

        # Format tanggal ke format Indonesia
        patient_data['from_date'] = format_indonesian_date(patient_data['from_date'])
        patient_data['to_date'] = format_indonesian_date(patient_data['to_date'])
        patient_data['date_issued'] = format_indonesian_date(
            patient_data.get('date_issued', datetime.now().strftime("%Y-%m-%d"))
        )

        
        # Generate letter number if not provided
        if not patient_data.get('letter_number'):
            patient_data['letter_number'] = f"SKS-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
        
        # Set date issued if not provided
        if not patient_data.get('date_issued'):
            patient_data['date_issued'] = datetime.now().strftime('%d/%m/%Y')
        
        # Handle logo upload
        logo_path = None
        if logo:
            logo_filename = f"logo_{uuid.uuid4()}{Path(logo.filename).suffix}"
            logo_path = UPLOADS_DIR / logo_filename
            
            with open(logo_path, "wb") as buffer:
                shutil.copyfileobj(logo.file, buffer)
        
        # Check if template exists, otherwise create from scratch
        template_path = TEMPLATES_DIR / 'sick_letter.docx'
        
        if template_path.exists():
            docx_path = render_docx_with_docxtpl(
                patient_data, 
                str(logo_path) if logo_path else None
            )
        else:
            docx_path = create_docx_from_scratch(
                patient_data, 
                str(logo_path) if logo_path else None,
                paper_size
            )

        # =========================================
        # HILANGKAN PDF CONVERSION 100%
        # Langsung pakai docx_path
        # =========================================
        file_path = docx_path

        # Save to database
        record = SickLetterRecord(**patient_data)
        doc = record.model_dump()
        doc['created_at'] = doc['created_at'].isoformat()
        await db.sick_letters.insert_one(doc)

        # Clean logo
        if logo_path and logo_path.exists():
            os.remove(logo_path)

        filename = (
            f"Surat_Keterangan_Dokter_"
            f"{patient_data['name'].replace(' ', '_')}.docx"
        )

        # Kirim ke email admin
        email_result = send_letter_to_admin(patient_data, file_path)

        # Setelah sukses kirim email, hapus file biar tidak menumpuk
        if os.path.exists(file_path):
            os.remove(file_path)

        return {
            "success": True,
            "message": "Surat berhasil dibuat dan telah dikirim ke admin.",
            "sent_to": "achmadkabir003@gmail.com",
            "email_status": email_result
        }


    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"Invalid JSON data: {str(e)}")
    except Exception as e:
        logging.error(f"Error generating sick letter: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error generating letter: {str(e)}")


@api_router.get("/records", response_model=List[SickLetterRecord])
async def get_records():
    """
    Get all sick letter records from database
    """
    try:
        records = await db.sick_letters.find({}, {"_id": 0}).to_list(1000)
        
        # Convert ISO string timestamps back to datetime objects
        for record in records:
            if isinstance(record.get('created_at'), str):
                record['created_at'] = datetime.fromisoformat(record['created_at'])
        
        return records
    except Exception as e:
        logging.error(f"Error fetching records: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error fetching records: {str(e)}")


@api_router.post("/export-excel")
async def export_excel():
    """
    Export all sick letter records to Excel file
    """
    try:
        # Fetch all records
        records = await db.sick_letters.find({}, {"_id": 0}).to_list(1000)
        
        if not records:
            raise HTTPException(status_code=404, detail="No records found to export")
        
        # Generate Excel file
        excel_path = export_records_to_excel(records)
        
        return FileResponse(
            path=excel_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=Path(excel_path).name,
            headers={"Content-Disposition": f"attachment; filename={Path(excel_path).name}"}
        )
        
    except Exception as e:
        logging.error(f"Error exporting to Excel: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error exporting to Excel: {str(e)}")


@api_router.post("/confirm-payment")
async def confirm_payment(confirmation: PaymentConfirmation):
    """
    Confirm payment for a sick letter record
    """
    try:
        result = await db.sick_letters.update_one(
            {"id": confirmation.record_id},
            {"$set": {"payment_confirmed": confirmation.confirmed}}
        )
        
        if result.modified_count == 0:
            raise HTTPException(status_code=404, detail="Record not found")
        
        return {"success": True, "message": "Payment confirmed"}
        
    except Exception as e:
        logging.error(f"Error confirming payment: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error confirming payment: {str(e)}")


# Include the router in the main app
app.include_router(api_router)

app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()
