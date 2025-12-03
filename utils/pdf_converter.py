"""PDF Conversion utilities"""
import os
import sys
import logging
from pathlib import Path

logger = logging.getLogger(__name__)


def convert_docx_to_pdf(docx_path: str, use_win32: bool = False) -> str:
    """
    Convert DOCX to PDF using available methods
    
    Args:
        docx_path: Path to the DOCX file
        use_win32: Use pywin32 COM automation (Windows only)
    
    Returns:
        Path to the generated PDF file
    """
    pdf_path = docx_path.replace('.docx', '.pdf')
    
    # Method 1: Try docx2pdf (works on Windows with Office installed)
    if use_win32 and sys.platform == 'win32':
        try:
            from docx2pdf import convert
            convert(docx_path, pdf_path)
            logger.info(f"PDF generated using docx2pdf: {pdf_path}")
            return pdf_path
        except Exception as e:
            logger.warning(f"docx2pdf conversion failed: {str(e)}")
    
    # Method 2: Try LibreOffice (Linux/cross-platform)
    if sys.platform in ['linux', 'linux2', 'darwin']:
        try:
            import subprocess
            # Try using LibreOffice for conversion
            result = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'pdf', 
                 '--outdir', str(Path(docx_path).parent), docx_path],
                capture_output=True,
                timeout=30
            )
            
            if result.returncode == 0 and os.path.exists(pdf_path):
                logger.info(f"PDF generated using LibreOffice: {pdf_path}")
                return pdf_path
            else:
                logger.warning(f"LibreOffice conversion failed: {result.stderr.decode()}")
        except Exception as e:
            logger.warning(f"LibreOffice conversion failed: {str(e)}")
    
    # Method 3: Fallback - Return DOCX path if no conversion method works
    # In production, you might want to use a cloud service or weasyprint for HTML->PDF
    logger.warning("No PDF conversion method available. Returning DOCX file.")
    logger.info("Consider installing LibreOffice for PDF conversion: apt-get install libreoffice")
    return docx_path


def convert_docx_to_pdf_win32(docx_path: str) -> str:
    """
    Convert DOCX to PDF using pywin32 COM automation (Windows only)
    This method requires Microsoft Word to be installed
    
    Args:
        docx_path: Path to the DOCX file
    
    Returns:
        Path to the generated PDF file
    """
    if sys.platform != 'win32':
        raise RuntimeError("pywin32 method only works on Windows")
    
    try:
        import win32com.client
        
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        
        doc = word.Documents.Open(os.path.abspath(docx_path))
        
        pdf_path = docx_path.replace('.docx', '.pdf')
        
        # wdFormatPDF = 17
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close()
        word.Quit()
        
        logger.info(f"PDF generated using pywin32: {pdf_path}")
        return pdf_path
        
    except Exception as e:
        logger.error(f"pywin32 conversion failed: {str(e)}")
        raise
