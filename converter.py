import subprocess
from pdf2docx import Converter
import os
import sys

def docx_to_pdf(input_path, output_dir=None, log_callback=None):
    if log_callback:
        log_callback("Starting DOCX to PDF conversion...")
    
    if not output_dir:
        output_dir = os.path.dirname(input_path) or "."
    
    os.makedirs(output_dir, exist_ok=True)
    
    base_name = os.path.basename(input_path)
    name_without_ext = os.path.splitext(base_name)[0]
    expected_output = os.path.join(output_dir, name_without_ext + ".pdf")
    
    if log_callback:
        log_callback(f"Running LibreOffice headless conversion to {output_dir}...")
    
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        input_path,
        "--outdir", output_dir
    ]
    
    result = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )
    
    if result.returncode != 0:
        error_msg = result.stderr or result.stdout or "LibreOffice failed with non-zero exit code"
        if log_callback:
            log_callback(f"Error during LibreOffice execution: {error_msg}")
        raise RuntimeError(f"LibreOffice failed: {error_msg}")
        
    if not os.path.exists(expected_output):
        if log_callback:
            log_callback("Error: LibreOffice completed but output PDF not found.")
        raise FileNotFoundError(f"Output PDF not found: {expected_output}")
        
    if log_callback:
        log_callback(f"Successfully converted to {expected_output}")
    return expected_output


def pdf_to_docx(input_path, output_dir=None, log_callback=None):
    if log_callback:
        log_callback("Starting PDF to DOCX conversion...")
        
    if not output_dir:
        output_dir = os.path.dirname(input_path) or "."
        
    os.makedirs(output_dir, exist_ok=True)
    
    base_name = os.path.basename(input_path)
    name_without_ext = os.path.splitext(base_name)[0]
    expected_output = os.path.join(output_dir, name_without_ext + ".docx")
    
    if log_callback:
        log_callback(f"Initializing pdf2docx Converter for {input_path}...")
    
    cv = Converter(input_path)
    
    if log_callback:
        log_callback(f"Converting pages to {expected_output}...")
        
    try:
        cv.convert(expected_output, start=0, end=None)
    finally:
        cv.close()
        
    if not os.path.exists(expected_output):
        raise FileNotFoundError(f"Output DOCX not found: {expected_output}")
        
    if log_callback:
        log_callback(f"Successfully converted to {expected_output}")
    return expected_output