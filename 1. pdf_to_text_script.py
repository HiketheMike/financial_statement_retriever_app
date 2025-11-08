import fitz         # PyMuPDF
from PIL import Image
import pytesseract
import io
from pathlib import Path
import os

# For Streamlit Community Cloud deployment, this line should be commented out
# as Tesseract will be installed as a system package and pytesseract will find it.
# For local Windows development, uncomment and set your Tesseract path:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def run_pdf_to_text_process(company_folder_name, periods_to_process, extraction_method):
    company_base_path = Path(company_folder_name)
    base_pdf_dir = company_base_path / "financial_statements"
    ocr_dir = company_base_path / "text_statements"

    ocr_dir.mkdir(parents=True, exist_ok=True)

    results = []
    results.append(f"--- Starting PDF Text Extraction Process ({extraction_method.upper()} method) ---")

    processed_any_pdf = False # Track if any PDF was successfully processed
    for period in periods_to_process:
        pdf_path = base_pdf_dir / f"{period}.pdf"
        out_txt = ocr_dir / f"{period}_ocr.txt"

        if not pdf_path.exists():
            # Changed: Force backslashes in output path
            error_message = f"Warning: PDF file not found for {period} at {str(pdf_path).replace('/', '\\')}. Skipping text extraction for this period."
            print(error_message)
            results.append(error_message)
            continue

        # Changed: Force backslashes in output path
        status_message = f"\nProcessing PDF for period: {period} ({str(pdf_path).replace('/', '\\')}) using {extraction_method.upper()}..."
        print(status_message)
        results.append(status_message)
        
        doc = None
        try:
            doc = fitz.open(pdf_path)
            with out_txt.open("w", encoding="utf-8") as fout:
                for pageno in range(len(doc)):
                    page = doc.load_page(pageno)
                    
                    if extraction_method.lower() == "ocr":
                        pix = page.get_pixmap(dpi=650)
                        img_bytes = pix.tobytes("png")
                        img = Image.open(io.BytesIO(img_bytes))
                        text = pytesseract.image_to_string(img, lang="vie", config="--psm 3") # Assuming 'vie' language pack
                    else: # extraction_method.lower() == "direct"
                        text = page.get_text("text")

                    fout.write(f"--- PAGE {pageno+1} ---\n")
                    fout.write(text + "\n\n")
            # Changed: Force backslashes in output path
            status_message = f"Text output for {period} saved to: {str(out_txt).replace('/', '\\')}"
            print(status_message)
            results.append(status_message)
            processed_any_pdf = True # Mark as successful for at least one PDF

        except Exception as e:
            # Changed: Force backslashes in output path
            error_message = f"An error occurred during {extraction_method.upper()} for {period} at {str(pdf_path).replace('/', '\\')}: {e}. Skipping this period."
            print(error_message)
            results.append(error_message)
        finally:
            if doc:
                doc.close()
    
    results.append("\n--- PDF Text Extraction Process Complete ---")
    
    if not processed_any_pdf:
        raise ValueError("No PDF files were successfully processed into text. Please check PDF paths and content.")
        
    return "\n".join(results)