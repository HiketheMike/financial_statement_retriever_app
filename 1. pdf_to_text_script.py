import fitz         # PyMuPDF
from PIL import Image
import pytesseract
import io
from pathlib import Path
import os

# point to your tesseract exe if not in PATH
# For Streamlit Community Cloud deployment, this line should be commented out
# as Tesseract will be installed as a system package and pytesseract will find it.
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# The try-except block for local Tesseract path is also not needed for cloud deployment
# as Tesseract is installed via packages.txt.
# try:
#     pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# except pytesseract.TesseractNotFoundError:
#     print("Tesseract-OCR is not found in the specified path or system PATH. OCR functionality may not work.")
#     print("Please install Tesseract-OCR or update the path in pdf_to_text_script.py.")


def run_pdf_to_text_process(company_folder_name, periods_to_process, extraction_method):
    # Change: Use Path(company_folder_name) to make it relative to the repo root
    company_base_path = Path(company_folder_name)
    base_pdf_dir = company_base_path / "financial_statements"
    ocr_dir = company_base_path / "text_statements"

    ocr_dir.mkdir(parents=True, exist_ok=True)

    results = []
    results.append(f"--- Starting PDF Text Extraction Process ({extraction_method.upper()} method) ---")

    for period in periods_to_process:
        pdf_path = base_pdf_dir / f"{period}.pdf"
        out_txt = ocr_dir / f"{period}_ocr.txt"

        if not pdf_path.exists():
            error_message = f"Warning: PDF file not found for {period} at {pdf_path}. Skipping text extraction for this period."
            print(error_message)
            results.append(error_message)
            continue

        status_message = f"\nProcessing PDF for period: {period} ({pdf_path}) using {extraction_method.upper()}..."
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
            status_message = f"Text output for {period} saved to: {out_txt}"
            print(status_message)
            results.append(status_message)

        except Exception as e:
            error_message = f"An error occurred during {extraction_method.upper()} for {period} at {pdf_path}: {e}. Skipping this period."
            print(error_message)
            results.append(error_message)
        finally:
            if doc:
                doc.close()
    
    results.append("\n--- PDF Text Extraction Process Complete ---")
    return "\n".join(results)