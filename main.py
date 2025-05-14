

# main.py
import pathlib
import logging
from datetime import datetime
import os # Import os needed for OS-specific date formatting in text_writer

# Use direct imports from your project structure
import pdf_parser
import data_mapper
import excel_writer  # Keep if you still want Excel output
import pdf_writer    # Keep if you still want the combined PDF output
import text_writer   # Contains the specific TXT writing function

# Import base directory configuration
try:
    from config import BASE_DIR
except ImportError:
    logging.error("Failed to import BASE_DIR from config.py. Ensure it exists.")
    # Fallback to current directory if config import fails, but this is not ideal
    BASE_DIR = pathlib.Path(__file__).parent

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,  # Ensure level is DEBUG
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    force=True  # Force this configuration, overriding others (Requires Python 3.8+)
)


# --- IMPORTANT: Set the PDF filename you want to process ---
# Update this variable to the actual PDF file you intend to run the script on.
# INPUT_PDF_FILENAME = "D & M KG-4501436938.pdf" # Example PDF from OCR D & M KG-453424501436938
INPUT_PDF_FILENAME = "D & M KG-451304501459759.pdf" # Example PDF from Translation.xlsx
# -----------------------------------------------------------

def process_order(pdf_file_path: pathlib.Path) -> bool:
    """
    Orchestrates the processing pipeline:
    1. Parse PDF to extract raw data.
    2. Map raw data to structured format (including preparing for TXT).
    3. (Optional) Write mapped data to Excel.
    4. (Optional) Write mapped data to a combined PDF.
    5. Write mapped data to the specialized Auftrag Export TXT format.

    Args:
        pdf_file_path (pathlib.Path): The absolute path to the input PDF file.

    Returns:
        bool: True if processing completed (even with warnings), False if a critical error occurred.
    """
    logging.info(f"Starting processing for PDF: {pdf_file_path.name}")
    output_dir = BASE_DIR  # Outputs will be saved in the same directory as the script
    logging.info(f"Output directory set to: {output_dir}")

    # 1. Parse PDF
    logging.info("Step 1: Parsing PDF...")
    extracted_data = pdf_parser.extract_data_from_pdf(pdf_file_path)
    if not extracted_data:
        logging.error("Failed to extract data from PDF. Aborting.")
        return False
    logging.info("PDF parsing successful.")
    logging.debug(f"Extracted data snippet: Header={extracted_data.get('header')}, Positions count={len(extracted_data.get('positions', []))}")

    # 2. Map Data
    logging.info("Step 2: Mapping extracted data...")
    # This step now also prepares data needed specifically for the TXT export, like WinkelFS_raw
    mapped_data = data_mapper.map_data_to_template(extracted_data)
    if not mapped_data:
        logging.error("Failed to map extracted data. Aborting.")
        return False
    logging.info("Data mapping successful.")
    logging.debug(f"Mapped data snippet: Kopf={mapped_data.get('kopf')}, Positions count={len(mapped_data.get('positionen', []))}")

    # 3. Generate Base Filename (used for all output files)
    logging.info("Step 3: Generating base filename...")
    kopf_values = mapped_data.get("kopf", {})
    today_str = datetime.now().strftime("%Y%m%d") # Format YYYYMMDD

    # Get order identifiers, providing fallbacks and cleaning them for use in filenames
    auftragsname = kopf_values.get("Auftragsname", "UnknownOrder")
    kunden_auftragsnr = kopf_values.get("Kunden-Auftrags-Nr", "UnknownRef")
    # Remove characters potentially problematic in filenames (allow letters, numbers, hyphen, underscore)
    auftragsname_clean = "".join(c if c.isalnum() or c in ('-', '_') else '_' for c in str(auftragsname))
    kunden_auftragsnr_clean = "".join(c if c.isalnum() or c in ('-', '_') else '_' for c in str(kunden_auftragsnr))

    base_filename = f"{today_str}_{auftragsname_clean}_{kunden_auftragsnr_clean}"
    logging.info(f"Generated base filename: {base_filename}")

    # 4. Write to Excel (Optional - Comment out if not needed)
    logging.info("Step 4: Writing data to Excel (Optional)...")
    excel_output_file = excel_writer.write_to_excel(mapped_data, output_dir, base_filename)
    if not excel_output_file:
         # Decide if Excel failure is critical. Log as error but continue?
         logging.error("Failed to write data to Excel file.")
    else:
         logging.info(f"Successfully generated Excel file -> {excel_output_file}")

    # 5. Write Combined PDF (Optional - Comment out if not needed)
    logging.info("Step 5: Writing data to Combined PDF (Optional)...")
    pdf_output_file = pdf_writer.write_combined_pdf(mapped_data, output_dir, base_filename)
    if not pdf_output_file:
        logging.warning("Failed to generate Combined PDF file.") # Treat as warning
    else:
         logging.info(f"Successfully generated Combined PDF -> {pdf_output_file}")

    # 6. Write Special Auftrag Export TXT (Primary TXT output)
    logging.info("Step 6: Writing Auftrag Export TXT data...")
    # Ensure we call the function designed for the specific semicolon-delimited format
    txt_output_file = text_writer.write_auftrag_export_txt(mapped_data, output_dir, base_filename)
    if not txt_output_file:
        # Treat TXT failure as critical error for this requirement
        logging.error("Failed to generate Auftrag Export TXT file. This might be a critical error.")
        # Depending on requirements, you might want to return False here
        # return False
    else:
         logging.info(f"Successfully generated Auftrag Export TXT file -> {txt_output_file}")

    logging.info(f"Processing finished for: {pdf_file_path.name}")
    return True # Return True indicating completion (even if optional steps failed)

# --- Main execution block ---
if __name__ == "__main__":
    # Construct the full path to the input PDF
    input_pdf_path = BASE_DIR / INPUT_PDF_FILENAME
    # Get the absolute path for clearer error messages
    input_pdf_path = input_pdf_path.resolve()

    # Check if the specified input PDF file actually exists
    if not input_pdf_path.is_file():
        logging.error("--- CRITICAL ERROR ---")
        logging.error(f"Input PDF file specified in script ('{INPUT_PDF_FILENAME}')")
        logging.error(f"was NOT found at the expected location:")
        logging.error(f"==> {input_pdf_path}")
        logging.error(f"Please ensure:")
        logging.error(f"  1. The file '{INPUT_PDF_FILENAME}' exists.")
        logging.error(f"  2. It is located in the same directory as the script ({BASE_DIR}).")
        logging.error(f"  3. The INPUT_PDF_FILENAME variable in main.py is set correctly.")
        logging.error("--- Aborting ---")
    else:
        # If the file exists, proceed with processing
        processing_successful = process_order(input_pdf_path)
        if processing_successful:
            logging.info("Script finished successfully.")
        else:
            logging.error("Script finished with critical errors.")

