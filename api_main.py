
# api_main.py
import fastapi
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
import shutil
import pathlib
import logging
import uuid # For unique temporary filenames
import os # For path manipulation
from datetime import datetime # Ensure datetime is imported
import re # Ensure re is imported

# --- Import your existing logic ---
# Use direct imports assuming all files are in the same root directory
try:
    import pdf_parser
    import data_mapper
    import excel_writer
    import pdf_writer
    import text_writer
    from pdf_auto.config import BASE_DIR # BASE_DIR should point to the project root
except ImportError as e:
     print(f"ERROR: Could not import processing modules. Ensure they are accessible.")
     print(f"Error details: {e}")
     # Consider exiting if core modules are missing
     # raise SystemExit("Core processing modules not found.")
     pass

# --- Configure Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Define Temporary and Output Directories ---

BASE_DIR = pathlib.Path(__file__).resolve().parent
# These directories should exist relative to where api_main.py is located
TEMP_DIR = BASE_DIR / "temp_files"
OUTPUT_DIR_API = BASE_DIR / "api_output" # Separate directory for API outputs
TEMP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR_API.mkdir(exist_ok=True)

# --- Create FastAPI app ---
app = FastAPI(title="PDF Processing API", description="Processes D&M KG PDF files.")

# --- Refactored Processing Logic ---
def run_processing_task(input_pdf_path: pathlib.Path, base_filename: str):
    """
    Runs the core PDF processing steps (Parse -> Map -> Write Outputs).
    Returns a dictionary with results including paths to generated files or error info.
    """
    results = {"success": False, "files": {}, "error": None}
    # Use the dedicated API output directory defined above
    output_dir = OUTPUT_DIR_API

    try:
        logging.info(f"API Task: Starting processing for {input_pdf_path.name} -> Output base: {base_filename}")

        # 1. Parse PDF
        extracted_data = pdf_parser.extract_data_from_pdf(input_pdf_path)
        if not extracted_data:
            # Check if positions list specifically is missing/empty
            if not extracted_data.get("positions"):
                 logging.error("PDF Parsing completed but found 0 positions.")
                 # You might still want to proceed to generate Kopf-only files
                 # Or raise specific error here. Let's raise for now.
                 raise ValueError("PDF Parsing failed to identify any position items.")
            else:
                 raise ValueError("PDF Parsing failed for unknown reasons.")

        logging.info(f"PDF Parsing successful. Found {len(extracted_data.get('positions',[]))} positions.")

        # 2. Map Data
        mapped_data = data_mapper.map_data_to_template(extracted_data)
        if not mapped_data:
            # This check might be redundant if parser ensures structure, but keep for safety
            raise ValueError("Failed to map extracted data (mapper returned None).")
        logging.info("Data mapping successful.")

        # 3. Write Excel
        excel_file = excel_writer.write_to_excel(mapped_data, output_dir, base_filename)
        if excel_file:
             results["files"]["excel"] = excel_file
             logging.info(f"Successfully generated Excel: {excel_file}")
        else:
             logging.warning("Failed to generate Excel file.") # Continue processing other formats

        # 4. Write Combined PDF
        pdf_file = pdf_writer.write_combined_pdf(mapped_data, output_dir, base_filename)
        if pdf_file:
             results["files"]["pdf"] = pdf_file
             logging.info(f"Successfully generated PDF: {pdf_file}")
        else:
             logging.warning("Failed to generate Combined PDF.")

        # 5. Write TXT
        txt_file = text_writer.write_auftrag_export_txt(mapped_data, output_dir, base_filename)
        if txt_file:
             results["files"]["txt"] = txt_file
             logging.info(f"Successfully generated TXT: {txt_file}")
        else:
             logging.warning("Failed to generate TXT file.")

        # Mark as overall success if at least one file was potentially generated
        # (or adjust based on which outputs are mandatory)
        if results["files"]:
             results["success"] = True
             logging.info(f"API Task: Processing finished for {base_filename}. Generated: {list(results['files'].keys())}")
        else:
             # If even Excel failed, consider it a failure
             results["success"] = False
             results["error"] = "Failed to generate any output files."
             logging.error(f"API Task: Processing failed for {base_filename}, no output files created.")


    except Exception as e:
        logging.error(f"API Task: Error during processing for {base_filename}: {e}", exc_info=True)
        results["error"] = f"Processing error: {str(e)}" # Provide clearer error
        results["success"] = False # Ensure success is false on exception
    # finally block removed for cleanup - cleanup happens in the endpoint now

    return results


# --- API Endpoint Definition ---
@app.post("/process_pdf/",
          summary="Process Uploaded PDF",
          description="Upload a D&M KG PDF file to extract data and generate Excel, PDF, and TXT reports.",
          response_description="Returns status and relative paths to generated files.",
          response_model=dict # Basic dict response for now
          )
async def process_pdf_endpoint(
    # background_tasks: BackgroundTasks, # Keep if needed for background option
    file: UploadFile = File(..., description="The D&M KG PDF file to process.") # Added description
    ):
    """
    API endpoint to upload a PDF, process it, and return file paths.
    """
    # --- Input Validation ---
    if not file or not file.filename:
         logging.warning("API: Received request with missing file or filename.")
         raise HTTPException(status_code=400, detail="No file or filename provided.")
    if not file.filename.lower().endswith(".pdf"):
        logging.warning(f"API: Received invalid file type: {file.filename}")
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a PDF.")

    temp_pdf_path = None # Initialize outside try
    try:
        # Create a unique temporary file path
        temp_id = uuid.uuid4()
        # Sanitize filename: replace non-alphanumeric (excluding . and -) with underscore
        safe_filename = re.sub(r'[^\w\.-]', '_', file.filename)
        temp_pdf_path = TEMP_DIR / f"{temp_id}_{safe_filename}"

        # Save the uploaded file temporarily
        logging.info(f"API: Receiving file {file.filename}...")
        with temp_pdf_path.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        logging.info(f"API: Saved PDF temporarily to {temp_pdf_path}")

        # --- Generate Base Filename for output files ---
        today_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_filename = f"{today_str}_{temp_id}" # Unique base name

        # --- Run processing SYNCHRONOUSLY ---
        # If processing takes > ~30-60s, consider background tasks
        logging.info(f"API: Starting synchronous processing task for {base_filename}...")
        results = run_processing_task(temp_pdf_path, base_filename)
        logging.info(f"API: Processing task finished for {base_filename}. Success: {results['success']}")


        if results["success"]:
             # Return paths relative to the API output directory
             relative_paths = {
                 key: pathlib.Path(path).name # Return only the filename
                 for key, path in results["files"].items()
             }
             logging.info(f"API: Returning success for {base_filename}. Files: {relative_paths}")
             return {
                 "message": "Processing successful",
                 "output_files": relative_paths, # Dictionary of {type: filename}
                 "base_filename": base_filename, # Useful for constructing download URLs
             }
        else:
            logging.error(f"API: Processing failed for {base_filename}. Error: {results.get('error')}")
            raise HTTPException(status_code=500, detail=f"Processing failed: {results.get('error', 'Unknown processing error')}")

    except HTTPException as http_exc:
         # Don't log again, just re-raise
         raise http_exc
    except Exception as e:
        logging.error(f"API Error: An unexpected error occurred processing {file.filename if file else 'No File'}: {e}", exc_info=True)
        # Return a generic error to the client
        raise HTTPException(status_code=500, detail=f"An internal server error occurred.")
    finally:
        # Ensure the uploaded temporary file handle is closed
        if file:
             await file.close()
             logging.debug(f"API: Closed uploaded file handle for {file.filename}")
        # Clean up temporary input file IF it was created
        if temp_pdf_path and temp_pdf_path.exists():
             try:
                  temp_pdf_path.unlink()
                  logging.info(f"API: Cleaned up input file: {temp_pdf_path}")
             except OSError as e:
                  logging.error(f"API: Error deleting temp file {temp_pdf_path}: {e}")


# --- Optional: Download Endpoint ---
@app.get("/download/{filename}",
         summary="Download Generated File",
         description="Downloads a previously generated output file (Excel, PDF, or TXT). Use the filename returned by the /process_pdf/ endpoint.",
         response_description="The requested file for download.",
         )
async def download_file(filename: str):
    """ Downloads a generated file from the API output directory. """
    # Basic security: prevent path traversal
    if ".." in filename or filename.startswith("/"):
         raise HTTPException(status_code=400, detail="Invalid filename.")

    file_path = OUTPUT_DIR_API / filename # Path relative to server output dir
    logging.info(f"API: Download request for: {filename} (Path: {file_path})")

    if file_path.is_file():
        # Guess media type based on extension
        media_type = 'application/octet-stream'
        low_filename = filename.lower()
        if low_filename.endswith(".xlsx"): media_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        elif low_filename.endswith(".pdf"): media_type = 'application/pdf'
        elif low_filename.endswith(".txt"): media_type = 'text/plain; charset=utf-8' # Specify charset

        return FileResponse(path=str(file_path), media_type=media_type, filename=filename) # Pass filename for browser
    else:
        logging.warning(f"API: Download failed - File not found: {file_path}")
        raise HTTPException(status_code=404, detail=f"File not found: {filename}")

# --- Root endpoint ---
@app.get("/", include_in_schema=False) # Hide from default docs
async def read_root():
    # Redirect to docs or provide simple status
    # from fastapi.responses import RedirectResponse
    # return RedirectResponse(url='/docs')
    return {"message": "PDF Processing API is running. Access /docs for details."}

# To run: uvicorn api_main:app --reload