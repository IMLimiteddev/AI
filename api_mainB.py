# api_main.py
import fastapi
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks, Depends
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import pathlib
import logging
import uuid
import os
import datetime
import re
from sqlalchemy.orm import Session # Import Session

# --- Your processing imports ---
import pdf_parser
import data_mapper
import excel_writer
import pdf_writer
import text_writer
from config import BASE_DIR

# --- Database Imports ---
import models # Import your models
import schemas # Import your Pydantic schemas (create schemas.py based on models)
from database import SessionLocal, engine, get_db # Import SessionLocal, engine, get_db

# --- Create DB Tables (only needed once, or use Alembic) ---
# Comment out after first run or use Alembic for migrations
# models.Base.metadata.create_all(bind=engine)
# logging.info("Database tables checked/created.")
# ---

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s')

TEMP_DIR = BASE_DIR / "temp_files"; TEMP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR_API = BASE_DIR / "api_output"; OUTPUT_DIR_API.mkdir(exist_ok=True)
INPUT_STORAGE_DIR = BASE_DIR / "input_storage"; INPUT_STORAGE_DIR.mkdir(exist_ok=True) # Store input PDFs persistently

app = FastAPI(title="PDF Processing API", description="Processes PDFs, stores results in DB.")

# --- CORS ---
origins = ["http://localhost:3000", "http://127.0.0.1:3000", "http://localhost", "http://127.0.0.1"] # Add others
app.add_middleware(CORSMiddleware, allow_origins=origins, allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# =============================================================================
# Background Task - Processing Logic (Modified)
# =============================================================================
def run_processing_task(db: Session, job_id: str, input_pdf_path: pathlib.Path, original_filename: str):
    """ Runs PDF processing and updates database. """
    output_dir = OUTPUT_DIR_API
    # Generate unique base filename for outputs based on job_id
    base_filename = f"{datetime.datetime.now():%Y%m%d_%H%M%S}_{job_id}"
    job = None # Initialize job variable

    try:
        # --- Get Job from DB ---
        job = db.query(models.UploadJob).filter(models.UploadJob.job_id == job_id).first()
        if not job:
            logging.error(f"Background Task: Job {job_id} not found in DB.")
            return # Cannot proceed

        # --- Update Job Status to Processing ---
        job.status = models.JobStatus.PROCESSING
        db.commit()
        logging.info(f"Background Task: Started processing job {job_id} for {original_filename}")

        # 1. Parse PDF
        extracted_data = pdf_parser.extract_data_from_pdf(input_pdf_path)
        if not extracted_data: raise ValueError("Failed to extract data from PDF.")

        # 2. Map Data
        mapped_data = data_mapper.map_data_to_template(extracted_data)
        if not mapped_data: raise ValueError("Failed to map extracted data.")

        output_file_records = [] # To store DB records for output files

        # 3. Write Excel & Record
        excel_path_str = excel_writer.write_to_excel(mapped_data, output_dir, base_filename)
        if excel_path_str:
            excel_filename = pathlib.Path(excel_path_str).name
            output_file_records.append(models.OutputFile(
                job_id=job.id, file_type=models.OutputFileType.EXCEL,
                filename=excel_filename, file_path=str(pathlib.Path(excel_path_str).relative_to(BASE_DIR)) # Store relative path maybe
            ))
            logging.info(f"Generated Excel: {excel_filename}")
        else: logging.warning("Failed to generate Excel file.")

        # 4. Write Combined PDF & Record
        pdf_path_str = pdf_writer.write_combined_pdf(mapped_data, output_dir, base_filename)
        if pdf_path_str:
            pdf_filename = pathlib.Path(pdf_path_str).name
            output_file_records.append(models.OutputFile(
                job_id=job.id, file_type=models.OutputFileType.PDF,
                filename=pdf_filename, file_path=str(pathlib.Path(pdf_path_str).relative_to(BASE_DIR))
            ))
            logging.info(f"Generated PDF: {pdf_filename}")
        else: logging.warning("Failed to generate Combined PDF.")

        # 5. Write TXT & Record
        txt_path_str = text_writer.write_auftrag_export_txt(mapped_data, output_dir, base_filename)
        if txt_path_str:
            txt_filename = pathlib.Path(txt_path_str).name
            output_file_records.append(models.OutputFile(
                job_id=job.id, file_type=models.OutputFileType.TXT,
                filename=txt_filename, file_path=str(pathlib.Path(txt_path_str).relative_to(BASE_DIR))
            ))
            logging.info(f"Generated TXT: {txt_filename}")
        else: logging.warning("Failed to generate TXT file.") # Maybe raise error if TXT is mandatory

        # --- Update Job Status to Completed ---
        job.status = models.JobStatus.COMPLETED
        # Add output file records to the session
        db.add_all(output_file_records)
        db.commit()
        logging.info(f"Background Task: Successfully completed job {job_id}")

    except Exception as e:
        logging.error(f"Background Task: Error processing job {job_id}: {e}", exc_info=True)
        if job: # Update status to FAILED if job object exists
            job.status = models.JobStatus.FAILED
            job.error_message = str(e)[:255] # Truncate error message if needed
            db.commit()
    finally:
        # Optional: Clean up the temporary input PDF (or keep it for debugging)
        # if input_pdf_path.exists():
        #     try: input_pdf_path.unlink()
        #     except OSError: logging.error(f"Error deleting temp file {input_pdf_path}")
        pass

# =============================================================================
# API Endpoints
# =============================================================================

@app.post("/process_pdf/",
          summary="Upload PDF for Processing",
          status_code=fastapi.status.HTTP_202_ACCEPTED) # Return 202 Accepted for background tasks
async def process_pdf_endpoint(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(..., description="The PDF file to process."),
    db: Session = Depends(get_db) # Inject DB session
    ) -> dict:
    """ Accepts PDF upload, stores metadata, and starts background processing. """
    if not file or not file.filename: raise HTTPException(status_code=400, detail="No file provided.")
    if not file.filename.lower().endswith(".pdf"): raise HTTPException(status_code=400, detail="Invalid file type.")

    temp_pdf_path = None
    persisted_input_path = None
    job_id = str(uuid.uuid4()) # Generate unique ID for this job

    try:
        # 1. Save Uploaded File Persistently (optional but good practice)
        safe_filename = re.sub(r'[^\w\._-]', '_', file.filename)
        persisted_input_path = INPUT_STORAGE_DIR / f"{job_id}_{safe_filename}"
        with persisted_input_path.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        logging.info(f"API: Saved input PDF persistently to {persisted_input_path}")

        # 2. Create Job Record in DB
        new_job = models.UploadJob(
            job_id=job_id,
            original_filename=file.filename,
            input_file_path=str(persisted_input_path.relative_to(BASE_DIR)), # Store relative path
            status=models.JobStatus.PENDING
        )
        db.add(new_job)
        db.commit()
        db.refresh(new_job) # Get the auto-generated ID, status, etc.
        logging.info(f"API: Created Job {job_id} (DB ID: {new_job.id}) for {file.filename}")

        # 3. Add Processing Task to Background
        # Pass DB session factory or handle session scoping carefully in task
        # Simpler approach: pass necessary data (job_id, path)
        # Complex: Pass SessionLocal and create session inside task (requires careful handling)
        # For simplicity here, we'll query the job object again inside the task
        # NOTE: Need to ensure the DB session used by the background task is managed correctly.
        # Creating a new session inside the task is often safer.
        background_tasks.add_task(run_processing_task, SessionLocal(), job_id, persisted_input_path, file.filename)
        logging.info(f"API: Added job {job_id} to background tasks.")

        # 4. Return Job ID to Client Immediately
        return {"message": "File upload accepted, processing started.", "job_id": job_id}

    except Exception as e:
        logging.error(f"API Error during upload/job creation for {file.filename}: {e}", exc_info=True)
        # Clean up persisted file if job creation failed
        if persisted_input_path and persisted_input_path.exists():
             try: persisted_input_path.unlink()
             except OSError: pass
        raise HTTPException(status_code=500, detail=f"Failed to start processing job: {e}")
    finally:
        if file: await file.close()


@app.get("/job_status/{job_id}",
         summary="Get Job Status and Results",
         response_model=schemas.JobStatusResponse) # Use a Pydantic model for response structure (define in schemas.py)
async def get_job_status(job_id: str, db: Session = Depends(get_db)):
    """ Poll this endpoint to check job status and get output filenames when completed. """
    job = db.query(models.UploadJob).filter(models.UploadJob.job_id == job_id).first()
    if not job:
        raise HTTPException(status_code=404, detail="Job not found")

    output_files_dict = {}
    if job.status == models.JobStatus.COMPLETED:
        outputs = db.query(models.OutputFile).filter(models.OutputFile.job_id == job.id).all()
        for output in outputs:
            output_files_dict[output.file_type.value] = output.filename # e.g., {"excel": "...", "pdf": ...}

    # You'll need to create schemas.py with Pydantic models like JobStatusResponse
    return schemas.JobStatusResponse(
        job_id=job.job_id,
        status=job.status,
        original_filename=job.original_filename,
        upload_time=job.upload_time,
        error_message=job.error_message,
        output_files=output_files_dict
    )

@app.get("/download/{filename}", summary="Download Generated File")
async def download_file(filename: str):
    """ Downloads a generated file from the API output directory. """
    if ".." in filename or "/" in filename or "\\" in filename or "\0" in filename:
         raise HTTPException(status_code=400, detail="Invalid filename.")

    # Files are expected in OUTPUT_DIR_API
    file_path = (OUTPUT_DIR_API / filename).resolve()

    # Security check: Ensure path didn't escape the output directory
    if not str(file_path).startswith(str(OUTPUT_DIR_API.resolve())):
        raise HTTPException(status_code=403, detail="Forbidden.")

    logging.info(f"API: Download request for: {filename} (Path: {file_path})")
    if file_path.is_file():
        media_type = 'application/octet-stream'
        low_filename = filename.lower()
        if low_filename.endswith(".xlsx"): media_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        elif low_filename.endswith(".pdf"): media_type = 'application/pdf'
        elif low_filename.endswith(".txt"): media_type = 'text/plain; charset=utf-8'
        return FileResponse(path=str(file_path), media_type=media_type, filename=filename)
    else:
        logging.warning(f"API: Download failed - File not found: {file_path}")
        raise HTTPException(status_code=404, detail=f"File not found: {filename}")

@app.get("/", include_in_schema=False)
async def read_root():
    from fastapi.responses import RedirectResponse
    return RedirectResponse(url='/docs') # Redirect base URL to API docs