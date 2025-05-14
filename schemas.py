# schemas.py
from pydantic import BaseModel
from datetime import datetime
from typing import Optional, Dict
import models # Import your SQLAlchemy models to reference the Enum

class OutputFileBase(BaseModel):
    file_type: models.OutputFileType # Use the Enum from models
    filename: str
    file_path: Optional[str] = None # Or remove if only returning filename

class OutputFileSchema(OutputFileBase):
    id: int
    job_id: int
    generated_time: datetime

    class Config:
        orm_mode = True # Enable SQLAlchemy object reading

# Response model for the status endpoint
class JobStatusResponse(BaseModel):
    job_id: str
    status: models.JobStatus # Use the Enum from models
    original_filename: Optional[str] = None
    upload_time: datetime
    error_message: Optional[str] = None
    output_files: Dict[str, str] = {} # e.g., {"excel": "filename1.xlsx", "txt": "filename2.txt"}

    class Config:
        orm_mode = True
        use_enum_values = True # Important for serializing Enum members as strings