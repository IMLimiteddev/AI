# models.py
from sqlalchemy import Column, Integer, String, DateTime, Enum as SQLEnum, ForeignKey, LargeBinary
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
import enum
from database import Base

class JobStatus(str, enum.Enum):
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"

class UploadJob(Base):
    __tablename__ = "upload_jobs"

    id = Column(Integer, primary_key=True, index=True)
    job_id = Column(String, unique=True, index=True, nullable=False) # Public facing Job ID (e.g., UUID)
    original_filename = Column(String, index=True)
    upload_time = Column(DateTime(timezone=True), server_default=func.now())
    status = Column(SQLEnum(JobStatus), default=JobStatus.PENDING)
    error_message = Column(String, nullable=True)
    # Option 1: Store file path
    input_file_path = Column(String, nullable=True)
    # Option 2: Store file content directly (use LargeBinary/BLOB type appropriate for your DB)
    # input_file_content = Column(LargeBinary, nullable=True)

    output_files = relationship("OutputFile", back_populates="job")

class OutputFileType(str, enum.Enum):
    EXCEL = "excel"
    PDF = "pdf"
    TXT = "txt"

class OutputFile(Base):
    __tablename__ = "output_files"

    id = Column(Integer, primary_key=True, index=True)
    job_id = Column(Integer, ForeignKey("upload_jobs.id")) # Link to the job
    file_type = Column(SQLEnum(OutputFileType))
    filename = Column(String, index=True)
    generated_time = Column(DateTime(timezone=True), server_default=func.now())
    # Option 1: Store file path
    file_path = Column(String, nullable=True)
    # Option 2: Store file content directly
    # file_content = Column(LargeBinary, nullable=True)

    job = relationship("UploadJob", back_populates="output_files")