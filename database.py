# database.py
from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import os

# Example for SQLite (simple local file)
# DATABASE_URL = "sqlite:///./pdf_jobs.db"
# engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False}) # Needed for SQLite + FastAPI

# Example for PostgreSQL (replace with your credentials/host)
DB_USER = os.getenv("DB_USER", "user")
DB_PASSWORD = os.getenv("DB_PASSWORD", "password")
DB_HOST = os.getenv("DB_HOST", "localhost")
DB_PORT = os.getenv("DB_PORT", "5432")
DB_NAME = os.getenv("DB_NAME", "pdf_processor_db")
DATABASE_URL = f"postgresql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
engine = create_engine(DATABASE_URL)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# Dependency to get DB session in API routes
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()