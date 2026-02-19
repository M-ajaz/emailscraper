"""
Database models for Recruitment Matching.

SQLAlchemy + SQLite backend storing candidates, job requisitions,
and match results.  DB file: backend/recruitment.db
"""

from datetime import datetime, timezone
from pathlib import Path

from sqlalchemy import (
    Column, Integer, Float, String, Text, Boolean, DateTime,
    ForeignKey, create_engine,
)
from sqlalchemy.orm import declarative_base, relationship, sessionmaker

DB_PATH = Path(__file__).resolve().parent / "recruitment.db"
DATABASE_URL = f"sqlite:///{DB_PATH}"

engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()


def _utcnow():
    return datetime.now(timezone.utc)


# ─── Models ──────────────────────────────────────────────────────────────────

class Candidate(Base):
    __tablename__ = "candidates"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String(255), nullable=False)
    email = Column(String(255), nullable=True)
    phone = Column(String(50), nullable=True)
    location = Column(String(255), nullable=True)
    titles = Column(Text, nullable=True)          # JSON array of strings
    skills = Column(Text, nullable=True)          # JSON array of strings
    years_exp = Column(Float, nullable=True)
    raw_resume_path = Column(String(500), nullable=True)
    source_email_uid = Column(String(100), nullable=True)
    notes = Column(Text, nullable=True)
    tags = Column(Text, nullable=True, default="[]")   # JSON array of strings
    created_at = Column(DateTime, default=_utcnow, nullable=False)

    matches = relationship("MatchResult", back_populates="candidate")


class JobRequisition(Base):
    __tablename__ = "job_requisitions"

    id = Column(Integer, primary_key=True, index=True)
    title = Column(String(255), nullable=False)
    required_skills = Column(Text, nullable=True)  # JSON array of strings
    min_exp = Column(Float, nullable=True)
    location = Column(String(255), nullable=True)
    remote_ok = Column(Boolean, default=False)
    jd_raw = Column(Text, nullable=True)
    created_at = Column(DateTime, default=_utcnow, nullable=False)

    matches = relationship("MatchResult", back_populates="job")


class MatchResult(Base):
    __tablename__ = "match_results"

    id = Column(Integer, primary_key=True, index=True)
    job_id = Column(Integer, ForeignKey("job_requisitions.id"), nullable=False)
    candidate_id = Column(Integer, ForeignKey("candidates.id"), nullable=False)
    score = Column(Float, nullable=False, default=0.0)
    match_reasons = Column(Text, nullable=True)    # JSON array of strings
    fit_level = Column(Text, nullable=True)        # e.g. "strong", "moderate", "weak"

    job = relationship("JobRequisition", back_populates="matches")
    candidate = relationship("Candidate", back_populates="matches")


class ScrapedEmail(Base):
    __tablename__ = "scraped_emails"

    id = Column(Integer, primary_key=True, index=True)
    uid = Column(Text, unique=True, nullable=False, index=True)
    folder = Column(Text, nullable=True)
    subject = Column(Text, nullable=True)
    sender = Column(Text, nullable=True)
    sender_email = Column(Text, nullable=True)
    date = Column(Text, nullable=True)
    body_text = Column(Text, nullable=True)
    body_html = Column(Text, nullable=True)
    has_attachments = Column(Boolean, default=False)
    attachment_count = Column(Integer, default=0)
    is_read = Column(Boolean, default=False)
    scraped_at = Column(DateTime, default=_utcnow, nullable=False)


class Attachment(Base):
    __tablename__ = "attachments"

    id = Column(Integer, primary_key=True, index=True)
    filename = Column(Text, unique=True, nullable=False, index=True)
    original_name = Column(Text, nullable=True)
    content_type = Column(Text, nullable=True)
    size = Column(Integer, default=0)
    email_uid = Column(Text, nullable=True)
    email_subject = Column(Text, nullable=True)
    email_sender = Column(Text, nullable=True)
    email_date = Column(Text, nullable=True)
    saved_at = Column(DateTime, default=_utcnow, nullable=False)


class SchedulerConfig(Base):
    __tablename__ = "scheduler_config"

    id = Column(Integer, primary_key=True, index=True)
    enabled = Column(Boolean, default=False)
    interval_minutes = Column(Integer, default=30)
    folder = Column(Text, default="INBOX")
    subject_filter = Column(Text, nullable=True)
    last_run_at = Column(DateTime, nullable=True)
    next_run_at = Column(DateTime, nullable=True)
    emails_found_last_run = Column(Integer, default=0)
    candidates_added_last_run = Column(Integer, default=0)


class Notification(Base):
    __tablename__ = "notifications"

    id = Column(Integer, primary_key=True, index=True)
    type = Column(Text, nullable=False)           # "new_candidate", "new_high_fit", "scrape_complete"
    title = Column(Text, nullable=False)
    message = Column(Text, nullable=True)
    job_id = Column(Integer, nullable=True)
    candidate_id = Column(Integer, nullable=True)
    is_read = Column(Boolean, default=False)
    created_at = Column(DateTime, default=_utcnow, nullable=False)


# ─── Table creation ──────────────────────────────────────────────────────────

def create_tables():
    """Create all tables in the SQLite database if they don't already exist."""
    Base.metadata.create_all(bind=engine)
    _migrate_tables()


def _migrate_tables():
    """Add columns that may be missing from older schemas."""
    import sqlite3
    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()
    # Candidates table migrations
    cursor.execute("PRAGMA table_info(candidates)")
    existing = {row[1] for row in cursor.fetchall()}
    if "notes" not in existing:
        cursor.execute("ALTER TABLE candidates ADD COLUMN notes TEXT")
    if "tags" not in existing:
        cursor.execute("ALTER TABLE candidates ADD COLUMN tags TEXT DEFAULT '[]'")
    # SchedulerConfig table — ensure it exists (create_all handles this,
    # but we seed a default row if the table is empty)
    cursor.execute("SELECT COUNT(*) FROM scheduler_config")
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            "INSERT INTO scheduler_config (enabled, interval_minutes, folder) "
            "VALUES (0, 30, 'INBOX')"
        )
    conn.commit()
    conn.close()


if __name__ == "__main__":
    create_tables()
    print(f"Database created at {DB_PATH}")
