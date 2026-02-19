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


# ─── Table creation ──────────────────────────────────────────────────────────

def create_tables():
    """Create all tables in the SQLite database if they don't already exist."""
    Base.metadata.create_all(bind=engine)


if __name__ == "__main__":
    create_tables()
    print(f"Database created at {DB_PATH}")
