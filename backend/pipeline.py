"""
Candidate ingestion pipeline.

Takes an attachment file + email context, runs the full extraction chain,
deduplicates, and persists the candidate to SQLite.
"""

import json
import logging
from pathlib import Path

from database import SessionLocal, Candidate, JobRequisition, Notification, create_tables
from parsers import (
    extract_text_from_pdf, extract_text_from_docx, extract_entities,
    extract_name_from_subject, extract_title_from_subject,
)
from extractors import extract_email_metadata, merge_profile, find_existing_candidate

logger = logging.getLogger(__name__)

_SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".doc"}


def process_attachment_into_candidate(
    attachment_filepath: str,
    email_uid: str,
    email_body: str,
    email_sender: str,
    email_subject: str = "",
) -> dict | None:
    """
    End-to-end pipeline: attachment file → parsed candidate → DB row.

    Parameters
    ----------
    attachment_filepath : str
        Path to the saved resume/CV file (PDF or DOCX).
    email_uid : str
        IMAP UID (or encoded email ID) for traceability.
    email_body : str
        Plain-text body of the email that carried the attachment.
    email_sender : str
        Sender email address from the email headers.
    email_subject : str
        Subject line of the email (used to extract candidate name / title
        from forwarded recruitment patterns).

    Returns
    -------
    dict  – the saved candidate record (with ``id``), or
    None  – if the candidate is a duplicate or the file is unsupported.
    """
    create_tables()

    filepath = Path(attachment_filepath)

    # ── 1. Validate file ─────────────────────────────────────────────────
    ext = filepath.suffix.lower()
    if ext not in _SUPPORTED_EXTENSIONS:
        logger.info("Skipping unsupported file type: %s", ext)
        return None

    if not filepath.exists():
        logger.error("Attachment file not found: %s", filepath)
        return None

    # ── 2. Extract text from document ────────────────────────────────────
    raw_text = ""
    try:
        if ext == ".pdf":
            raw_text = extract_text_from_pdf(filepath)
        elif ext in (".docx", ".doc"):
            raw_text = extract_text_from_docx(filepath)
    except Exception:
        logger.exception("Failed to extract text from %s", filepath)

    if not raw_text.strip():
        logger.warning("No text extracted from %s", filepath)

    # ── 3. Extract entities from resume text ─────────────────────────────
    try:
        resume_data = extract_entities(raw_text)
    except Exception:
        logger.exception("Entity extraction failed for %s", filepath)
        resume_data = {"skills": [], "years_exp": None, "titles": [], "locations": []}

    # Attach file-level metadata the merge step expects
    resume_data["raw_resume_path"] = str(filepath)
    resume_data["source_email_uid"] = email_uid

    # ── 4. Extract metadata from the email body ──────────────────────────
    try:
        email_data = extract_email_metadata(email_body or "")
    except Exception:
        logger.exception("Email metadata extraction failed")
        email_data = {
            "name": None, "email": None, "phone": None,
            "location": None, "role_applied": None,
        }

    # Use header sender address as fallback email
    if not email_data.get("email") and email_sender:
        email_data["email"] = email_sender.strip().lower()

    # ── 4b. Extract name and title from email subject line ─────────────
    if email_subject:
        subject_name = extract_name_from_subject(email_subject)
        if subject_name and not resume_data.get("name") and not email_data.get("name"):
            email_data["name"] = subject_name

        subject_title = extract_title_from_subject(email_subject)
        if subject_title:
            existing_titles = [t.lower() for t in resume_data.get("titles", [])]
            if subject_title.lower() not in existing_titles:
                resume_data.setdefault("titles", []).append(subject_title)

    # ── 5. Merge resume + email into a unified profile ───────────────────
    try:
        profile = merge_profile(resume_data, email_data)
    except Exception:
        logger.exception("Profile merge failed")
        return None

    # ── 6. Deduplicate by email address ──────────────────────────────────
    candidate_email = profile.get("email")
    if candidate_email:
        try:
            existing = find_existing_candidate(candidate_email)
            if existing:
                logger.info(
                    "Duplicate candidate (id=%s, email=%s) — skipping",
                    existing.id, candidate_email,
                )
                return None
        except Exception:
            logger.exception("Dedup lookup failed for %s", candidate_email)

    # ── 7. Persist to database ───────────────────────────────────────────
    # Candidate.name is NOT NULL — derive a fallback if parsing found nothing
    name = profile.get("name") or _name_from_email(candidate_email) or "Unknown"

    candidate = Candidate(
        name=name,
        email=candidate_email,
        phone=profile.get("phone"),
        location=profile.get("location"),
        titles=json.dumps(profile.get("titles", [])),
        skills=json.dumps(profile.get("skills", [])),
        years_exp=profile.get("years_exp"),
        raw_resume_path=profile.get("raw_resume_path"),
        source_email_uid=profile.get("source_email_uid"),
    )

    db = SessionLocal()
    try:
        db.add(candidate)
        db.commit()
        db.refresh(candidate)
        logger.info("Saved candidate id=%s name=%s", candidate.id, candidate.name)

        # ── Notifications ─────────────────────────────────────────────
        # Always create a "new_candidate" notification
        try:
            db.add(Notification(
                type="new_candidate",
                title="New Candidate",
                message=f"{candidate.name} added from email",
                candidate_id=candidate.id,
            ))

            # Quick skill match against all jobs
            cand_skills = {s.lower() for s in json.loads(candidate.skills or "[]")}
            if cand_skills:
                jobs = db.query(JobRequisition).all()
                for job in jobs:
                    job_skills = {s.lower() for s in json.loads(job.required_skills or "[]")}
                    if not job_skills:
                        continue
                    overlap = len(cand_skills & job_skills)
                    score = round((overlap / len(job_skills)) * 100)
                    if score >= 75:
                        db.add(Notification(
                            type="new_high_fit",
                            title="High-Fit Match",
                            message=f"{candidate.name} matches {job.title} ({score}%)",
                            candidate_id=candidate.id,
                            job_id=job.id,
                        ))

            db.commit()
        except Exception:
            logger.exception("Failed to create notifications for candidate %s", candidate.id)

        return {
            "id": candidate.id,
            "name": candidate.name,
            "email": candidate.email,
            "phone": candidate.phone,
            "location": candidate.location,
            "titles": json.loads(candidate.titles or "[]"),
            "skills": json.loads(candidate.skills or "[]"),
            "years_exp": candidate.years_exp,
            "raw_resume_path": candidate.raw_resume_path,
            "source_email_uid": candidate.source_email_uid,
            "created_at": candidate.created_at.isoformat() if candidate.created_at else None,
        }
    except Exception:
        db.rollback()
        logger.exception("Failed to save candidate to database")
        return None
    finally:
        db.close()


def _name_from_email(addr: str | None) -> str | None:
    """Derive a human-readable name from an email address as last resort."""
    if not addr or "@" not in addr:
        return None
    local = addr.split("@")[0]
    # "john.smith" → "John Smith", "jane_doe" → "Jane Doe"
    parts = local.replace("_", ".").replace("-", ".").split(".")
    return " ".join(p.capitalize() for p in parts if p)
