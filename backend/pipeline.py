"""
Candidate ingestion pipeline.

Takes an attachment file + email context, runs the full extraction chain,
deduplicates, and persists the candidate to SQLite.
"""

import json
import logging
from pathlib import Path

from database import SessionLocal, Candidate, create_tables
from parsers import extract_text_from_pdf, extract_text_from_docx, extract_entities
from extractors import extract_email_metadata, merge_profile, find_existing_candidate

logger = logging.getLogger(__name__)

_SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".doc"}


def process_attachment_into_candidate(
    attachment_filepath: str,
    email_uid: str,
    email_body: str,
    email_sender: str,
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
