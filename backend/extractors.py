"""
Email metadata extraction, profile merging, and candidate deduplication.

- extract_email_metadata(email_body) – regex extraction from cover-letter / email text
- merge_profile(resume_data, email_data) – combine resume + email, resume wins ties
- find_existing_candidate(email_address) – dedupe check against candidates table
"""

import re
from typing import Optional

from database import SessionLocal, Candidate


# ─── Compiled patterns (module-level for reuse) ─────────────────────────────

_EMAIL_RE = re.compile(
    r'\b[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}\b'
)

# Phone: international and US/UK formats
# +1 (555) 123-4567, 555-123-4567, (555) 123 4567, +44 20 7946 0958, etc.
_PHONE_RE = re.compile(
    r'(?:\+\d{1,3}[\s\-.]?)?' 	         # optional country code
    r'(?:\(?\d{2,4}\)?[\s\-.]?)?'        # optional area code
    r'\d{3,4}[\s\-.]?\d{3,4}'            # main number
)

# "Applying for <role>", "Position: <role>", "Role: <role>", "interested in the <role> position"
_ROLE_PATTERNS = [
    # "interested in the Senior Engineer position", "applying for the role of Data Scientist"
    re.compile(
        r'(?:appl(?:y|ying)\s+for|application\s+for|interest(?:ed)?\s+in)'
        r'\s+(?:the\s+)?(?:position\s+of\s+|role\s+of\s+)?'
        r'(.+?)'
        r'(?:\s+(?:position|role|opening|job|vacancy)(?:\s|[.,;])|[.\n,])',
        re.IGNORECASE,
    ),
    # "Position: Senior Engineer" / "Role: Data Analyst"
    re.compile(
        r'(?:position|role|job\s+title)\s*:\s*(.+?)(?:\s*[.\n,|]|$)',
        re.IGNORECASE,
    ),
    # Subject-line style: "Application — Senior Engineer"
    re.compile(
        r'(?:application|apply|candidate)\s*[\-–—:]\s*(.+?)(?:\s*[.\n,|]|$)',
        re.IGNORECASE,
    ),
]

# Name from greeting / sign-off lines
_NAME_PATTERNS = [
    # "Dear Hiring Manager, my name is John Smith"
    re.compile(
        r'(?:my\s+name\s+is|i\s+am|this\s+is)\s+'
        r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})',
    ),
    # Sign-off: "Regards,\nJohn Smith" or "Best,\n  John Smith"
    re.compile(
        r'(?:regards|sincerely|best|thanks|cheers|respectfully'
        r'|thank\s+you|warm\s+regards|kind\s+regards)\s*,?\s*\n\s*'
        r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})',
    ),
    # "From: John Smith" header remnant in forwarded mail
    re.compile(
        r'(?:^|\n)\s*(?:from|name)\s*:\s*'
        r'([A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,3})',
    ),
]

# "Based in <location>", "Location: <location>", "from <City>, <ST>"
_LOCATION_PATTERNS = [
    re.compile(
        r'(?:based\s+in|located\s+in|location\s*:\s*|residing\s+in|from)\s+'
        r'([A-Z][A-Za-z\s.]+(?:,\s*[A-Z]{2})?)',
        re.IGNORECASE,
    ),
    # "City, ST" or "City, State" standalone
    re.compile(
        r'\b([A-Z][a-z]+(?:\s[A-Z][a-z]+)?,\s*[A-Z]{2})\b',
    ),
]


# ─── Email metadata extraction ──────────────────────────────────────────────

def extract_email_metadata(email_body: str) -> dict:
    """
    Extract candidate metadata from an email / cover-letter body using regex.

    Returns
    -------
    dict with keys:
        name          : str | None
        email         : str | None
        phone         : str | None
        location      : str | None
        role_applied  : str | None
    """
    result: dict = {
        "name": None,
        "email": None,
        "phone": None,
        "location": None,
        "role_applied": None,
    }

    if not email_body:
        return result

    # --- Email address ---
    email_match = _EMAIL_RE.search(email_body)
    if email_match:
        result["email"] = email_match.group(0).lower()

    # --- Phone number ---
    phone_match = _PHONE_RE.search(email_body)
    if phone_match:
        raw = phone_match.group(0).strip()
        # Only accept if it has enough digits to be a real phone number
        digits = re.sub(r'\D', '', raw)
        if len(digits) >= 7:
            result["phone"] = raw

    # --- Name ---
    for pat in _NAME_PATTERNS:
        m = pat.search(email_body)
        if m:
            name = m.group(1).strip()
            # Reject if it looks like a sentence fragment (too many words)
            if 1 <= len(name.split()) <= 4:
                result["name"] = name
                break

    # --- Role applied ---
    for pat in _ROLE_PATTERNS:
        m = pat.search(email_body)
        if m:
            role = m.group(1).strip()
            # Clean trailing noise
            role = re.sub(r'\s+(?:at|with|for)\s+.*$', '', role, flags=re.IGNORECASE)
            if 2 <= len(role) <= 80:
                result["role_applied"] = role
                break

    # --- Location ---
    for pat in _LOCATION_PATTERNS:
        m = pat.search(email_body)
        if m:
            loc = m.group(1).strip().rstrip('.,;')
            # Reject overly long matches (sentence fragments)
            if 2 <= len(loc) <= 60:
                result["location"] = loc
                break

    return result


# ─── Profile merging ────────────────────────────────────────────────────────

def merge_profile(resume_data: dict, email_data: dict) -> dict:
    """
    Merge resume-parsed data with email-parsed metadata into a single
    candidate profile.  Resume data wins where both sources have a value;
    email data fills gaps.

    Parameters
    ----------
    resume_data : dict
        Output from ``parsers.extract_entities()`` plus optional extra keys
        (name, email, phone, raw_resume_path, source_email_uid).
    email_data : dict
        Output from ``extract_email_metadata()``.

    Returns
    -------
    dict with keys matching the Candidate model:
        name, email, phone, location, titles, skills,
        years_exp, raw_resume_path, source_email_uid
    """
    def _pick(key: str, *sources):
        """Return the first non-empty value across sources for *key*."""
        for src in sources:
            val = src.get(key)
            if val is not None and val != "" and val != []:
                return val
        return None

    # Locations: resume gives a list, email gives a single string.
    # Prefer resume list, fall back to email scalar.
    resume_locations = resume_data.get("locations", [])
    email_location = email_data.get("location")
    if resume_locations:
        location = ", ".join(resume_locations)
    elif email_location:
        location = email_location
    else:
        location = None

    # Role from email can supplement titles from resume
    titles = list(resume_data.get("titles", []))
    role_applied = email_data.get("role_applied")
    if role_applied and role_applied.lower() not in {t.lower() for t in titles}:
        titles.append(role_applied)

    return {
        "name": _pick("name", resume_data, email_data),
        "email": _pick("email", resume_data, email_data),
        "phone": _pick("phone", resume_data, email_data),
        "location": location,
        "titles": titles,
        "skills": resume_data.get("skills", []),
        "years_exp": resume_data.get("years_exp"),
        "raw_resume_path": resume_data.get("raw_resume_path"),
        "source_email_uid": resume_data.get("source_email_uid"),
    }


# ─── Deduplication ──────────────────────────────────────────────────────────

def find_existing_candidate(email_address: str) -> Optional[Candidate]:
    """
    Look up a candidate by email address (case-insensitive).

    Returns the existing ``Candidate`` row or ``None``.
    """
    if not email_address:
        return None
    db = SessionLocal()
    try:
        return (
            db.query(Candidate)
            .filter(Candidate.email.ilike(email_address.strip()))
            .first()
        )
    finally:
        db.close()
