"""
Resume / document parsing utilities.

- extract_text_from_pdf(filepath)  – via pdfplumber
- extract_text_from_docx(filepath) – via python-docx
- extract_entities(raw_text)       – regex + keyword matching (no external APIs)
- extract_name_from_subject(subject) – parse candidate name from forwarded email subjects
"""

import re
from pathlib import Path

import pdfplumber
from docx import Document as DocxDocument


# ─── Text extraction ────────────────────────────────────────────────────────

def extract_text_from_pdf(filepath: str | Path) -> str:
    """Extract all text from a PDF file using pdfplumber."""
    pages = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                pages.append(text)
    return "\n".join(pages)


def extract_text_from_docx(filepath: str | Path) -> str:
    """Extract all text from a DOCX file using python-docx."""
    doc = DocxDocument(str(filepath))
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(paragraphs)


# ─── Skill keywords (lowercase) ─────────────────────────────────────────────
# Grouped by domain for readability; flattened into a set at module load.

_SKILL_GROUPS = {
    "languages": [
        "python", "java", "javascript", "typescript", "c#", "c++", "go",
        "golang", "rust", "ruby", "php", "swift", "kotlin", "scala",
        "perl", "r", "matlab", "dart", "lua", "shell", "bash",
        "powershell", "sql", "html", "css", "sass", "less",
        "embedded c", "vhdl", "verilog",
    ],
    "frontend": [
        "react", "reactjs", "react.js", "angular", "angularjs", "vue",
        "vuejs", "vue.js", "svelte", "nextjs", "next.js", "nuxt",
        "nuxtjs", "gatsby", "remix", "tailwind", "tailwindcss",
        "bootstrap", "jquery", "webpack", "vite", "redux", "mobx",
        "graphql", "apollo",
    ],
    "backend": [
        "node", "nodejs", "node.js", "express", "expressjs", "fastapi",
        "flask", "django", "spring", "spring boot", "springboot",
        "rails", "laravel", "asp.net", ".net", "dotnet", "nestjs",
        "fastify", "gin", "fiber", "actix",
    ],
    "data_ml": [
        "pandas", "numpy", "scipy", "scikit-learn", "sklearn",
        "tensorflow", "pytorch", "keras", "opencv", "spark",
        "pyspark", "hadoop", "hive", "airflow", "kafka",
        "machine learning", "deep learning", "nlp",
        "natural language processing", "computer vision",
        "data science", "data engineering", "data analysis",
        "power bi", "powerbi", "tableau", "looker", "dbt",
        "etl", "data pipeline", "simulink", "signal processing",
    ],
    "cloud_devops": [
        "aws", "azure", "gcp", "google cloud", "heroku",
        "digitalocean", "docker", "kubernetes", "k8s", "terraform",
        "ansible", "jenkins", "github actions", "gitlab ci", "ci/cd",
        "cicd", "linux", "nginx", "apache", "cloudflare",
        "serverless", "lambda", "ec2", "s3", "ecs", "eks",
        "fargate", "cloudformation",
    ],
    "databases": [
        "mysql", "postgresql", "postgres", "mongodb", "redis",
        "elasticsearch", "sqlite", "oracle", "sql server",
        "dynamodb", "cassandra", "couchdb", "neo4j", "mariadb",
        "firebase", "firestore", "supabase",
    ],
    "tools_practices": [
        "git", "github", "gitlab", "bitbucket", "jira",
        "confluence", "agile", "scrum", "kanban", "tdd",
        "rest", "restful", "soap", "microservices",
        "api", "oauth", "jwt", "websocket", "grpc",
        "rabbitmq", "celery", "selenium", "cypress",
        "playwright", "jest", "pytest", "unittest", "mocha",
    ],
    "mobile": [
        "android", "ios", "react native", "flutter", "xamarin",
        "swiftui", "objective-c", "cordova", "ionic",
    ],
    "electrical_hardware": [
        "autocad", "plc", "altium", "labview", "solidworks",
        "fpga", "pcb", "rf", "power electronics",
    ],
}

# Flatten to a single set; sort longest-first so multi-word skills match before
# their single-word substrings (e.g. "spring boot" before "spring").
SKILLS = sorted(
    {s for group in _SKILL_GROUPS.values() for s in group},
    key=len,
    reverse=True,
)

# Pre-compile a regex for each skill.  Word boundaries work for most terms;
# special-case entries that contain punctuation (e.g. "c++", "c#", "node.js").
_SKILL_PATTERNS: list[tuple[str, re.Pattern]] = []
for _skill in SKILLS:
    if re.search(r'[+#./]', _skill):
        # Escape the literal and anchor with lookaround so "c++" doesn't need
        # a trailing word-boundary (which wouldn't match after '+').
        _pat = re.compile(r'(?<![a-zA-Z])' + re.escape(_skill) + r'(?![a-zA-Z])', re.IGNORECASE)
    else:
        _pat = re.compile(r'\b' + re.escape(_skill) + r'\b', re.IGNORECASE)
    _SKILL_PATTERNS.append((_skill, _pat))


# ─── Title keywords ─────────────────────────────────────────────────────────

_TITLE_PATTERNS: list[re.Pattern] = [
    re.compile(p, re.IGNORECASE) for p in [
        # "Senior Software Engineer", "Lead Data Scientist", etc.
        r'\b(?:senior|junior|jr\.?|sr\.?|lead|principal|staff|chief|head of|vp of|director of|manager of|associate|intern)?'
        r'\s*'
        r'(?:'
            r'software engineer(?:ing)?|software developer|web developer|'
            r'full[\s-]?stack (?:developer|engineer)|'
            r'front[\s-]?end (?:developer|engineer)|'
            r'back[\s-]?end (?:developer|engineer)|'
            r'devops engineer|sre|site reliability engineer|'
            r'cloud (?:engineer|architect)|solutions? architect|'
            r'data (?:engineer|scientist|analyst)|'
            r'machine learning engineer|ml engineer|ai engineer|'
            r'mobile (?:developer|engineer)|'
            r'ios developer|android developer|'
            r'qa engineer|test engineer|sdet|'
            r'security engineer|infosec engineer|'
            r'database administrator|dba|'
            r'systems? (?:engineer|administrator|admin)|'
            r'network engineer|'
            r'product manager|project manager|program manager|'
            r'engineering manager|technical lead|tech lead|team lead|'
            r'scrum master|'
            r'ux designer|ui designer|ux/ui designer|product designer|'
            r'business analyst|data architect|'
            r'cto|cio|vp engineering|director of engineering|'
            r'electrical engineer(?:ing)?|'
            r'mechanical engineer(?:ing)?|'
            r'hardware engineer(?:ing)?|'
            r'embedded (?:systems? )?engineer|'
            r'firmware engineer|'
            r'rf engineer|'
            r'controls? engineer|'
            r'power engineer|'
            r'design engineer|'
            r'manufacturing engineer|'
            r'process engineer|'
            r'validation engineer|'
            r'test engineer|'
            r'field (?:service )?engineer|'
            r'applications? engineer'
        r')',
    ]
]


# ─── Years-of-experience patterns ───────────────────────────────────────────

_YOE_PATTERNS = [
    # "5+ years of experience", "10 yrs", "3-5 years", "over 8 years",
    # "more than 6 years", "15+ yrs of work"
    re.compile(
        r'(?:over\s+|more\s+than\s+)?'
        r'(\d{1,2})\s*[\-–to]*\s*\d{0,2}\s*\+?\s*'
        r'(?:years?|yrs?|yr)\b'
        r'(?:\s+of\s+(?:experience|exp|work|professional))?',
        re.IGNORECASE,
    ),
    # "experience: 7 years"
    re.compile(
        r'experience\s*(?::|\s)\s*(\d{1,2})\s*\+?\s*(?:years?|yrs?)',
        re.IGNORECASE,
    ),
    # "X-year career" / "X year career"
    re.compile(
        r'(\d{1,2})[\s-]*year\s+career',
        re.IGNORECASE,
    ),
    # "X years experience" (without "of")
    re.compile(
        r'(\d{1,2})\s*\+?\s*(?:years?|yrs?)\s+experience',
        re.IGNORECASE,
    ),
]


# ─── Location patterns ──────────────────────────────────────────────────────

_LOCATIONS = [
    # Major US cities / metros
    "New York", "Los Angeles", "Chicago", "Houston", "Phoenix", "Philadelphia",
    "San Antonio", "San Diego", "Dallas", "San Jose", "Austin", "Jacksonville",
    "San Francisco", "Seattle", "Denver", "Boston", "Nashville", "Portland",
    "Las Vegas", "Atlanta", "Miami", "Minneapolis", "Charlotte", "Raleigh",
    "Salt Lake City", "Pittsburgh", "Detroit", "Baltimore", "Tampa",
    "Washington DC", "Washington D.C.",
    # Bay Area / Silicon Valley
    "Bay Area", "Silicon Valley", "Palo Alto", "Mountain View", "Sunnyvale",
    "Cupertino", "Menlo Park", "Redmond",
    # International cities
    "London", "Berlin", "Paris", "Amsterdam", "Dublin", "Toronto", "Vancouver",
    "Montreal", "Sydney", "Melbourne", "Singapore", "Hong Kong", "Tokyo",
    "Bangalore", "Bengaluru", "Hyderabad", "Mumbai", "Delhi", "Pune",
    "Tel Aviv", "Dubai", "Zurich", "Stockholm", "Copenhagen", "Oslo",
    "Helsinki", "Warsaw", "Prague", "Budapest", "Bucharest", "Lisbon",
    "Barcelona", "Madrid", "Milan", "Rome",
    # Countries
    "United States", "USA", "Canada", "United Kingdom", "UK",
    "Germany", "France", "Netherlands", "Ireland", "Australia", "India",
    "Israel", "UAE", "Switzerland", "Sweden", "Denmark", "Norway",
    "Finland", "Poland", "Czech Republic", "Hungary", "Romania",
    "Portugal", "Spain", "Italy", "Japan", "South Korea", "Brazil",
    "Mexico", "Argentina", "Remote",
]

# US state abbreviations – matched case-sensitively and only after a comma
# or pipe to avoid false positives ("IN", "OR", "ME", etc.).
_US_STATES = [
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
    "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD",
    "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
    "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY",
]

# Build patterns: longest first, word-boundary anchored
_LOCATION_PATTERNS: list[tuple[str, re.Pattern]] = []
for _loc in sorted(_LOCATIONS, key=len, reverse=True):
    _LOCATION_PATTERNS.append((
        _loc,
        re.compile(r'\b' + re.escape(_loc) + r'\b', re.IGNORECASE),
    ))
# State codes: require ", XX" or "| XX" or " XX" after a city-like word context
# and case-sensitive match
for _st in _US_STATES:
    _LOCATION_PATTERNS.append((
        _st,
        re.compile(r'(?:,\s*|\|\s*)' + _st + r'\b'),
    ))

# "City, ST" / "City ST" pattern — captures the full "City, ST" or "City ST" pair
# for US locations not in the named list
_CITY_STATE_RE = re.compile(
    r'\b([A-Z][a-z]+(?:\s[A-Z][a-z]+)?)'     # City (1-2 capitalized words)
    r'[,\s]\s*'                                # comma or space separator
    r'([A-Z]{2})\b'                            # 2-letter state code
)


# ─── Subject-line name extraction ────────────────────────────────────────────

# Patterns for forwarded recruitment email subjects like:
#   "Fw: Protingent Candidate - Dan T. Tran - Electrical Engineer - LYNK #28651"
#   "Re: Candidate - Jane Smith - Software Engineer"
#   "Fwd: ABC Corp Candidate - John Doe - Data Analyst - REQ #12345"

_SUBJECT_NAME_PATTERNS = [
    # "Candidate - FirstName [M.] LastName - Title"
    re.compile(
        r'[Cc]andidate\s*[-–—:]\s*'
        r'([A-Z][a-zA-Z]+(?:\s+[A-Z]\.?)?\s+[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)?)'
        r'\s*[-–—]',
    ),
    # "Candidate: FirstName LastName -"
    re.compile(
        r'[Cc]andidate\s*:\s*'
        r'([A-Z][a-zA-Z]+(?:\s+[A-Z]\.?)?\s+[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)?)'
        r'\s*[-–—]',
    ),
    # "Candidate - FirstName LastName" (at end of subject, no trailing dash)
    re.compile(
        r'[Cc]andidate\s*[-–—:]\s*'
        r'([A-Z][a-zA-Z]+(?:\s+[A-Z]\.?)?\s+[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z]+)?)'
        r'\s*$',
    ),
]

# "- Title -" or "- Title - LYNK" pattern to extract job title from subject
_SUBJECT_TITLE_PATTERNS = [
    # After the name segment: "- Title - LYNK/REQ/#"
    re.compile(
        r'[-–—]\s*'
        r'([A-Z][A-Za-z /&]+(?:Engineer|Developer|Analyst|Manager|Designer|Architect|Scientist|Lead|Director|Administrator|Specialist|Consultant|Coordinator|Technician|Intern)(?:\s+\w+)?)'
        r'\s*(?:[-–—]|$)',
        re.IGNORECASE,
    ),
]


def extract_name_from_subject(subject: str) -> str | None:
    """
    Parse candidate name from forwarded recruitment email subjects.

    Handles patterns like:
      "Fw: Protingent Candidate - Dan T. Tran - Electrical Engineer - LYNK #28651"
      "Re: Candidate - Jane Smith - Software Engineer"

    Returns the candidate name string or None if no pattern matched.
    """
    if not subject:
        return None

    # Strip common Fw/Fwd/Re prefixes
    cleaned = re.sub(r'^(?:(?:Fw|Fwd|Re)\s*:\s*)+', '', subject, flags=re.IGNORECASE).strip()

    for pat in _SUBJECT_NAME_PATTERNS:
        m = pat.search(cleaned)
        if m:
            name = m.group(1).strip()
            # Reject if too many words (likely a sentence, not a name)
            if 2 <= len(name.split()) <= 4:
                return name

    return None


def extract_title_from_subject(subject: str) -> str | None:
    """
    Parse job title from forwarded recruitment email subjects.

    Handles patterns like:
      "Fw: Protingent Candidate - Dan T. Tran - Electrical Engineer - LYNK #28651"

    Returns the title string or None.
    """
    if not subject:
        return None

    cleaned = re.sub(r'^(?:(?:Fw|Fwd|Re)\s*:\s*)+', '', subject, flags=re.IGNORECASE).strip()

    for pat in _SUBJECT_TITLE_PATTERNS:
        m = pat.search(cleaned)
        if m:
            title = m.group(1).strip()
            if 2 <= len(title) <= 80:
                return title

    return None


# ─── Phone patterns (for resume text) ───────────────────────────────────────

_PHONE_PATTERNS = [
    # +91-XXXXXXXXXX or +91 XXXXXXXXXX (Indian format)
    re.compile(r'\+91[\s\-.]?\d{5}[\s\-.]?\d{5}\b'),
    # +1 (555) 123-4567, (555) 123-4567, 555-123-4567 (US/Canada)
    re.compile(
        r'(?:\+1[\s\-.]?)?'
        r'\(?\d{3}\)?[\s\-.]?'
        r'\d{3}[\s\-.]?\d{4}\b'
    ),
    # +44 20 7946 0958 (UK / international)
    re.compile(
        r'\+\d{1,3}[\s\-.]?'
        r'\d{2,4}[\s\-.]?'
        r'\d{3,4}[\s\-.]?\d{3,4}\b'
    ),
]


# ─── Entity extraction ──────────────────────────────────────────────────────

def extract_entities(raw_text: str) -> dict:
    """
    Extract structured fields from raw resume / JD text.

    Returns
    -------
    dict with keys:
        skills     : list[str]  – matched technical skills
        years_exp  : int | None – parsed years of experience
        titles     : list[str]  – extracted job titles
        locations  : list[str]  – city / country / state matches
        phone      : str | None – first phone number found
        name       : str | None – always None (set by pipeline from subject/email)
    """
    if not raw_text:
        return {"skills": [], "years_exp": None, "titles": [], "locations": [], "phone": None, "name": None}

    # --- Skills ---
    found_skills: list[str] = []
    seen_skills: set[str] = set()
    for canonical, pat in _SKILL_PATTERNS:
        if pat.search(raw_text):
            key = canonical.lower()
            if key not in seen_skills:
                seen_skills.add(key)
                found_skills.append(canonical)

    # --- Years of experience ---
    years_exp: int | None = None
    for pat in _YOE_PATTERNS:
        m = pat.search(raw_text)
        if m:
            try:
                val = int(m.group(1))
                if years_exp is None or val > years_exp:
                    years_exp = val
            except (ValueError, IndexError):
                pass

    # --- Titles ---
    found_titles: list[str] = []
    seen_titles: set[str] = set()
    for pat in _TITLE_PATTERNS:
        for m in pat.finditer(raw_text):
            title = m.group(0).strip()
            key = title.lower()
            if key and key not in seen_titles:
                seen_titles.add(key)
                found_titles.append(title)

    # --- Locations ---
    found_locations: list[str] = []
    seen_locations: set[str] = set()
    for canonical, pat in _LOCATION_PATTERNS:
        if pat.search(raw_text):
            key = canonical.lower()
            if key not in seen_locations:
                seen_locations.add(key)
                found_locations.append(canonical)

    # Also try the generic "City, ST" / "City ST" pattern
    for m in _CITY_STATE_RE.finditer(raw_text):
        city = m.group(1)
        state = m.group(2)
        if state in _US_STATES:
            loc_str = f"{city}, {state}"
            key = loc_str.lower()
            if key not in seen_locations:
                seen_locations.add(key)
                # Remove bare city or bare state if "City, ST" is more specific
                city_key = city.lower()
                state_key = state.lower()
                for bare in (city_key, state_key):
                    if bare in seen_locations:
                        found_locations = [l for l in found_locations if l.lower() != bare]
                        seen_locations.discard(bare)
                seen_locations.add(key)
                found_locations.append(loc_str)

    # --- Phone ---
    phone: str | None = None
    for pat in _PHONE_PATTERNS:
        m = pat.search(raw_text)
        if m:
            raw = m.group(0).strip()
            digits = re.sub(r'\D', '', raw)
            if len(digits) >= 7:
                phone = raw
                break

    return {
        "skills": found_skills,
        "years_exp": years_exp,
        "titles": found_titles,
        "locations": found_locations,
        "phone": phone,
        "name": None,
    }
