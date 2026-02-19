"""
Candidate-to-job matching engine.

Scoring rubric (100 pts max):
    Skill overlap   50 pts  –  (matched / required) * 50
    Experience      20 pts  –  full if candidate.years_exp >= job.min_exp
    Location        15 pts  –  full if job.remote_ok OR locations match
    Title           15 pts  –  full if any candidate title appears in job title
"""

import json
import logging

from database import SessionLocal, Candidate, JobRequisition, MatchResult

logger = logging.getLogger(__name__)


# ─── Single-pair scoring ────────────────────────────────────────────────────

def _score_candidate(candidate: Candidate, job: JobRequisition) -> dict:
    """
    Score one candidate against one job.

    Returns
    -------
    dict with keys: score (float 0-100), match_reasons (list[str]), fit_level (str)
    """
    reasons: list[str] = []
    score = 0.0

    cand_skills = {s.lower() for s in (json.loads(candidate.skills) if candidate.skills else [])}
    job_skills = {s.lower() for s in (json.loads(job.required_skills) if job.required_skills else [])}
    cand_titles = [t.lower() for t in (json.loads(candidate.titles) if candidate.titles else [])]
    job_title_lower = (job.title or "").lower()
    cand_exp = candidate.years_exp or 0
    min_exp = job.min_exp or 0
    cand_loc = (candidate.location or "").lower()
    job_loc = (job.location or "").lower()

    # ── 1. Skill overlap  (max 50 pts) ──────────────────────────────────
    if job_skills:
        matched = cand_skills & job_skills
        missing = job_skills - cand_skills
        skill_score = round((len(matched) / len(job_skills)) * 50, 1)
        score += skill_score
        if matched:
            reasons.append(f"Skills matched ({len(matched)}/{len(job_skills)}): "
                           f"{', '.join(sorted(matched))}")
        if missing:
            reasons.append(f"Skills missing ({len(missing)}): "
                           f"{', '.join(sorted(missing))}")
    else:
        score += 25
        reasons.append("No specific skills required — partial credit")

    # ── 2. Experience  (max 20 pts) ─────────────────────────────────────
    if min_exp > 0:
        if cand_exp >= min_exp:
            score += 20
            reasons.append(f"Experience {cand_exp:.0f}y meets requirement ({min_exp:.0f}y)")
        else:
            gap = min_exp - cand_exp
            reasons.append(f"Experience {cand_exp:.0f}y is {gap:.0f}y short of requirement ({min_exp:.0f}y)")
    else:
        score += 20
        reasons.append("No minimum experience required")

    # ── 3. Location  (max 15 pts) ───────────────────────────────────────
    if job.remote_ok:
        score += 15
        reasons.append("Remote OK — location flexible")
    elif job_loc and cand_loc:
        job_tokens = set(job_loc.replace(",", " ").split())
        cand_tokens = set(cand_loc.replace(",", " ").split())
        if job_tokens & cand_tokens:
            score += 15
            reasons.append(f"Location match: {candidate.location}")
        else:
            reasons.append(f"Location mismatch: {candidate.location or 'unknown'} vs {job.location}")
    elif not job_loc:
        score += 15
        reasons.append("No location requirement")
    else:
        reasons.append("Candidate location unknown")

    # ── 4. Title relevance  (max 15 pts) ────────────────────────────────
    if job_title_lower and cand_titles:
        title_hit = any(ct in job_title_lower for ct in cand_titles)
        if title_hit:
            score += 15
            matching = [ct for ct in cand_titles if ct in job_title_lower]
            reasons.append(f"Title match: {matching[0]}")
        else:
            reasons.append(f"Title mismatch: candidate titles "
                           f"[{', '.join(cand_titles[:3])}] not found in '{job.title}'")
    elif not job_title_lower:
        score += 15
    else:
        reasons.append("Candidate has no titles on file")

    # ── Final ───────────────────────────────────────────────────────────
    score = min(round(score, 1), 100.0)

    if score >= 75:
        fit_level = "high"
    elif score >= 45:
        fit_level = "medium"
    else:
        fit_level = "low"

    return {"score": score, "match_reasons": reasons, "fit_level": fit_level}


# ─── Helpers ────────────────────────────────────────────────────────────────

def _candidate_to_dict(c: Candidate) -> dict:
    return {
        "id": c.id,
        "name": c.name,
        "email": c.email,
        "phone": c.phone,
        "location": c.location,
        "titles": json.loads(c.titles) if c.titles else [],
        "skills": json.loads(c.skills) if c.skills else [],
        "years_exp": c.years_exp,
        "raw_resume_path": c.raw_resume_path,
        "source_email_uid": c.source_email_uid,
        "notes": c.notes or "",
        "tags": json.loads(c.tags) if c.tags else [],
        "created_at": c.created_at.isoformat() if c.created_at else None,
    }


def _job_to_dict(j: JobRequisition) -> dict:
    return {
        "id": j.id,
        "title": j.title,
        "required_skills": json.loads(j.required_skills) if j.required_skills else [],
        "min_exp": j.min_exp,
        "location": j.location,
        "remote_ok": j.remote_ok,
        "jd_raw": j.jd_raw,
        "created_at": j.created_at.isoformat() if j.created_at else None,
    }


# ─── Public API ─────────────────────────────────────────────────────────────

def run_match(job_id: int) -> list:
    """
    Match all candidates against a job requisition.

    - Fetches the job and all candidates from the DB.
    - Scores every candidate.
    - Replaces any previous match_results for this job.
    - Returns a list sorted by score (descending), each entry containing
      candidate info, score, match_reasons, and fit_level.

    Raises
    ------
    ValueError  if the job_id does not exist or there are no candidates.
    """
    db = SessionLocal()
    try:
        job = db.query(JobRequisition).filter(JobRequisition.id == job_id).first()
        if not job:
            raise ValueError(f"Job with id={job_id} not found")

        candidates = db.query(Candidate).all()
        if not candidates:
            raise ValueError("No candidates in database")

        # Clear previous results for this job
        db.query(MatchResult).filter(MatchResult.job_id == job_id).delete()

        results = []
        for c in candidates:
            match = _score_candidate(c, job)
            mr = MatchResult(
                job_id=job_id,
                candidate_id=c.id,
                score=match["score"],
                match_reasons=json.dumps(match["match_reasons"]),
                fit_level=match["fit_level"],
            )
            db.add(mr)
            results.append({
                "candidate": _candidate_to_dict(c),
                "score": match["score"],
                "match_reasons": match["match_reasons"],
                "fit_level": match["fit_level"],
            })

        db.commit()

        # Back-fill match IDs from the newly persisted rows
        saved = (
            db.query(MatchResult)
            .filter(MatchResult.job_id == job_id)
            .all()
        )
        id_map = {mr.candidate_id: mr.id for mr in saved}
        for r in results:
            r["match_id"] = id_map.get(r["candidate"]["id"])

        # Sort descending by score
        results.sort(key=lambda r: r["score"], reverse=True)

        logger.info(
            "Matched %d candidates against job id=%d (%s)",
            len(results), job_id, job.title,
        )
        return results
    finally:
        db.close()
