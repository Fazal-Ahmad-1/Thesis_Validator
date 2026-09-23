"""
Orchestrates the DOCX reader, rule loader, and validators (page, font,
paragraph, and pagination) to produce a single validation result for
a document.
"""

from validator.docx_reader import read_docx
from validator.rule_loader import load_rules
from validator.page_validator import validate_page_settings
from validator.font_validator import validate_font_settings
from validator.paragraph_validator import validate_paragraph_settings
from validator.pagination_validator import validate_pagination_settings
from validator.heading_validator import validate_heading_settings

def _calculate_score(blocking_error_count: int) -> int:
    """
    Simple, deterministic formatting-compliance score from 0 to 100.

    Starts at 100 and subtracts 10 points per blocking error, floored
    at 0. Informational notices never affect this score. This score
    represents formatting compliance with the selected rule profile
    only — it says nothing about writing quality, grammar, or academic
    content.
    """
    return max(0, 100 - (10 * blocking_error_count))


def _score_label(score: int) -> str:
    """Map a score to a plain compliance label."""
    if score >= 90:
        return "Excellent"
    if score >= 75:
        return "Good"
    if score >= 50:
        return "Needs Improvement"
    return "Major Formatting Issues"


def _has_blocking_issue(errors: list[dict], category: str) -> bool:
    """Return True when a category has at least one blocking issue."""
    return any(
        error.get("category") == category
        and error.get("severity", "error") != "info"
        for error in errors
    )


def _build_check_summary(rules: dict, errors: list[dict]) -> dict:
    """
    Build explicit per-check applicability/status metadata for the frontend.

    Status values:
        - passed
        - needs_attention
        - not_applicable

    Applicability is derived from the selected rule profile, not from whether
    an issue happened to be found in the document.
    """
    checks = {}

    # These validators are part of the core V1 validation pipeline and are
    # applicable to every current profile.
    for category in ("Page", "Font", "Paragraph"):
        checks[category] = {
            "status": (
                "needs_attention"
                if _has_blocking_issue(errors, category)
                else "passed"
            )
        }

    pagination_rules = rules.get("rules", {}).get("pagination", {})
    pagination_enforced = any(
        key in pagination_rules
        and pagination_rules.get(key) not in (None, "", False)
        for key in (
            "require_page_field",
            "position",
            "alignment",
            "preliminary_format",
            "body_format",
            "body_starts_at",
        )
    )

    if pagination_enforced:
        checks["Pagination"] = {
            "status": (
                "needs_attention"
                if _has_blocking_issue(errors, "Pagination")
                else "passed"
            )
        }
    else:
        checks["Pagination"] = {
            "status": "not_applicable",
            "reason": "Pagination requirements are not enforced by this profile.",
        }

    heading_rules = rules.get("rules", {}).get("headings", {})
    heading_levels = [
        key for key in heading_rules
        if key != "_note" and isinstance(heading_rules.get(key), dict)
    ]

    if heading_levels:
        checks["Heading"] = {
            "status": (
                "needs_attention"
                if _has_blocking_issue(errors, "Heading")
                else "passed"
            )
        }
    else:
        checks["Heading"] = {
            "status": "not_applicable",
            "reason": "Heading requirements are not configured for this profile.",
        }

    return checks


def validate_document(file_path: str, profile_name: str) -> dict:
    """
    Validate a DOCX file against a named university rule profile.

    Args:
        file_path: Path to the .docx file to validate.
        profile_name: Rule profile identifier, e.g. "asu_graduate".

    Returns:
        A dictionary containing:
            - university
            - program
            - document_type
            - profile_name
            - errors: the full combined list (blocking + informational),
              kept for backward compatibility
            - blocking_errors: entries with severity != "info"
            - informational_notices: entries with severity == "info"
            - blocking_error_count: len(blocking_errors)
            - informational_notice_count: len(informational_notices)
            - is_valid: True only when blocking_error_count is 0
            - score: 0-100 formatting-compliance score
            - score_label: a plain label for the score
            - checks: per-category applicability and status metadata

    Raises:
        FileNotFoundError: If the DOCX file or rule profile is missing.
        ValueError: If the DOCX file or rule profile is invalid.
    """
    document = read_docx(file_path)
    rules = load_rules(profile_name)

    page_errors = validate_page_settings(document, rules)
    font_errors = validate_font_settings(document, rules)
    paragraph_errors = validate_paragraph_settings(document, rules)
    pagination_results = validate_pagination_settings(document, rules)
    heading_errors = validate_heading_settings(document, rules)

    errors = page_errors + font_errors + paragraph_errors + pagination_results + heading_errors

    # Split the combined list without altering or duplicating any issue
    # objects — each item is a reference to the same dict either way.
    blocking_errors = [e for e in errors if e.get("severity", "error") != "info"]
    informational_notices = [e for e in errors if e.get("severity") == "info"]

    score = _calculate_score(len(blocking_errors))
    score_label = _score_label(score)
    checks = _build_check_summary(rules, errors)

    return {
        "university": rules.get("university", "Unknown"),
        "program": rules.get("program", "Unknown"),
        "document_type": rules.get("document_type", "Unknown"),
        "profile_name": profile_name,
        "errors": errors,
        "blocking_errors": blocking_errors,
        "informational_notices": informational_notices,
        "blocking_error_count": len(blocking_errors),
        "informational_notice_count": len(informational_notices),
        "is_valid": len(blocking_errors) == 0,
        "score": score,
        "score_label": score_label,
        "checks": checks,
    }