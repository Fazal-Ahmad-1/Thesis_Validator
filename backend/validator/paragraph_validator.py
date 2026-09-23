"""
Validates paragraph line spacing against a university's rule profile.

Approach and limitations (read before relying on this for edge cases):

paragraph_format.line_spacing can be:
    - a Length object (Pt/Emu-based, an int subclass) -> "exactly" or
      "at least" point-based spacing, NOT a multiple. Comparing this
      numerically against 1.5/2.0 would be meaningless, so it is
      deliberately SKIPPED rather than guessed at.
    - a plain int or float, e.g. 1.5 or 2.0 -> a "multiple" spacing
      value (what we validate here)
    - None -> not set directly on this paragraph; the paragraph
      inherits spacing from its style

When line_spacing is None, this validator walks the paragraph's style
and its base_style chain looking for the first explicit numeric
(multiple) line_spacing value. If none is found (or only a point-based
Length is found), the paragraph is skipped rather than flagged as a
violation. This is a deliberate MVP trade-off, not a bug.
"""

from typing import Optional
from docx.shared import Length
from validator.paragraph_scanner import iter_all_paragraphs
from validator.rule_helpers import get_configured_heading_style_names


def _resolve_spacing_multiple(value) -> Optional[float]:
    """
    Interpret a raw paragraph_format.line_spacing value.

    Returns:
        - a float, if the value is a numeric "multiple" spacing value
        - None if the value is None, or a point-based Length (not a multiple)
    """
    if value is None:
        return None

    # Length is an int subclass, so check for it explicitly and first,
    # rather than relying on it simply not being a float.
    if isinstance(value, Length):
        return None

    if isinstance(value, bool):
        # Guard against the unlikely case of a bool sneaking in
        # (bool is also an int subclass in Python).
        return None

    if isinstance(value, (int, float)):
        return float(value)

    return None


def _get_effective_line_spacing(paragraph) -> Optional[float]:
    """
    Resolve the effective numeric line-spacing multiple for a paragraph.

    Returns a float (e.g. 1.5, 2.0) if it can be reliably determined,
    or None if it can't (point-based spacing, or no spacing set anywhere
    in the style chain).
    """
    direct = _resolve_spacing_multiple(paragraph.paragraph_format.line_spacing)
    if direct is not None:
        return direct
    if paragraph.paragraph_format.line_spacing is not None:
        # A value was set directly but it's point-based (Length) —
        # not a multiple, so don't fall back to the style chain.
        return None

    # Walk the paragraph style's base_style chain, guarding against
    # circular references.
    visited = set()
    style = paragraph.style
    while style is not None and id(style) not in visited:
        visited.add(id(style))

        raw_style_spacing = style.paragraph_format.line_spacing
        style_spacing = _resolve_spacing_multiple(raw_style_spacing)
        if style_spacing is not None:
            return style_spacing
        if raw_style_spacing is not None:
            # Point-based value found in the style chain — not a
            # multiple we can validate.
            return None

        style = style.base_style

    return None


def _format_spacing(value: float) -> str:
    """Display a spacing value without unnecessary trailing zeros (e.g. 2 not 2.0)."""
    return f"{value:g}"


def validate_paragraph_settings(document, rules: dict) -> list:
    """
    Check paragraph line spacing against the expected value in the rule profile.

    Args:
        document: A python-docx Document object.
        rules: The full rule profile dictionary (with a top-level "rules" key).

    Returns:
        A list of structured error dictionaries, one per unique actual
        line-spacing violation found, each including an occurrence
        count and one example location.

    Raises:
        ValueError: If "rules", "rules.paragraph", or
            "rules.paragraph.line_spacing" is missing, or line_spacing
            is not a supported numeric format.
    """
    if "rules" not in rules:
        raise ValueError("Rule profile is missing the 'rules' section.")

    if "paragraph" not in rules["rules"]:
        raise ValueError("Rule profile is missing the 'rules.paragraph' section.")

    paragraph_rules = rules["rules"]["paragraph"]

    if "line_spacing" not in paragraph_rules:
        raise ValueError(
            "Rule profile is missing 'rules.paragraph.line_spacing'."
        )

    expected_spacing = paragraph_rules["line_spacing"]

    if isinstance(expected_spacing, bool) or not isinstance(expected_spacing, (int, float)):
        raise ValueError(
            "Unsupported line_spacing format: expected a numeric value "
            f"(e.g. 1.5 or 2.0), got: {expected_spacing!r}"
        )

    expected_spacing = float(expected_spacing)
    tolerance = 0.05

    # Group violations by actual spacing value to avoid duplicate/noisy errors.
    violations = {}
    heading_style_names = get_configured_heading_style_names(rules)

    for paragraph_index, paragraph in enumerate(iter_all_paragraphs(document)):
        if paragraph.style is not None and paragraph.style.name in heading_style_names:
            continue
        text = paragraph.text.strip()
        if not text:
            continue  # ignore empty/whitespace-only paragraphs

        actual_spacing = _get_effective_line_spacing(paragraph)
        if actual_spacing is None:
            # Can't confidently determine effective spacing; skip
            # rather than risk a false violation (see module docstring).
            continue

        if abs(actual_spacing - expected_spacing) <= tolerance:
            continue

        key = round(actual_spacing, 2)
        if key not in violations:
            snippet = text if len(text) <= 40 else text[:40] + "..."
            violations[key] = {
                "actual": actual_spacing,
                "count": 0,
                "example_paragraph": paragraph_index,
                "example_text": snippet,
            }
        violations[key]["count"] += 1

    errors = []
    for v in violations.values():
        errors.append({
            "category": "Paragraph",
            "type": "Formatting",
            "rule": "Line Spacing",
            "severity": "error",
            "message": (
                f"Some paragraphs use {_format_spacing(v['actual'])} line "
                f"spacing instead of the required {_format_spacing(expected_spacing)}."
            ),
            "expected": _format_spacing(expected_spacing),
            "actual": _format_spacing(v["actual"]),
            "location": f"Paragraph {v['example_paragraph']}",
            "text_snippet": v["example_text"],
            "occurrences": v["count"],
            "suggestion": f"Change the affected paragraphs to {_format_spacing(expected_spacing)} line spacing in Word.",
        })

    return errors