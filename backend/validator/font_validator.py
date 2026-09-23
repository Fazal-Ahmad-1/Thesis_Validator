"""
Validates font family and size used in a DOCX document against a
university's allowed font rules.

Supports two rule schemas found across the current rule profiles:

1. "allowed" list schema (e.g. asu_graduate.json):
       "font": {
           "allowed": [
               {"family": "Arial", "size": 10},
               {"family": "Times New Roman", "size": 12}
           ]
       }
   Any of the listed family/size combinations is acceptable.

2. "family" + "body_size" schema (e.g. amity_selected.json,
   mumbai_engineering.json):
       "font": {
           "family": "Times New Roman",
           "body_size": 12
       }
   Only that single family/size combination is acceptable.

Both schemas are normalized internally into the same allowed set of
(family, size) combinations, so the rest of the validation logic and
the issue structure are identical regardless of which schema a
profile uses. If neither schema is recognized, a ValueError is raised
rather than silently skipping the check.
"""

from typing import Optional
from validator.paragraph_scanner import iter_all_paragraphs
from validator.rule_helpers import get_configured_heading_style_names


def _get_effective_font_name(run, paragraph) -> Optional[str]:
    if run.font.name:
        return run.font.name

    style = run.style
    while style is not None:
        if style.font.name:
            return style.font.name
        style = style.base_style

    style = paragraph.style
    while style is not None:
        if style.font.name:
            return style.font.name
        style = style.base_style

    return None


def _get_effective_font_size(run, paragraph) -> Optional[float]:
    if run.font.size is not None:
        return run.font.size.pt

    style = run.style
    while style is not None:
        if style.font.size is not None:
            return style.font.size.pt
        style = style.base_style

    style = paragraph.style
    while style is not None:
        if style.font.size is not None:
            return style.font.size.pt
        style = style.base_style

    return None


def _format_size(size: float) -> str:
    return f"{size:g}"


def _extract_allowed_combos(font_rules: dict) -> list:
    """
    Normalize either supported font rule schema into a list of
    {"family": str, "size": float} dicts.

    Raises:
        ValueError: If neither supported schema is present.
    """
    # Schema 1: "allowed" list of family/size combinations.
    if "allowed" in font_rules:
        if not isinstance(font_rules["allowed"], list):
            raise ValueError(
                "Unsupported font rule format: 'rules.font.allowed' must "
                "be a list of {family, size} combinations."
            )

        combos = []
        for combo in font_rules["allowed"]:
            family = combo.get("family")
            size = combo.get("size")
            if not family or size is None:
                raise ValueError(
                    "Unsupported font rule format: each entry in 'allowed' "
                    "must have a 'family' and a 'size'."
                )
            combos.append({"family": family, "size": float(size)})
        return combos

    # Schema 2: single "family" + "body_size" pair.
    if "family" in font_rules and "body_size" in font_rules:
        family = font_rules["family"]
        size = font_rules["body_size"]
        if not family or size is None:
            raise ValueError(
                "Unsupported font rule format: 'family' and 'body_size' "
                "must both have values."
            )
        return [{"family": family, "size": float(size)}]

    # Neither recognized schema was found — fail loudly rather than
    # silently skipping font validation.
    raise ValueError(
        "Unsupported font rule format: expected either 'rules.font.allowed' "
        "(a list of {family, size} combinations) or 'rules.font.family' "
        "with 'rules.font.body_size'."
    )


def validate_font_settings(document, rules: dict) -> list:
    if "rules" not in rules:
        raise ValueError("Rule profile is missing the 'rules' section.")

    if "font" not in rules["rules"]:
        raise ValueError("Rule profile is missing the 'rules.font' section.")

    font_rules = rules["rules"]["font"]

    allowed_combo_list = _extract_allowed_combos(font_rules)
    allowed_combos = {
        (combo["family"].strip().lower(), combo["size"]) for combo in allowed_combo_list
    }

    violations = {}
    heading_style_names = get_configured_heading_style_names(rules)

    for paragraph_index, paragraph in enumerate(iter_all_paragraphs(document)):
        if paragraph.style is not None and paragraph.style.name in heading_style_names:
            continue
        for run in paragraph.runs:
            text = run.text.strip()
            if not text:
                continue

            family = _get_effective_font_name(run, paragraph)
            size = _get_effective_font_size(run, paragraph)

            if family is None or size is None:
                continue

            key = (family.strip().lower(), float(size))
            if key in allowed_combos:
                continue

            if key not in violations:
                snippet = text if len(text) <= 40 else text[:40] + "..."
                violations[key] = {
                    "family": family,
                    "size": size,
                    "count": 0,
                    "example_paragraph": paragraph_index,
                    "example_text": snippet,
                }
            violations[key]["count"] += 1

    allowed_description = ", ".join(
        f"{combo['family']} {_format_size(combo['size'])}pt" for combo in allowed_combo_list
    )

    errors = []
    for v in violations.values():
        actual_description = f"{v['family']} {_format_size(v['size'])}pt"
        errors.append({
            "category": "Font",
            "type": "Font",
            "rule": "Font",
            "severity": "error",
            "message": (
                f"Some text uses {actual_description}, which isn't an "
                f"approved font for this profile."
            ),
            "expected": f"One of: {allowed_description}",
            "actual": actual_description,
            "location": f"Paragraph {v['example_paragraph']}",
            "text_snippet": v["example_text"],
            "occurrences": v["count"],
            "suggestion": (
                "Change the affected text to one of the approved font "
                "family and size combinations."
            ),
        })

    return errors