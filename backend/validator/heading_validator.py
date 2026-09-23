"""
Validates heading formatting (font size, bold, uppercase) against a
university's rule profile, using Word's built-in heading styles as the
sole detection mechanism.

Design constraints:
    - Headings are detected ONLY by an explicitly configured Word style
      name (e.g. "Heading 1"). Text content, font size, boldness, or
      capitalization are NEVER used to infer that a paragraph is a
      heading — only paragraph.style.name is checked.
    - Heading validation is entirely optional: if a rule profile has no
      "headings" section, this validator returns an empty list rather
      than raising, since not every profile defines heading rules.
    - A heading level's rule config must include a "style" key naming
      the Word style to match. If it doesn't, that level is skipped
      rather than guessed at.
    - Only the specific sub-properties present in a heading level's
      rule config are checked (e.g. a level with no "uppercase" key
      never gets an uppercase check).
    - A heading level that never appears in the document is NOT
      flagged — its absence is not an error.
    - Effective font size and boldness are resolved by walking the
      same run -> run.style -> paragraph.style chain used elsewhere in
      this project (see font_validator.py). If a value can't be
      resolved anywhere in that chain for a given run, that run is
      excluded from the comparison rather than treated as a violation.
    - Each heading PARAGRAPH is checked once per property, regardless
      of how many runs it's split into:
        - If every resolvable run agrees and matches the expected
          value, the heading passes that property.
        - If every resolvable run agrees but disagrees with the
          expected value, that's one violation for that heading/property.
        - If resolvable runs disagree with EACH OTHER (inconsistent
          formatting within the same heading), that's also one
          violation for that heading/property, reported as "Mixed"
          rather than picking one run's value arbitrarily.
      In all cases, a heading spanning several runs produces at most
      one violation per property — never one per run.
"""

from typing import Optional

MIXED = "Mixed"


def _get_effective_bold(run, paragraph) -> Optional[bool]:
    """Resolve the effective bold state for a run, walking the style chain."""
    if run.font.bold is not None:
        return run.font.bold

    style = run.style
    while style is not None:
        if style.font.bold is not None:
            return style.font.bold
        style = style.base_style

    style = paragraph.style
    while style is not None:
        if style.font.bold is not None:
            return style.font.bold
        style = style.base_style

    return None


def _get_effective_size(run, paragraph) -> Optional[float]:
    """Resolve the effective font size (in points) for a run, walking the style chain."""
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


def _resolve_all_run_values(paragraph):
    """
    Collect the effective size and bold value from every non-empty run
    in a heading paragraph (not just the first one).

    Returns:
        (sizes, bolds) — two lists of the resolved values found across
        all runs. Runs where a property can't be resolved anywhere in
        the style chain simply don't contribute a value for that
        property; they are not treated as violations.
    """
    sizes = []
    bolds = []

    for run in paragraph.runs:
        if not run.text.strip():
            continue

        size = _get_effective_size(run, paragraph)
        if size is not None:
            sizes.append(size)

        bold = _get_effective_bold(run, paragraph)
        if bold is not None:
            bolds.append(bold)

    return sizes, bolds


def _evaluate_property(resolved_values, expected):
    """
    Compare all resolved values for a property against the expected
    value for a single heading paragraph.

    Returns:
        None if there's nothing to report (no resolvable values, or
        all resolvable runs already match expected).
        Otherwise, the "actual" value to report: either the single
        (wrong) value every run agreed on, or MIXED if the runs within
        this heading disagree with each other.
    """
    if not resolved_values:
        # Can't resolve this property anywhere in the heading — skip
        # rather than guess.
        return None

    unique_values = set(resolved_values)

    if len(unique_values) > 1:
        # Runs within the same heading disagree with each other.
        return MIXED

    value = next(iter(unique_values))
    if value != expected:
        return value

    return None


def validate_heading_settings(document, rules: dict) -> list:
    """
    Check heading formatting against a rule profile's "headings" section.

    Args:
        document: A python-docx Document object.
        rules: The full rule profile dictionary.

    Returns:
        A list of structured error dictionaries (category/type/severity/
        message/expected/actual/location/suggestion), or an empty list
        if the profile defines no heading rules.
    """
    rule_sections = rules.get("rules", {})
    headings_rules = rule_sections.get("headings")

    if not headings_rules:
        return []

    errors = []

    for level_key, level_rule in headings_rules.items():
        if level_key.startswith("_"):
            # Skip documentation/metadata keys (e.g. "_note").
            continue

        style_name = level_rule.get("style")
        if not style_name:
            # No explicit Word style mapping for this level — we cannot
            # reliably detect these headings, so skip rather than guess.
            continue

        matching_paragraphs = [
            (index, paragraph)
            for index, paragraph in enumerate(document.paragraphs)
            if paragraph.style is not None and paragraph.style.name == style_name
        ]

        if not matching_paragraphs:
            # This heading level simply isn't used in the document.
            # Absence is not a violation.
            continue

        expected_size = level_rule.get("size")
        expected_bold = level_rule.get("bold")
        expected_uppercase = level_rule.get("uppercase")

        # Grouped by property + actual value (or MIXED), one entry per
        # heading PARAGRAPH regardless of run count — consistent with
        # font_validator's "group by actual value" approach.
        size_violations = {}
        bold_violations = {}
        uppercase_violations = {}

        for paragraph_index, paragraph in matching_paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue

            if expected_uppercase is not None:
                is_uppercase = text == text.upper()
                if is_uppercase != expected_uppercase:
                    key = is_uppercase
                    if key not in uppercase_violations:
                        uppercase_violations[key] = {
                            "count": 0,
                            "example_paragraph": paragraph_index,
                            "example_text": text if len(text) <= 40 else text[:40] + "...",
                        }
                    uppercase_violations[key]["count"] += 1

            sizes, bolds = _resolve_all_run_values(paragraph)

            if expected_size is not None:
                size_result = _evaluate_property(sizes, float(expected_size))
                if size_result is not None:
                    key = size_result  # either a float or MIXED
                    if key not in size_violations:
                        size_violations[key] = {
                            "count": 0,
                            "example_paragraph": paragraph_index,
                            "example_text": text if len(text) <= 40 else text[:40] + "...",
                        }
                    size_violations[key]["count"] += 1

            if expected_bold is not None:
                bold_result = _evaluate_property(bolds, expected_bold)
                if bold_result is not None:
                    key = bold_result  # either a bool or MIXED
                    if key not in bold_violations:
                        bold_violations[key] = {
                            "count": 0,
                            "example_paragraph": paragraph_index,
                            "example_text": text if len(text) <= 40 else text[:40] + "...",
                        }
                    bold_violations[key]["count"] += 1

        level_label = level_key.replace("_", " ")

        for actual_size, v in size_violations.items():
            actual_description = MIXED if actual_size == MIXED else f"{_format_size(actual_size)}pt"
            message = (
                f"Some {level_label} text has mixed font sizes within the same heading."
                if actual_size == MIXED
                else (
                    f"Some {level_label} text is {actual_description} instead of "
                    f"the required {_format_size(expected_size)}pt."
                )
            )
            errors.append({
                "category": "Heading",
                "type": "Formatting",
                "rule": f"{level_label.title()} Size",
                "severity": "error",
                "message": message,
                "expected": f"{_format_size(expected_size)}pt",
                "actual": actual_description,
                "location": f"Paragraph {v['example_paragraph']}",
                "text_snippet": v["example_text"],
                "occurrences": v["count"],
                "suggestion": f"Change the affected {level_label} text to a consistent {_format_size(expected_size)}pt.",
            })

        for actual_bold, v in bold_violations.items():
            if actual_bold == MIXED:
                actual_description = MIXED
                message = f"Some {level_label} text has inconsistent bold formatting within the same heading."
            else:
                actual_description = "Bold" if actual_bold else "Not bold"
                message = (
                    f"Some {level_label} text is "
                    f"{'bold' if actual_bold else 'not bold'} instead of "
                    f"{'bold' if expected_bold else 'not bold'}."
                )
            errors.append({
                "category": "Heading",
                "type": "Formatting",
                "rule": f"{level_label.title()} Bold",
                "severity": "error",
                "message": message,
                "expected": "Bold" if expected_bold else "Not bold",
                "actual": actual_description,
                "location": f"Paragraph {v['example_paragraph']}",
                "text_snippet": v["example_text"],
                "occurrences": v["count"],
                "suggestion": (
                    f"Make the affected {level_label} text consistently "
                    f"{'bold' if expected_bold else 'not bold'} in Word."
                ),
            })

        for is_uppercase, v in uppercase_violations.items():
            errors.append({
                "category": "Heading",
                "type": "Formatting",
                "rule": f"{level_label.title()} Uppercase",
                "severity": "error",
                "message": (
                    f"Some {level_label} text is "
                    f"{'uppercase' if is_uppercase else 'not uppercase'} instead of "
                    f"{'uppercase' if expected_uppercase else 'not uppercase'}."
                ),
                "expected": "Uppercase" if expected_uppercase else "Not uppercase",
                "actual": "Uppercase" if is_uppercase else "Not uppercase",
                "location": f"Paragraph {v['example_paragraph']}",
                "text_snippet": v["example_text"],
                "occurrences": v["count"],
                "suggestion": (
                    f"Change the affected {level_label} text to "
                    f"{'uppercase' if expected_uppercase else 'title/sentence case'} in Word."
                ),
            })

    return errors