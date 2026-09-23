"""
Conservative V1 pagination validator.

IMPORTANT LIMITATION: python-docx cannot determine the final rendered
page count, or which page number appears on any given rendered page,
because that depends entirely on Word/LibreOffice layout at render
time. This validator does NOT attempt to verify rendered page numbers
or rendered page counts.

It only checks facts that are explicitly represented in the DOCX XML:
    - whether a PAGE field exists in the footer (or header)
    - the alignment of the paragraph containing that field
    - the explicit page-numbering format/start declared on a section
      (w:pgNumType in sectPr), when the document declares one

Position handling: if the rule profile specifies "top" or "bottom",
only that location (header or footer respectively) is searched. If
position is omitted, both header and footer are searched, and finding
the field in either one satisfies the check — no default location is
assumed, since that would risk a false violation.

Alignment is only checked when the rule profile provides an explicit
"alignment" value. If a profile doesn't specify one, no alignment
check is performed at all.

When a check cannot be reliably answered from the document's
structure (e.g. no section explicitly declares a numbering format, or
a page-number paragraph's alignment is inherited from its style rather
than set directly), this validator reports that as a limitation
(severity "info") rather than guessing or reporting a false violation.
Consumers of the errors list should treat "info" items as notices, not
blocking violations.
"""

import re
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}
ALIGNMENT_NAMES = {value: key for key, value in ALIGNMENT_MAP.items()}


def _paragraph_has_page_field(paragraph) -> bool:
    """Return True if this paragraph contains a Word PAGE field (simple or complex)."""
    p = paragraph._p

    # Simple field: <w:fldSimple w:instr="PAGE ...">
    for fld_simple in p.findall(qn("w:fldSimple")):
        instr = fld_simple.get(qn("w:instr")) or ""
        if re.search(r"\bPAGE\b", instr, re.IGNORECASE):
            return True

    # Complex field: <w:instrText>PAGE ...</w:instrText>
    for instr_text_el in p.iter(qn("w:instrText")):
        if instr_text_el.text and re.search(r"\bPAGE\b", instr_text_el.text, re.IGNORECASE):
            return True

    return False


def _find_page_field_paragraphs(container):
    """Return the paragraphs in a header/footer that contain a PAGE field."""
    return [p for p in container.paragraphs if _paragraph_has_page_field(p)]


def _get_section_page_num_type(section):
    """
    Read the explicit page numbering format/start declared in a section's
    properties (w:pgNumType), if present.

    Returns a dict {"fmt": str or None, "start": int or None}, or None if
    the section does not explicitly declare a page numbering format.
    """
    sect_pr = section._sectPr
    pg_num_type = sect_pr.find(qn("w:pgNumType"))
    if pg_num_type is None:
        return None

    fmt = pg_num_type.get(qn("w:fmt"))
    start_raw = pg_num_type.get(qn("w:start"))
    start = int(start_raw) if start_raw is not None else None

    return {"fmt": fmt, "start": start}


def validate_pagination_settings(document, rules: dict) -> list:
    """
    Check pagination-related structure against the rule profile.

    Args:
        document: A python-docx Document object.
        rules: The full rule profile dictionary (with a top-level "rules" key).

    Returns:
        A list of structured result dictionaries in the same shape used
        by the other validators (category/type/rule/severity/message/
        expected/actual/location/occurrences/suggestion). "severity" is
        explicitly "error" or "info" on every entry.

    Raises:
        ValueError: If "rules" or "rules.pagination" is missing.
    """
    if "rules" not in rules:
        raise ValueError("Rule profile is missing the 'rules' section.")

    if "pagination" not in rules["rules"]:
        raise ValueError("Rule profile is missing the 'rules.pagination' section.")

    pagination_rules = rules["rules"]["pagination"]

    require_page_field = pagination_rules.get("require_page_field", False)
    position = pagination_rules.get("position")  # "top" / "bottom" / None
    expected_alignment = pagination_rules.get("alignment")  # "center" / etc. / None
    preliminary_format = pagination_rules.get("preliminary_format")  # e.g. "lowerRoman"
    body_format = pagination_rules.get("body_format")  # e.g. "decimal"
    body_starts_at = pagination_rules.get("body_starts_at")  # e.g. 1

    results = []
    sections = document.sections

    # --- Checks 1 & 2: PAGE field presence and alignment ---
    if require_page_field:
        field_paragraphs = []  # list of (section_index, container_name, paragraph)

        for index, section in enumerate(sections):
            if position == "top":
                containers = [("header", section.header)]
            elif position == "bottom":
                containers = [("footer", section.footer)]
            else:
                # No position specified: search both header and footer.
                # Do not assume a default location, since that could
                # produce a false violation.
                containers = [("header", section.header), ("footer", section.footer)]

            for container_name, container in containers:
                for paragraph in _find_page_field_paragraphs(container):
                    field_paragraphs.append((index, container_name, paragraph))

        if not field_paragraphs:
            if position == "top":
                container_description = "header"
            elif position == "bottom":
                container_description = "footer"
            else:
                container_description = "header or footer"
            results.append({
                "category": "Pagination",
                "type": "Formatting",
                "rule": "Page Number Field",
                "severity": "error",
                "message": f"No page number field was found in the document {container_description}.",
                "expected": f"A PAGE field present in the document {container_description}",
                "actual": "No PAGE field found",
                "location": "Document sections",
                "occurrences": len(sections),
                "suggestion": f"Add a PAGE field to the {container_description} so Word can display the page number.",
            })
        elif expected_alignment:
            expected_wd_alignment = ALIGNMENT_MAP.get(expected_alignment.lower())
            misaligned = {}
            unverifiable_count = 0

            for section_index, container_name, paragraph in field_paragraphs:
                actual_alignment = paragraph.alignment
                if actual_alignment is None:
                    # Alignment may be inherited from the paragraph style
                    # rather than set directly on this paragraph. V1 does
                    # not walk the style chain here; report this as an
                    # informational notice rather than guessing.
                    unverifiable_count += 1
                    continue

                if actual_alignment != expected_wd_alignment:
                    actual_name = ALIGNMENT_NAMES.get(actual_alignment, str(actual_alignment))
                    if actual_name not in misaligned:
                        misaligned[actual_name] = {
                            "count": 0,
                            "example_section": section_index,
                            "example_container": container_name,
                        }
                    misaligned[actual_name]["count"] += 1

            for actual_name, info in misaligned.items():
                results.append({
                    "category": "Pagination",
                    "type": "Formatting",
                    "rule": "Page Number Alignment",
                    "severity": "error",
                    "message": (
                        f"The page number is {actual_name}-aligned instead of "
                        f"{expected_alignment}-aligned."
                    ),
                    "expected": expected_alignment,
                    "actual": actual_name,
                    "location": f"Section {info['example_section'] + 1} {info['example_container']}",
                    "occurrences": info["count"],
                    "suggestion": f"Change the page number paragraph's alignment to {expected_alignment} in Word.",
                })

            if unverifiable_count:
                results.append({
                    "category": "Pagination",
                    "type": "Formatting",
                    "rule": "Page Number Alignment",
                    "severity": "info",
                    "message": (
                        "The page number's alignment is inherited from its "
                        "paragraph style rather than set directly, so it "
                        "cannot be reliably verified."
                    ),
                    "expected": expected_alignment,
                    "actual": (
                        "Alignment not explicitly set on the page-number "
                        "paragraph (inherited from style); cannot be "
                        "reliably verified"
                    ),
                    "location": "Document sections",
                    "occurrences": unverifiable_count,
                    "suggestion": (
                        f"Manually confirm the page number is {expected_alignment}-aligned, "
                        "since this can't be automatically verified from the document structure."
                    ),
                })
        # If expected_alignment is not set, no alignment check is
        # performed at all, per profile requirements.

    # --- Checks 3 & 4: Numbering format / start (only when explicitly declared) ---
    if preliminary_format or body_format or body_starts_at is not None:
        declared = [
            (index, _get_section_page_num_type(section))
            for index, section in enumerate(sections)
        ]
        explicitly_declared = [(i, d) for i, d in declared if d is not None]

        if not explicitly_declared:
            results.append({
                "category": "Pagination",
                "type": "Formatting",
                "rule": "Page Numbering Format",
                "severity": "info",
                "message": (
                    "No section explicitly declares a page numbering format "
                    "in the document, so the required preliminary/main-text "
                    "numbering format cannot be reliably verified."
                ),
                "expected": (
                    f"Preliminary pages: {preliminary_format or 'N/A'}, "
                    f"main text: {body_format or 'N/A'}"
                ),
                "actual": (
                    "No section explicitly declares a page numbering format "
                    "(w:pgNumType) in the document XML; this cannot be "
                    "reliably verified without rendering the document"
                ),
                "location": "Document sections",
                "occurrences": len(sections),
                "suggestion": (
                    "Manually confirm your preliminary and main text sections use "
                    "the required page numbering formats, since this can't be "
                    "automatically verified from the document structure."
                ),
            })
        else:
            # V1 treats the last section that explicitly declares a
            # format as the main-text section.
            last_index, last_decl = explicitly_declared[-1]

            if body_format and last_decl["fmt"] != body_format:
                actual_fmt = last_decl["fmt"] or "Not set"
                results.append({
                    "category": "Pagination",
                    "type": "Formatting",
                    "rule": "Main Text Page Numbering Format",
                    "severity": "error",
                    "message": (
                        f"The main text section uses '{actual_fmt}' page "
                        f"numbering instead of the required '{body_format}'."
                    ),
                    "expected": body_format,
                    "actual": actual_fmt,
                    "location": f"Section {last_index + 1}",
                    "occurrences": 1,
                    "suggestion": (
                        f"Set the main text section's page numbering format "
                        f"to {body_format} in Word's Page Number Format settings."
                    ),
                })

            if body_starts_at is not None:
                if last_decl["start"] is None:
                    results.append({
                        "category": "Pagination",
                        "type": "Formatting",
                        "rule": "Main Text Page Numbering Start",
                        "severity": "info",
                        "message": (
                            "The main text section doesn't explicitly declare "
                            "a starting page number, so this cannot be "
                            "reliably verified."
                        ),
                        "expected": str(body_starts_at),
                        "actual": (
                            "Start value not explicitly declared for this "
                            "section; cannot be reliably verified"
                        ),
                        "location": f"Section {last_index + 1}",
                        "occurrences": 1,
                        "suggestion": (
                            f"Manually confirm the main text section starts numbering "
                            f"at {body_starts_at}, since this can't be automatically verified."
                        ),
                    })
                elif last_decl["start"] != body_starts_at:
                    results.append({
                        "category": "Pagination",
                        "type": "Formatting",
                        "rule": "Main Text Page Numbering Start",
                        "severity": "error",
                        "message": (
                            f"The main text section starts numbering at "
                            f"{last_decl['start']} instead of {body_starts_at}."
                        ),
                        "expected": str(body_starts_at),
                        "actual": str(last_decl["start"]),
                        "location": f"Section {last_index + 1}",
                        "occurrences": 1,
                        "suggestion": f"Set the main text section to start page numbering at {body_starts_at} in Word.",
                    })

            if preliminary_format and len(explicitly_declared) > 1:
                earlier_declarations = explicitly_declared[:-1]
                mismatched = [
                    (i, d) for i, d in earlier_declarations if d["fmt"] != preliminary_format
                ]
                if mismatched:
                    example_index = mismatched[0][0]
                    results.append({
                        "category": "Pagination",
                        "type": "Formatting",
                        "rule": "Preliminary Page Numbering Format",
                        "severity": "error",
                        "message": (
                            f"One or more preliminary sections don't use "
                            f"'{preliminary_format}' page numbering as required."
                        ),
                        "expected": preliminary_format,
                        "actual": "Mismatched or unset format in an earlier section",
                        "location": f"Section {example_index + 1}",
                        "occurrences": len(mismatched),
                        "suggestion": (
                            f"Set the preliminary section(s) page numbering "
                            f"format to {preliminary_format} in Word."
                        ),
                    })
            elif preliminary_format and len(explicitly_declared) == 1:
                results.append({
                    "category": "Pagination",
                    "type": "Formatting",
                    "rule": "Preliminary Page Numbering Format",
                    "severity": "info",
                    "message": (
                        "Only one section explicitly declares a numbering "
                        "format, so preliminary vs. main-text numbering "
                        "cannot be reliably verified."
                    ),
                    "expected": preliminary_format,
                    "actual": (
                        "Only one section explicitly declares a numbering "
                        "format; preliminary vs. main-text formatting "
                        "cannot be reliably verified"
                    ),
                    "location": "Document sections",
                    "occurrences": 1,
                    "suggestion": (
                        "Manually confirm your preliminary pages use the required "
                        "numbering format, since this can't be automatically verified."
                    ),
                })

    return results