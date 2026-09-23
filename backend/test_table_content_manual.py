"""
Diagnostic test: does the current implementation scan formatting inside
Word TABLE CELLS, or only top-level body paragraphs?

This is NOT a pass/fail regression test. It establishes current behavior
before deciding whether to expand table support.

Run from backend/ with the venv active:
    python test_table_content_manual.py

No production files are modified.
"""

import os

from docx import Document
from docx.shared import Pt, Inches, Cm, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from validator.validator import validate_document


PROFILE_SETTINGS = {
    "asu_graduate": {
        "unit": Inches,
        "width": 8.5,
        "height": 11.0,
        "top": 1.0,
        "bottom": 1.0,
        "left": 1.25,
        "right": 1.25,
        "font_family": "Times New Roman",
        "font_size": 12,
        "line_spacing": 2.0,
        "align_footer": True,
    },
    "amity_selected": {
        "unit": Cm,
        "width": 21.0,
        "height": 29.7,
        "top": 2.5,
        "bottom": 2.5,
        "left": 2.5,
        "right": 2.5,
        "font_family": "Times New Roman",
        "font_size": 12,
        "line_spacing": 1.5,
        "align_footer": True,
    },
    "mumbai_engineering": {
        "unit": Mm,
        "width": 210,
        "height": 297,
        "top": 15,
        "bottom": 22,
        "left": 30,
        "right": 20,
        "font_family": "Times New Roman",
        "font_size": 12,
        "line_spacing": 1.5,
        "align_footer": False,
    },
}


def add_page_field(paragraph):
    """Add a PAGE field to a paragraph."""
    p = paragraph._p

    fld_simple = OxmlElement("w:fldSimple")
    fld_simple.set(qn("w:instr"), "PAGE")

    run_el = OxmlElement("w:r")
    text_el = OxmlElement("w:t")
    text_el.text = "1"

    run_el.append(text_el)
    fld_simple.append(run_el)
    p.append(fld_simple)


def build_document_with_bad_table_cell(profile_name):
    """
    Build an otherwise compliant document containing one table cell
    with deliberately incorrect font and line spacing.
    """
    settings = PROFILE_SETTINGS[profile_name]
    unit = settings["unit"]

    document = Document()
    section = document.sections[0]

    # Correct page settings.
    section.page_width = unit(settings["width"])
    section.page_height = unit(settings["height"])
    section.top_margin = unit(settings["top"])
    section.bottom_margin = unit(settings["bottom"])
    section.left_margin = unit(settings["left"])
    section.right_margin = unit(settings["right"])

    # Correct body paragraph.
    body = document.add_paragraph()
    body_run = body.add_run("This is a compliant body paragraph.")
    body_run.font.name = settings["font_family"]
    body_run.font.size = Pt(settings["font_size"])
    body.paragraph_format.line_spacing = settings["line_spacing"]

    # Create table.
    table = document.add_table(rows=1, cols=2)

    # First cell is irrelevant and left compliant/default.
    table.cell(0, 0).text = "Metric"

    # Second cell contains deliberately bad formatting.
    bad_cell_paragraph = table.cell(0, 1).paragraphs[0]

    bad_cell_run = bad_cell_paragraph.add_run(
        "Deliberately wrong font and line spacing inside a table cell."
    )

    # Explicitly incorrect font.
    bad_cell_run.font.name = "Calibri"
    bad_cell_run.font.size = Pt(11)

    # Explicitly incorrect paragraph spacing.
    bad_cell_paragraph.paragraph_format.line_spacing = 1.0

    # Add PAGE field so pagination does not muddy the diagnostic.
    footer_paragraph = section.footer.paragraphs[0]

    if settings["align_footer"]:
        footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_page_field(footer_paragraph)

    return document


def run_diagnostic(profile_name):
    path = f"_tmp_table_{profile_name}.docx"

    try:
        document = build_document_with_bad_table_cell(profile_name)
        document.save(path)

        result = validate_document(path, profile_name)

        font_errors = [
            error
            for error in result["blocking_errors"]
            if error.get("category") == "Font"
        ]

        paragraph_errors = [
            error
            for error in result["blocking_errors"]
            if error.get("category") == "Paragraph"
        ]

        # Look specifically for the values deliberately injected
        # into the table cell.
        font_detected = any(
            "Calibri" in str(error.get("actual", ""))
            for error in font_errors
        )

        spacing_detected = any(
            str(error.get("actual", "")).strip() in {"1", "1.0", "1.00"}
            for error in paragraph_errors
        )

        print(f"--- {profile_name} ---")
        print(
            f"  is_valid={result['is_valid']} "
            f"blocking_error_count={result['blocking_error_count']} "
            f"score={result['score']}"
        )

        print(
            "  Font validator detected table-cell "
            f"Calibri 11pt: {font_detected}"
        )

        print(
            "  Paragraph validator detected table-cell "
            f"1.0 spacing: {spacing_detected}"
        )

        if font_errors:
            print("  Font errors:")
            for error in font_errors:
                print(
                    f"    - {error.get('actual')} "
                    f"at {error.get('location')}"
                )

        if paragraph_errors:
            print("  Paragraph errors:")
            for error in paragraph_errors:
                print(
                    f"    - {error.get('actual')} "
                    f"at {error.get('location')}"
                )

        print("  Page/Pagination results intentionally ignored.")
        print()

        return font_detected, spacing_detected

    finally:
        if os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                pass


if __name__ == "__main__":
    print(
        "DIAGNOSTIC: Does the current implementation scan "
        "table-cell content?\n"
    )

    print(
        "This is NOT a pass/fail test. "
        "It establishes current behavior.\n"
    )

    print("=" * 70)

    outcomes = {}

    for profile_name in (
        "asu_graduate",
        "amity_selected",
        "mumbai_engineering",
    ):
        outcomes[profile_name] = run_diagnostic(profile_name)

    print("=" * 70)
    print("\nSUMMARY")

    any_detected = False

    for profile_name, (font_detected, spacing_detected) in outcomes.items():
        print(
            f"  {profile_name}: "
            f"font_detected={font_detected}, "
            f"spacing_detected={spacing_detected}"
        )

        if font_detected or spacing_detected:
            any_detected = True

    print()

    if any_detected:
        print(
            "CONCLUSION: Table-cell content IS scanned by at least "
            "one validator in at least one profile."
        )
    else:
        print(
            "CONCLUSION: Table-cell content is NOT scanned by "
            "font_validator.py or paragraph_validator.py in any "
            "of the three profiles. Deliberately non-compliant "
            "text inside a table cell currently passes silently."
        )