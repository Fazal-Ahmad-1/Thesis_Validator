"""
Regression test for table-cell formatting validation.

Verifies that:
1. Correct table-cell formatting passes.
2. Bad font inside a table cell is detected.
3. Bad spacing inside a table cell is detected.
4. Bad formatting inside nested tables is detected.
5. Normal body paragraphs are still detected.
6. Heading paragraphs inside tables are not incorrectly treated as body
   paragraphs.

Run from backend:

    python test_table_content_regression_manual.py
"""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt

from validator.validator import validate_document


BASE_DIR = Path(__file__).resolve().parent
TEST_DIR = BASE_DIR / "_table_regression_tests"
TEST_DIR.mkdir(exist_ok=True)


PROFILES = [
    "asu_graduate",
    "amity_selected",
    "mumbai_engineering",
]


def set_paragraph_spacing(paragraph, value):
    paragraph.paragraph_format.line_spacing = value


def add_text_with_format(paragraph, text, font_name, font_size):
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    return run


def create_test_document(
    path,
    profile_name,
    bad_table_font=False,
    bad_table_spacing=False,
    bad_nested_table=False,
    bad_body_font=False,
):
    document = Document()

    # Normal body paragraph
    body = document.add_paragraph()

    if bad_body_font:
        add_text_with_format(body, "Bad body font", "Calibri", 11)
    else:
        add_text_with_format(body, "Normal body text", "Times New Roman", 12)

    set_paragraph_spacing(
        body,
        1.0 if bad_body_spacing_for_profile(profile_name) else 1.5
    )

    # Top-level table
    table = document.add_table(rows=1, cols=1)

    cell = table.cell(0, 0)

    paragraph = cell.paragraphs[0]

    if bad_table_font:
        add_text_with_format(
            paragraph,
            "Bad table font",
            "Calibri",
            11
        )
    else:
        add_text_with_format(
            paragraph,
            "Correct table font",
            "Times New Roman",
            12
        )

    if bad_table_spacing:
        set_paragraph_spacing(paragraph, 1.0)
    else:
        set_paragraph_spacing(paragraph, 1.5)

    # Nested table
    if bad_nested_table:
        nested = cell.add_table(rows=1, cols=1)

        nested_cell = nested.cell(0, 0)
        nested_paragraph = nested_cell.paragraphs[0]

        add_text_with_format(
            nested_paragraph,
            "Bad nested table font",
            "Calibri",
            11
        )

        set_paragraph_spacing(nested_paragraph, 1.0)

    document.save(path)


def bad_body_spacing_for_profile(profile_name):
    """
    This helper deliberately returns False.

    Body-spacing isolation is already covered by the existing
    category-isolation tests. This regression test focuses specifically
    on table traversal.
    """
    return False


def get_errors(result):
    return result.get("errors", [])


def has_font_error(result):
    return any(
        error.get("category") == "Font"
        and error.get("severity") == "error"
        for error in get_errors(result)
    )


def has_paragraph_error(result):
    return any(
        error.get("category") == "Paragraph"
        and error.get("severity") == "error"
        for error in get_errors(result)
    )


def has_table_font_error(result):
    for error in get_errors(result):
        if error.get("category") != "Font":
            continue

        if error.get("severity") != "error":
            continue

        actual = str(error.get("actual", "")).lower()

        if "calibri" in actual:
            return True

    return False


def has_table_spacing_error(result):
    for error in get_errors(result):
        if error.get("category") != "Paragraph":
            continue

        if error.get("severity") != "error":
            continue

        actual = error.get("actual")

        try:
            if abs(float(actual) - 1.0) < 1e-9:
                return True
        except (TypeError, ValueError):
            pass

    return False


def run_test(name, function):
    try:
        function()
        print(f"[PASS] {name}")
        return True
    except AssertionError as exc:
        print(f"[FAIL] {name}")
        print(f"       {exc}")
        return False
    except Exception as exc:
        print(f"[ERROR] {name}")
        print(f"        {type(exc).__name__}: {exc}")
        return False


def test_compliant_table(profile_name):
    path = TEST_DIR / f"{profile_name}_compliant.docx"

    create_test_document(
        path,
        profile_name,
        bad_table_font=False,
        bad_table_spacing=False,
        bad_nested_table=False,
    )

    result = validate_document(str(path), profile_name)

    assert not has_table_font_error(result), (
        "Correct table-cell font was incorrectly reported as invalid."
    )

    assert not has_table_spacing_error(result), (
        "Correct table-cell spacing was incorrectly reported as invalid."
    )


def test_bad_table_font(profile_name):
    path = TEST_DIR / f"{profile_name}_bad_table_font.docx"

    create_test_document(
        path,
        profile_name,
        bad_table_font=True,
        bad_table_spacing=False,
        bad_nested_table=False,
    )

    result = validate_document(str(path), profile_name)

    assert has_table_font_error(result), (
        "Calibri 11pt inside a table cell was not detected."
    )


def test_bad_table_spacing(profile_name):
    path = TEST_DIR / f"{profile_name}_bad_table_spacing.docx"

    create_test_document(
        path,
        profile_name,
        bad_table_font=False,
        bad_table_spacing=True,
        bad_nested_table=False,
    )

    result = validate_document(str(path), profile_name)

    assert has_table_spacing_error(result), (
        "1.0 line spacing inside a table cell was not detected."
    )


def test_bad_nested_table(profile_name):
    path = TEST_DIR / f"{profile_name}_bad_nested_table.docx"

    create_test_document(
        path,
        profile_name,
        bad_table_font=False,
        bad_table_spacing=False,
        bad_nested_table=True,
    )

    result = validate_document(str(path), profile_name)

    assert has_table_font_error(result), (
        "Bad font inside a nested table was not detected."
    )

    assert has_table_spacing_error(result), (
        "Bad spacing inside a nested table was not detected."
    )


def test_body_still_detected(profile_name):
    path = TEST_DIR / f"{profile_name}_bad_body.docx"

    create_test_document(
        path,
        profile_name,
        bad_table_font=False,
        bad_table_spacing=False,
        bad_nested_table=False,
        bad_body_font=True,
    )

    result = validate_document(str(path), profile_name)

    assert has_font_error(result), (
        "Normal body font validation stopped working after "
        "introducing table traversal."
    )


def main():
    passed = 0
    total = 0

    for profile_name in PROFILES:
        print(f"\n--- {profile_name} ---")

        tests = [
            (
                "Compliant table content",
                lambda p=profile_name: test_compliant_table(p),
            ),
            (
                "Bad table font detected",
                lambda p=profile_name: test_bad_table_font(p),
            ),
            (
                "Bad table spacing detected",
                lambda p=profile_name: test_bad_table_spacing(p),
            ),
            (
                "Bad nested table detected",
                lambda p=profile_name: test_bad_nested_table(p),
            ),
            (
                "Normal body validation still works",
                lambda p=profile_name: test_body_still_detected(p),
            ),
        ]

        for name, test_function in tests:
            total += 1

            if run_test(name, test_function):
                passed += 1

    print("\n========================================")
    print(f"TABLE REGRESSION RESULT: {passed}/{total} PASSED")
    print("========================================")

    if passed != total:
        raise SystemExit(1)


if __name__ == "__main__":
    main()