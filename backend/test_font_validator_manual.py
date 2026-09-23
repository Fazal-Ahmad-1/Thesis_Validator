"""
Manual test script for font_validator.py.

Run from the backend/ directory with the venv active:
    python test_font_validator_manual.py

This is a manual test only (no pytest, no extra dependencies).
It does not modify any existing project files.
"""

from validator.docx_reader import read_docx
from validator.rule_loader import load_rules
from validator.font_validator import validate_font_settings

SAMPLE_DOCX = r"C:\Users\shafqat ali\Downloads\sample.docx"
PROFILE_NAME = "asu_graduate"


def run_font_validation_test():
    """Load the sample document and rules, then print every font violation."""
    document = read_docx(SAMPLE_DOCX)
    rules = load_rules(PROFILE_NAME)

    errors = validate_font_settings(document, rules)

    if not errors:
        print("No font violations found. Document fonts match the allowed rules.")
        return

    print(f"Found {len(errors)} unique font violation(s):\n")

    for index, error in enumerate(errors, start=1):
        print(f"Violation {index}:")
        print(f"  rule:         {error['rule']}")
        print(f"  expected:     {error['expected']}")
        print(f"  actual:       {error['actual']}")
        print(f"  location:     {error['location']}")
        print(f"  text_snippet: {error['text_snippet']}")
        print(f"  occurrences:  {error['occurrences']}")
        print()

    print(f"Total number of unique violations: {len(errors)}")


def run_missing_rules_guard_check():
    """Confirm validate_font_settings raises ValueError when 'rules' is missing."""
    document = read_docx(SAMPLE_DOCX)

    print("\nGuard check: calling validate_font_settings with no 'rules' key...")
    try:
        validate_font_settings(document, {"university": "Test University"})
        print("UNEXPECTED: No error was raised. This should not happen.")
    except ValueError as e:
        print(f"Guard check passed. Raised ValueError as expected: {e}")


if __name__ == "__main__":
    run_font_validation_test()
    run_missing_rules_guard_check()