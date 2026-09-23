"""
Manual integration test for validate_document().

Run from the backend/ directory with the venv active:
    python test_validator_manual.py

This is a manual test only (no pytest, no extra dependencies).
It does not modify any existing project files or DOCX files.
"""

import sys
from pathlib import Path

from validator.validator import validate_document

PROFILE_NAME = "asu_graduate"

# Directory the script is run relative to (backend/).
BACKEND_DIR = Path(__file__).resolve().parent

# Which DOCX fixture to validate. Change this to point at a different
# fixture, e.g. "test_data/malformed_asu_sample.docx".
SAMPLE_DOCX = BACKEND_DIR / "test_data" / "malformed_asu_sample.docx"


def print_result(result: dict) -> None:
    """Print the validate_document() result clearly, separating blocking
    errors from informational notices."""
    issues = result.get("errors", [])

    blocking_errors = [issue for issue in issues if issue.get("severity") != "info"]
    informational_notices = [issue for issue in issues if issue.get("severity") == "info"]

    print("Document Validation Result")
    print("---------------------------")
    print(f"University:            {result.get('university')}")
    print(f"Program:               {result.get('program')}")
    print(f"Document type:         {result.get('document_type')}")
    print(f"Profile name:          {result.get('profile_name')}")
    print(f"Is valid:              {result.get('is_valid')}")
    print(f"Total issues:          {len(issues)}")
    print(f"Blocking errors:       {len(blocking_errors)}")
    print(f"Informational notices: {len(informational_notices)}")
    print()

    if not issues:
        print("No validation issues found.")
        return

    for index, issue in enumerate(issues, start=1):
        label = "INFO" if issue.get("severity") == "info" else "ERROR"
        print(f"[{label}] Issue {index}:")
        print(f"  rule:        {issue.get('rule')}")
        print(f"  expected:    {issue.get('expected')}")
        print(f"  actual:      {issue.get('actual')}")
        print(f"  location:    {issue.get('location', 'N/A')}")
        if "occurrences" in issue:
            print(f"  occurrences: {issue.get('occurrences')}")
        print()


def run_validator_integration_test():
    """Validate the configured sample DOCX and print the result."""
    if not SAMPLE_DOCX.exists():
        print(
            "Configured sample DOCX not found.\n"
            f"Expected a file at: {SAMPLE_DOCX}\n"
            "Update SAMPLE_DOCX in this script, or generate/place the "
            "fixture at that path, then re-run this test."
        )
        sys.exit(1)

    print(f"Using sample DOCX: {SAMPLE_DOCX}\n")

    result = validate_document(str(SAMPLE_DOCX), PROFILE_NAME)
    print_result(result)


if __name__ == "__main__":
    run_validator_integration_test()