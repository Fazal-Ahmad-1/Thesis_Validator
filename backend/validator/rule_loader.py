"""
Loads university formatting rule files (JSON) from backend/rules/.
"""

import json
import re
from pathlib import Path

# backend/validator/rule_loader.py -> parents[1] is backend/
RULES_DIR = Path(__file__).resolve().parents[1] / "rules"

# Only allow simple names: letters, digits, underscores, hyphens.
# This blocks path separators, "..", absolute paths, and similar tricks.
VALID_PROFILE_NAME = re.compile(r"^[A-Za-z0-9_-]+$")


def load_rules(profile_name: str) -> dict:
    """
    Load a university rule profile as a dictionary.

    Args:
        profile_name: Rule profile identifier, e.g. "asu_graduate".
            Do not include the .json extension. Must contain only
            letters, digits, underscores, and hyphens.

    Returns:
        The parsed rule set as a Python dictionary.

    Raises:
        ValueError: If profile_name has an invalid format (e.g. path
            separators or ".."), or if the rule file contains invalid JSON.
        FileNotFoundError: If no matching rule file exists.
    """
    if not VALID_PROFILE_NAME.match(profile_name):
        raise ValueError(
            f"Invalid rule profile name: '{profile_name}'. "
            "Only letters, digits, underscores, and hyphens are allowed."
        )

    rule_file = (RULES_DIR / f"{profile_name}.json").resolve()

    # Defense in depth: make sure the resolved path is still inside RULES_DIR.
    if RULES_DIR.resolve() not in rule_file.parents:
        raise ValueError(f"Invalid rule profile name: '{profile_name}'.")

    if not rule_file.exists():
        raise FileNotFoundError(
            f"Rule profile '{profile_name}' not found at: {rule_file}"
        )

    try:
        with rule_file.open("r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"Rule file '{rule_file.name}' contains invalid JSON: {exc}"
        ) from exc