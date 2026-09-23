def validate_page_settings(document, rules):
    """
    Check document page size and margins against university rules.
    """

    errors = []

    section = document.sections[0]

    # Convert Word measurements from EMU to inches
    page_width = section.page_width.inches
    page_height = section.page_height.inches

    top_margin = section.top_margin.inches
    bottom_margin = section.bottom_margin.inches
    left_margin = section.left_margin.inches
    right_margin = section.right_margin.inches

    # Formatting rules are nested under the top-level "rules" key,
    # alongside metadata like university/program/document_type.
    if "rules" not in rules:
        raise ValueError("Rule profile is missing the 'rules' section.")

    if "page" not in rules["rules"]:
        raise ValueError("Rule profile is missing the 'rules.page' section.")

    page_rules = rules["rules"]["page"]

    # Check page size
    expected_size = page_rules["size"]

    if expected_size == "A4":
        expected_width = 8.27
        expected_height = 11.69
    elif expected_size == "Letter":
        expected_width = 8.5
        expected_height = 11.0
    else:
        expected_width = None
        expected_height = None

    if expected_width is not None:
        tolerance = 0.05

        if (
            abs(page_width - expected_width) > tolerance
            or abs(page_height - expected_height) > tolerance
        ):
            errors.append({
                "category": "Page",
                "type": "Formatting",
                "rule": "Page Size",
                "severity": "error",
                "message": "The page size does not match the selected formatting requirements.",
                "expected": expected_size,
                "actual": f"{page_width:.2f} x {page_height:.2f} inches",
                "suggestion": f"Change the document page size to {expected_size} in Word.",
            })

    # Convert margin rules to inches
    margin_rules = page_rules["margins"]
    unit = margin_rules["unit"]

    if unit == "cm":
        conversion = 1 / 2.54
    elif unit == "mm":
        conversion = 1 / 25.4
    else:
        conversion = 1

    expected_top = margin_rules["top"] * conversion
    expected_bottom = margin_rules["bottom"] * conversion
    expected_left = margin_rules["left"] * conversion
    expected_right = margin_rules["right"] * conversion

    # Check margins
    margin_tolerance = 0.03

    margins = {
        "Top Margin": (top_margin, expected_top),
        "Bottom Margin": (bottom_margin, expected_bottom),
        "Left Margin": (left_margin, expected_left),
        "Right Margin": (right_margin, expected_right)
    }

    for name, (actual, expected) in margins.items():
        if abs(actual - expected) > margin_tolerance:
            errors.append({
                "category": "Page",
                "type": "Formatting",
                "rule": name,
                "severity": "error",
                "message": f"The {name.lower()} does not match the selected formatting requirements.",
                "expected": f"{expected:.2f} inches",
                "actual": f"{actual:.2f} inches",
                "suggestion": f"Change the {name.lower()} to {expected:.2f} inches in Word.",
            })

    return errors