def validate_page_settings(document, rules):
    """
    Check page size and margins against university rules
    for every section in the document.
    """

    errors = []

    # Formatting rules are nested under the top-level "rules" key,
    # alongside metadata like university/program/document_type.
    if "rules" not in rules:
        raise ValueError("Rule profile is missing the 'rules' section.")

    if "page" not in rules["rules"]:
        raise ValueError("Rule profile is missing the 'rules.page' section.")

    page_rules = rules["rules"]["page"]

    # ---------------------------------------------------------
    # Check page size requirements
    # ---------------------------------------------------------
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

    page_tolerance = 0.05

    # ---------------------------------------------------------
    # Convert margin rules to inches
    # ---------------------------------------------------------
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

    margin_tolerance = 0.03

    # ---------------------------------------------------------
    # Validate every document section
    # ---------------------------------------------------------
    for section_number, section in enumerate(document.sections, start=1):

        # Convert Word measurements to inches
        page_width = section.page_width.inches
        page_height = section.page_height.inches

        top_margin = section.top_margin.inches
        bottom_margin = section.bottom_margin.inches
        left_margin = section.left_margin.inches
        right_margin = section.right_margin.inches

        # -----------------------------------------------------
        # Page size
        # -----------------------------------------------------
        if expected_width is not None:
            if (
                abs(page_width - expected_width) > page_tolerance
                or abs(page_height - expected_height) > page_tolerance
            ):
                errors.append({
                    "category": "Page",
                    "type": "Formatting",
                    "rule": "Page Size",
                    "severity": "error",
                    "message": (
                        f"Section {section_number}: "
                        "The page size does not match the selected "
                        "formatting requirements."
                    ),
                    "expected": expected_size,
                    "actual": (
                        f"{page_width:.2f} x "
                        f"{page_height:.2f} inches"
                    ),
                    "suggestion": (
                        f"Change the page size of section "
                        f"{section_number} to {expected_size} in Word."
                    ),
                })

        # -----------------------------------------------------
        # Margins
        # -----------------------------------------------------
        margins = {
            "Top Margin": (top_margin, expected_top),
            "Bottom Margin": (bottom_margin, expected_bottom),
            "Left Margin": (left_margin, expected_left),
            "Right Margin": (right_margin, expected_right),
        }

        for name, (actual, expected) in margins.items():
            if abs(actual - expected) > margin_tolerance:
                errors.append({
                    "category": "Page",
                    "type": "Formatting",
                    "rule": name,
                    "severity": "error",
                    "message": (
                        f"Section {section_number}: "
                        f"The {name.lower()} does not match "
                        "the selected formatting requirements."
                    ),
                    "expected": f"{expected:.2f} inches",
                    "actual": f"{actual:.2f} inches",
                    "suggestion": (
                        f"Change the {name.lower()} of section "
                        f"{section_number} to {expected:.2f} inches in Word."
                    ),
                })

    return errors