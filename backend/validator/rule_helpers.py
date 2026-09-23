"""Small helpers shared by validators when interpreting rule profiles."""


def get_configured_heading_style_names(rules: dict) -> set[str]:
    """Return exact Word style names explicitly configured for headings."""
    heading_rules = rules.get("rules", {}).get("headings", {})
    if not isinstance(heading_rules, dict):
        return set()

    names = set()
    for level_rule in heading_rules.values():
        if not isinstance(level_rule, dict):
            continue
        style_name = level_rule.get("style")
        if isinstance(style_name, str) and style_name.strip():
            names.add(style_name.strip())
    return names
