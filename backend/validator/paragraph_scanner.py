"""
Shared paragraph traversal helper.

Yields paragraphs from:
1. Top-level document body
2. Top-level tables
3. Nested tables inside table cells

Paragraphs are deduplicated by their underlying XML element so that
merged table cells cannot cause the same paragraph to be validated twice.
"""


def _iter_cell_paragraphs(cell, seen):
    """
    Yield unique paragraphs directly inside a table cell and
    recursively traverse any nested tables.
    """

    for paragraph in cell.paragraphs:
        paragraph_id = id(paragraph._p)

        if paragraph_id not in seen:
            seen.add(paragraph_id)
            yield paragraph

    for nested_table in cell.tables:
        yield from _iter_table_paragraphs(nested_table, seen)


def _iter_table_paragraphs(table, seen):
    """
    Recursively yield unique paragraphs from a table,
    including paragraphs inside nested tables.
    """

    for row in table.rows:
        for cell in row.cells:
            yield from _iter_cell_paragraphs(cell, seen)


def iter_all_paragraphs(document):
    """
    Yield every unique paragraph that should be inspected by
    body-level formatting validators.

    Includes:
    - Normal document paragraphs
    - Paragraphs inside tables
    - Paragraphs inside nested tables
    """

    seen = set()

    # Normal document paragraphs
    for paragraph in document.paragraphs:
        paragraph_id = id(paragraph._p)

        if paragraph_id not in seen:
            seen.add(paragraph_id)
            yield paragraph

    # Top-level tables and their nested tables
    for table in document.tables:
        yield from _iter_table_paragraphs(table, seen)