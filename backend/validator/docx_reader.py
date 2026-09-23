import zipfile
from pathlib import Path

from docx import Document
from docx.opc.exceptions import PackageNotFoundError


def read_docx(file_path: str) -> Document:
    """
    Read and return a DOCX document.

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the file is not a .docx or is not a valid DOCX.
    """
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"DOCX file not found: {path}")

    if path.suffix.lower() != ".docx":
        raise ValueError(f"Expected a .docx file, got: {path.suffix}")

    try:
        return Document(path)

    except (PackageNotFoundError, zipfile.BadZipFile) as exc:
        # A .docx file is a ZIP archive under the hood.
        # python-docx raises PackageNotFoundError for some malformed
        # packages, while a file that isn't a ZIP archive at all
        # raises zipfile.BadZipFile.
        #
        # Both cases mean the uploaded file is not a valid DOCX.
        raise ValueError(
            "The uploaded file is not a valid DOCX document."
        ) from exc