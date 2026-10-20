#!/usr/bin/env python3

import argparse
import bz2
import gzip
import lzma
import re
import sys
from pathlib import Path
from urllib.parse import quote
import tempfile
import subprocess
import logging
import warnings
import xml.etree.ElementTree as ET
import zipfile
import shutil
import tarfile
from typing import BinaryIO

SUPPORTED_EXTENSIONS = {
    ".txt",
    ".md",
    ".pdf",
    ".doc",
    ".docx",
    ".odt",
    ".odp",
    ".ppt",
    ".pptx",
    ".xlsx",
    ".xls",
    ".ods",
}

DEFAULT_EXTENSIONS = {
    ".txt",
    ".md",
    ".pdf",
    ".docx",
    ".odt",
    ".odp",
    ".pptx",
}

ARCHIVE_SUFFIXES = (
    ".tar.gz",
    ".tar.bz2",
    ".tar.xz",
    ".tgz",
    ".tbz2",
    ".txz",
    ".zip",
    ".jar",
    ".7z",
    ".tar",
    ".gz",
    ".bz2",
    ".xz",
)

VALID_FORMATS = {extension.lstrip(".") for extension in SUPPORTED_EXTENSIONS}

ODF_TEXT_NAMESPACE = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"


def configure_logging():
    """Reduce warning noise from PDF extraction dependencies.

    The command-line output is intended for search results, so warnings emitted
    by pypdf are suppressed while errors remain visible.
    """
    # Keep third-party parser warnings from obscuring search output.
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    warnings.filterwarnings("ignore", module=r"pypdf(\.|$)")


def get_allowed_extensions(extension_spec: str | None) -> set[str]:
    """Normalize an extension option, including additive and extensionless forms.

    Args:
        extension_spec: Comma-separated extensions. A leading `+` includes the
            default extensions, and empty entries select files without an
            extension. `None` selects only the defaults.

    Returns:
        Lowercase extensions with leading dots, including an empty string when
        extensionless files are selected.
    """
    if extension_spec is None:
        return DEFAULT_EXTENSIONS

    include_defaults = extension_spec.startswith("+")
    if include_defaults:
        extension_spec = extension_spec[1:]

    allowed_extensions = set(DEFAULT_EXTENSIONS) if include_defaults else set()
    for extension in extension_spec.split(","):
        extension = extension.strip().lower().lstrip(".")
        allowed_extensions.add(f".{extension}" if extension else "")

    return allowed_extensions


def configure_colors(mode: str) -> tuple[str, str]:
    """Choose ANSI color codes based on the requested color mode.

    Args:
        mode: One of `"always"`, `"never"`, or `"auto"`. Auto mode
            enables colors only when standard output is connected to a TTY.

    Returns:
        A pair containing the color-start and reset escape sequences, or two
        empty strings when color output is disabled.
    """
    # Auto mode avoids emitting terminal escapes when output is redirected.
    use_color = mode == "always" or (mode == "auto" and sys.stdout.isatty())

    if use_color:
        return "\033[31m", "\033[0m"

    return "", ""


def hyperlink(path: Path) -> str:
    """Format a file path as a terminal OSC 8 hyperlink.

    Args:
        path: File path to encode as a clickable terminal link.

    Returns:
        The OSC 8 hyperlink prefix and resolved file URI.
    """
    # Resolve first so the URI identifies the file independently of the cwd.
    uri = path.resolve().as_uri()
    return f"====> {uri} "


def extract_pdf_text_pypdf(path: Path) -> list[str]:
    """Extract page text from a PDF, marking each page in the result.

    Args:
        path: PDF file to read.

    Returns:
        Extracted text as lines, with a page marker before each page's text.

    Raises:
        RuntimeError: If the pypdf dependency is unavailable.
    """
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError("PDF support requires: python -m pip install pypdf") from exc

    lines = []

    reader = PdfReader(str(path))

    for page_number, page in enumerate(reader.pages, start=1):
        # Keep the PDF page number available in the extracted text.
        lines.append(f"[PDF page {page_number}]")

        text = page.extract_text() or ""
        lines.extend(text.splitlines())

    return lines


def extract_pdf_text_pymupdf(path: Path) -> list[str]:
    """Extract text from a PDF, marking each page in the result.

    Args:
        path: PDF file to read.

    Returns:
        Extracted text as lines, with a page marker before each page's text.

    Raises:
        RuntimeError: If the pypdf dependency is unavailable.
    """
    try:
        import pymupdf
    except ImportError as exc:
        raise RuntimeError(
            "PDF support requires: python -m pip install pymupdf"
        ) from exc

    lines = []

    with pymupdf.open(path) as doc:
        for page_number, page in enumerate(doc, start=1):
            lines.append(f"[PDF page {page_number}]")

            text = page.get_text("text") or ""
            lines.extend(text.splitlines())

    return lines


def extract_pdf_text(path: Path) -> list[str]:
    """Extract PDF text, preferring PyMuPDF and falling back to pypdf.

    Args:
        path: PDF file to read.

    Returns:
        Extracted text as lines, with a page marker before each page's text.
    """
    try:
        return extract_pdf_text_pymupdf(path)
    except Exception:
        return extract_pdf_text_pypdf(path)


def extract_doc_text(path: Path) -> list[str]:
    """Extract lines from a legacy Word `.doc` file using antiword.

    Args:
        path: Legacy Word document to read.

    Returns:
        The converted document text split into lines.

    Raises:
        RuntimeError: If antiword is missing or cannot process the document.
    """
    try:
        # antiword provides text conversion for the binary .doc format.
        result = subprocess.run(
            ["antiword", str(path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=True,
        )
    except FileNotFoundError as exc:
        raise RuntimeError(
            "DOC support requires antiword. "
            "Install it with: sudo apt install antiword"
        ) from exc
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(
            f"antiword could not extract text: {exc.stderr.strip()}"
        ) from exc

    return result.stdout.splitlines()


def extract_docx_text(path: Path) -> list[str]:
    """Extract paragraph text from a Word `.docx` document.

    Args:
        path: Word document to read.

    Returns:
        Paragraph contents split into lines; empty paragraphs are preserved.

    Raises:
        RuntimeError: If python-docx is unavailable.
        Exception: If the document cannot be opened or parsed.
    """
    try:
        from docx import Document
    except ImportError as exc:
        raise RuntimeError(
            "DOCX support requires: " "python -m pip install python-docx"
        ) from exc

    # Preserve blank paragraphs so extracted line positions reflect the document.
    document = Document(str(path))
    lines = []

    for paragraph in document.paragraphs:
        lines.extend(paragraph.text.splitlines() or [""])

    return lines


def extract_pptx_text(path: Path) -> list[str]:
    """Extract visible text, tables, and speaker notes from a presentation.

    Args:
        path: PowerPoint `.pptx` file to read.

    Returns:
        Extracted slide text as lines, preceded by slide and notes markers.

    Raises:
        RuntimeError: If python-pptx is unavailable.
    """
    try:
        from pptx import Presentation
    except ImportError as exc:
        raise RuntimeError(
            "PowerPoint support requires: " "python -m pip install python-pptx"
        ) from exc

    presentation = Presentation(str(path))
    lines = []

    def extract_shape(shape):
        """Collect text recursively from one shape and its descendants.

        Args:
            shape: PowerPoint shape to inspect.

        Returns:
            Non-empty text paragraphs and table rows found in the shape.
        """
        extracted = []

        # Read both standard text frames and tabular content on the shape.
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                text = paragraph.text.strip()
                if text:
                    extracted.append(text)

        if shape.has_table:
            for row in shape.table.rows:
                cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if cells:
                    extracted.append(" | ".join(cells))

        # Shape type 6 represents a group of shapes.
        if shape.shape_type == 6:
            for child in shape.shapes:
                extracted.extend(extract_shape(child))

        return extracted

    for slide_number, slide in enumerate(presentation.slides, start=1):
        # Extract visible slide text.
        lines.append(f"[Slide {slide_number}]")

        for shape in slide.shapes:
            lines.extend(extract_shape(shape))

        # Extract speaker notes, when present.
        if slide.has_notes_slide:
            notes_slide = slide.notes_slide

            # notes_text_frame contains the actual notes text,
            # excluding slide images and slide-number placeholders.
            notes_frame = notes_slide.notes_text_frame

            if notes_frame is not None:
                notes_text = notes_frame.text.strip()

                if notes_text:
                    lines.append(f"[Speaker Notes - Slide {slide_number}]")
                    lines.extend(notes_text.splitlines())

    return lines


def extract_ppt_text(path: Path) -> list[str]:
    """Convert a legacy `.ppt` presentation and extract its text.

    LibreOffice performs the conversion because python-pptx reads only the
    newer `.pptx` format.

    Args:
        path: Legacy PowerPoint presentation to convert.

    Returns:
        Extracted slide text and speaker notes as lines.

    Raises:
        RuntimeError: If LibreOffice is missing, conversion fails, or times out.
    """
    try:
        with tempfile.TemporaryDirectory(prefix="docgrep_") as temp_dir:
            # Store the conversion in a temporary directory to avoid altering input files.
            result = subprocess.run(
                [
                    "libreoffice",
                    "--headless",
                    "--convert-to",
                    "pptx",
                    "--outdir",
                    temp_dir,
                    str(path.resolve()),
                ],
                capture_output=True,
                text=True,
                timeout=120,
            )

            converted = Path(temp_dir) / (path.stem + ".pptx")

            if result.returncode != 0 or not converted.is_file():
                details = result.stderr.strip() or result.stdout.strip()

                raise RuntimeError(
                    "LibreOffice could not convert the PPT file. " + details
                )

            return extract_pptx_text(converted)

    except FileNotFoundError as exc:
        raise RuntimeError(
            "Legacy PPT support requires LibreOffice. "
            "Install it with: sudo apt install libreoffice-impress"
        ) from exc
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(
            "LibreOffice conversion timed out after 120 seconds."
        ) from exc


def extract_odf_text(path: Path) -> list[str]:
    """Extract paragraphs and headings from an OpenDocument text file.

    Args:
        path: OpenDocument Text or Presentation file containing `content.xml`.

    Returns:
        Paragraph and heading text split into lines.

    Raises:
        RuntimeError: If the archive does not contain `content.xml`.
    """
    # ODF text documents are ZIP containers whose document content is in XML.
    with zipfile.ZipFile(path) as archive:
        try:
            content = archive.read("content.xml")
        except KeyError as exc:
            raise RuntimeError(
                "OpenDocument file does not contain content.xml"
            ) from exc

    root = ET.fromstring(content)
    lines = []

    def extract_element_text(element) -> str:
        """Convert an XML element subtree, including ODF spacing markers, to text.

        Args:
            element: XML element whose child content should be serialized.

        Returns:
            Text represented by the element, preserving spaces, tabs, and breaks.
        """
        extracted = [element.text or ""]

        for child in element:
            if child.tag == f"{ODF_TEXT_NAMESPACE}s":
                extracted.append(" " * int(child.get(f"{ODF_TEXT_NAMESPACE}c", "1")))
            elif child.tag == f"{ODF_TEXT_NAMESPACE}tab":
                extracted.append("\t")
            elif child.tag == f"{ODF_TEXT_NAMESPACE}line-break":
                extracted.append("\n")
            else:
                extracted.append(extract_element_text(child))

            extracted.append(child.tail or "")

        return "".join(extracted)

    paragraph_tags = {
        f"{ODF_TEXT_NAMESPACE}p",
        f"{ODF_TEXT_NAMESPACE}h",
    }
    for element in root.iter():
        if element.tag in paragraph_tags:
            lines.extend(extract_element_text(element).splitlines() or [""])

    return lines


def extract_xlsx_text(path: Path, separator: str = "|") -> list[str]:
    """Extract worksheet rows from an Excel `.xlsx` workbook.

    Each output line begins with the sheet name, followed by cell values joined
    with `separator`.

    Args:
        path: Workbook to read.
        separator: String used between the sheet name and cell values.

    Returns:
        One delimited string for each worksheet row.

    Raises:
        RuntimeError: If openpyxl is unavailable.
    """
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError(
            "Excel support requires: python -m pip install openpyxl"
        ) from exc

    # Read cached cell results in streaming mode to limit memory use.
    workbook = load_workbook(filename=str(path), read_only=True, data_only=True)
    lines = []

    try:
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows(values_only=True):
                values = ["" if value is None else str(value) for value in row]
                lines.append(separator.join([sheet.title, *values]))
    finally:
        workbook.close()

    return lines


def extract_xls_text(path: Path, separator: str = "|") -> list[str]:
    """Extract rows and worksheet names from a legacy Excel `.xls` workbook.

    Args:
        path: Workbook to read.
        separator: String used between the sheet name and cell values.

    Returns:
        One delimited string for each worksheet row, with values normalized
        according to their Excel cell type.

    Raises:
        RuntimeError: If xlrd is unavailable.
    """
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError(
            "Legacy Excel support requires: python -m pip install xlrd"
        ) from exc

    # Open workbook. Note that Cell types need normalization
    # because dates and booleans are not plain strings.
    workbook = xlrd.open_workbook(filename=str(path))
    lines = []

    for sheet in workbook.sheets():
        for row_idx in range(sheet.nrows):
            values = []

            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)

                if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                    value = ""
                elif cell.ctype == xlrd.XL_CELL_DATE:
                    value = xlrd.xldate_as_datetime(
                        cell.value, workbook.datemode
                    ).isoformat()
                elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                    value = str(bool(cell.value)).lower()
                elif cell.ctype == xlrd.XL_CELL_ERROR:
                    value = xlrd.error_text_from_code.get(cell.value, str(cell.value))
                else:
                    value = str(cell.value)

                values.append(value)

            lines.append(separator.join([sheet.name, *values]))

    return lines


def extract_ods_text(path: Path, separator: str = "|") -> list[str]:
    """Extract worksheet names and cell contents from an ODS spreadsheet.

    Visible paragraph text takes precedence; typed cell attributes provide a
    fallback for values without textual paragraphs.

    Args:
        path: OpenDocument Spreadsheet file to read.
        separator: String used between the sheet name and cell values.

    Returns:
        One delimited string for each worksheet row.

    Raises:
        RuntimeError: If odfpy is unavailable.
    """
    try:
        from odf.opendocument import load
        from odf.table import Table, TableRow, TableCell, CoveredTableCell
        from odf.text import P, Span, S, Tab, LineBreak
    except ImportError as exc:
        raise RuntimeError("ODS support requires: python -m pip install odfpy") from exc

    # Load once, then serialize each row into the common line-based search format.
    document = load(str(path))
    lines = []

    def extract_paragraph_text(paragraph) -> str:
        """Convert an ODF paragraph and its inline nodes to text.

        Args:
            paragraph: ODF paragraph element to serialize.

        Returns:
            Paragraph text with explicit spaces, tabs, and line breaks retained.
        """
        parts = []

        def visit(node):
            """Append a node's text and recursively visit nested elements.

            Args:
                node: ODF DOM node to serialize into the enclosing result.
            """
            if node.nodeType == node.TEXT_NODE:
                parts.append(node.data)
            elif node.nodeType == node.ELEMENT_NODE:
                if node.tagName == "text:s":
                    count = int(node.getAttribute("c") or 1)
                    parts.append(" " * count)
                elif node.tagName == "text:tab":
                    parts.append("\t")
                elif node.tagName == "text:line-break":
                    parts.append("\n")
                else:
                    for child in node.childNodes:
                        visit(child)

        for child in paragraph.childNodes:
            visit(child)

        return "".join(parts)

    def extract_cell_value(cell) -> str:
        """Return visible paragraph text or the cell's typed stored value.

        Args:
            cell: ODF table cell to inspect.

        Returns:
            Cell text or a string representation of its stored value.
        """
        paragraphs = cell.getElementsByType(P)
        text = "\n".join(extract_paragraph_text(paragraph) for paragraph in paragraphs)

        # Prefer visible text when present.
        if text:
            return text

        # Otherwise, inspect the ODS cell's typed value attributes.
        value_type = cell.getAttribute("valuetype")

        if value_type == "string":
            return cell.getAttribute("stringvalue") or ""

        if value_type == "float":
            return cell.getAttribute("value") or ""

        if value_type == "currency":
            return cell.getAttribute("value") or ""

        if value_type == "percentage":
            return cell.getAttribute("value") or ""

        if value_type == "boolean":
            return cell.getAttribute("booleanvalue") or ""

        if value_type == "date":
            return cell.getAttribute("datevalue") or ""

        if value_type == "time":
            return cell.getAttribute("timevalue") or ""

        # Fallback for files with missing or unexpected value types.
        for attribute in (
            "stringvalue",
            "value",
            "datevalue",
            "timevalue",
            "booleanvalue",
        ):
            value = cell.getAttribute(attribute)
            if value is not None:
                return str(value)

        return ""

    for sheet in document.spreadsheet.getElementsByType(Table):
        sheet_name = sheet.getAttribute("name")

        for row in sheet.getElementsByType(TableRow):
            values = []

            for cell in row.childNodes:
                if cell.nodeType != cell.ELEMENT_NODE:
                    continue

                if cell.qname not in (TableCell().qname, CoveredTableCell().qname):
                    continue

                # Covered cells represent merged cells and have no value.
                if cell.qname == CoveredTableCell().qname:
                    values.append("")
                    continue

                values.append(extract_cell_value(cell))

                # Account for repeated columns in ODS files.
                repeat = int(cell.getAttribute("numbercolumnsrepeated") or 1)
                values.extend([""] * (repeat - 1))

            lines.append(separator.join([sheet_name, *values]))

    return lines


def extract_text(path: Path) -> list[str]:
    """Dispatch a file to the extractor for its extension.

    Supported office and PDF formats use dedicated extractors; other files are
    decoded as UTF-8 text with undecodable bytes replaced.

    Args:
        path: File whose text should be extracted.

    Returns:
        Extracted content as a list of lines.
    """
    # Normalize the suffix so extension matching is case-insensitive.
    suffix = path.suffix.lower()

    if suffix == ".pdf":
        return extract_pdf_text(path)

    elif suffix == ".doc":
        return extract_doc_text(path)

    elif suffix == ".docx":
        return extract_docx_text(path)

    elif suffix in {".odt", ".odp"}:
        return extract_odf_text(path)

    elif suffix == ".pptx":
        return extract_pptx_text(path)

    elif suffix == ".xlsx":
        return extract_xlsx_text(path)

    elif suffix == ".xls":
        return extract_xls_text(path)

    elif suffix == ".ods":
        return extract_ods_text(path)

    elif suffix == ".ppt":
        return extract_ppt_text(path)

    else:
        return path.read_text(
            encoding="utf-8",
            errors="replace",
        ).splitlines()

    return []


def collect_files(
    inputs: list[str],
    formats: str | None = None,
    filename_pattern: re.Pattern[str] | None = None,
) -> list[Path]:
    """Find unique files matching extension and optional filename filters.

    Directories are searched recursively. Nonexistent paths are reported to
    standard error and skipped.

    Args:
        inputs: File and directory paths to search.
        formats: Comma-separated extensions. A leading `+` adds to the default
            extensions, and empty entries select files without an extension.
            `None` selects the default extensions.
        filename_pattern: Optional regular expression matched against each
            filename stem.

    Returns:
        Resolved matching file paths in sorted order.
    """

    allowed_extensions = get_allowed_extensions(formats)

    files = set()

    for item in inputs:
        path = Path(item).expanduser()

        if not path.exists():
            print(
                f"Warning: not found: {item}",
                file=sys.stderr,
            )
            continue

        if path.is_file():
            if path.suffix.lower() in allowed_extensions and (
                filename_pattern is None or filename_pattern.search(path.stem)
            ):
                files.add(path.resolve())
            continue

        if path.is_dir():
            for candidate in path.rglob("*"):
                if (
                    candidate.is_file()
                    and candidate.suffix.lower() in allowed_extensions
                    and (
                        filename_pattern is None
                        or filename_pattern.search(candidate.stem)
                    )
                ):
                    files.add(candidate.resolve())

    return sorted(files)


def archive_suffix(filename: str) -> str | None:
    """Identify a supported archive suffix in a filename.

    Compound suffixes such as `.tar.gz` are checked before simple suffixes
    to avoid treating them as generic compressed streams.

    Args:
        filename: Filename or path basename to inspect.

    Returns:
        The matching suffix in lowercase, or `None` when unsupported.
    """
    # ARCHIVE_SUFFIXES is ordered from compound to simple suffixes.
    filename = filename.lower()
    return next(
        (suffix for suffix in ARCHIVE_SUFFIXES if filename.endswith(suffix)),
        None,
    )


def collect_archives(
    inputs: list[str],
    filename_pattern: re.Pattern[str] | None = None,
) -> list[Path]:
    """Find supported archives at the given paths, independently of documents.

    Directories are searched recursively, and the optional filename expression
    is matched against the archive name with its complete archive suffix removed.

    Args:
        inputs: File and directory paths to inspect.
        filename_pattern: Optional regular expression matched against archive
            filename stems.

    Returns:
        Resolved matching archive paths in sorted order.
    """
    # Deduplicate overlapping input paths before returning stable output.
    archives = set()

    for item in inputs:
        path = Path(item).expanduser()

        if not path.exists():
            continue

        candidates = (
            [path] if path.is_file() else path.rglob("*") if path.is_dir() else []
        )
        for candidate in candidates:
            suffix = archive_suffix(candidate.name)
            if not candidate.is_file() or suffix is None:
                continue

            stem = candidate.name[: -len(suffix)]
            if filename_pattern is None or filename_pattern.search(stem):
                archives.add(candidate.resolve())

    return sorted(archives)


def archive_member_names(path: Path) -> tuple[str, list[str]]:
    """List searchable regular-file members and classify the archive container
    based on compression method.

    Args:
        path: Archive path to inspect.

    Returns:
        A pair of container type (`"zip"`, `"tar"`, `"7z"`, or
        `"compressed"`) and member names. Unsupported suffixes return an
        empty type and an empty list.

    Raises:
        RuntimeError: If 7z support is requested without py7zr.
    """
    # Resolve the format from the full suffix before choosing a container reader.
    suffix = archive_suffix(path.name)
    if suffix is None:
        return "", []

    if suffix in {".zip", ".jar"}:
        with zipfile.ZipFile(path) as archive:
            return "zip", [
                info.filename for info in archive.infolist() if not info.is_dir()
            ]

    if suffix in {
        ".tar",
        ".tar.gz",
        ".tar.bz2",
        ".tar.xz",
        ".tgz",
        ".tbz2",
        ".txz",
    }:
        with tarfile.open(path, "r:*") as archive:
            return "tar", [
                member.name for member in archive.getmembers() if member.isfile()
            ]

    if suffix == ".7z":
        try:
            import py7zr
        except ImportError as exc:
            raise RuntimeError(
                "7z support requires: python -m pip install py7zr"
            ) from exc

        with py7zr.SevenZipFile(path, mode="r") as archive:
            return "7z", [
                item.filename for item in archive.list() if not item.is_directory
            ]

    member_name = path.name[: -len(suffix)]
    return "compressed", [member_name]


def copy_archive_member(
    archive_path: Path,
    archive_type: str,
    member_name: str,
    destination: BinaryIO,
) -> None:
    """Stream one archive member into a binary destination.

    The destination is supplied by the caller, allowing extracted data to be
    staged without writing archive-controlled paths to disk.

    Args:
        archive_path: Archive containing the member.
        archive_type: Container kind returned by :func:`archive_member_names`.
        member_name: Exact member path within the archive.
        destination: Writable binary file-like object.

    Raises:
        RuntimeError: If the member cannot be read safely or the format is
            unsupported.
    """
    if archive_type == "zip":
        # Stream ZIP members directly rather than extracting the whole archive.
        with zipfile.ZipFile(archive_path) as archive:
            with archive.open(member_name) as source:
                shutil.copyfileobj(source, destination)
        return

    if archive_type == "tar":
        with tarfile.open(archive_path, "r:*") as archive:
            source = archive.extractfile(member_name)
            if source is None:
                raise RuntimeError(f"Cannot read archive member {member_name}")
            with source:
                shutil.copyfileobj(source, destination)
        return

    if archive_type == "7z":
        import py7zr

        # py7zr extracts to disk, so verify the resolved member remains in the
        # temporary root before opening it.
        with tempfile.TemporaryDirectory(prefix="egrepd_7z_") as temp_dir:
            root = Path(temp_dir).resolve()
            with py7zr.SevenZipFile(archive_path, mode="r") as archive:
                archive.extract(path=root, targets=[member_name])

            extracted = (root / member_name).resolve()
            if not extracted.is_relative_to(root) or not extracted.is_file():
                raise RuntimeError(f"Cannot safely read archive member {member_name}")
            with extracted.open("rb") as source:
                shutil.copyfileobj(source, destination)
        return

    suffix = archive_suffix(archive_path.name)
    if suffix not in {".gz", ".bz2", ".xz"}:
        raise RuntimeError(f"Unsupported archive format: {archive_path}")

    open_compressed = {
        ".gz": gzip.open,
        ".bz2": bz2.open,
        ".xz": lzma.open,
    }[suffix]
    with open_compressed(archive_path, "rb") as source:
        shutil.copyfileobj(source, destination)


def is_searchable_archive_member(
    member_name: str,
    allowed_extensions: set[str],
    filename_pattern: re.Pattern[str] | None,
    compressed_stream: bool = False,
) -> bool:
    """Check whether an archive member passes extension and filename filters.

    A compressed stream has no internal filename, so it is treated as a text
    file for extension filtering.

    Args:
        member_name: Member path as stored in the archive.
        allowed_extensions: Permitted lowercase extensions, including dots.
        filename_pattern: Optional regular expression applied to the member stem.
        compressed_stream: Whether the member represents a single compressed
            stream without a stored filename.

    Returns:
        `True` when the member should be searched.
    """
    # Compressed streams inherit a text extension because their inner name is absent.
    member_path = Path(member_name)
    suffix = member_path.suffix.lower()

    if not suffix and compressed_stream:
        suffix = ".txt"

    return suffix in allowed_extensions and (
        filename_pattern is None or filename_pattern.search(member_path.stem)
    )


def merge_ranges(
    ranges: list[tuple[int, int]],
) -> list[tuple[int, int]]:
    """Combine overlapping or adjacent inclusive integer ranges.

    Args:
        ranges: Start/end pairs to combine; endpoints are treated as inclusive.

    Returns:
        Sorted, non-overlapping ranges. Empty input produces an empty list.
    """
    # Sorting allows each range to be merged against only the previous result.
    merged = []

    for start, end in sorted(ranges):
        if merged and start <= merged[-1][1] + 1:
            merged[-1] = (
                merged[-1][0],
                max(merged[-1][1], end),
            )
        else:
            merged.append((start, end))

    return merged


def truncate_match_line(
    line: str,
    pattern: re.Pattern,
    matchlen: int,
    red: str,
    reset: str,
) -> str:
    """Build excerpts around every match, merging intersecting excerpts.

    Excerpts are at most `matchlen` characters long unless a match itself is
    shorter. Gaps and omitted text are represented by `...`.

    Args:
        line: Full source line containing zero or more matches.
        pattern: Compiled expression used to locate and highlight matches.
        matchlen: Target maximum length of each excerpt.
        red: ANSI escape sequence used to begin match highlighting; empty
            disables highlighting.
        reset: ANSI escape sequence that ends highlighting.

    Returns:
        A concatenation of excerpts, ellipses, and optional ANSI highlights.
    """
    # Center a window on each match, expanding it when the match is longer.
    ranges = []
    for match in pattern.finditer(line):
        window_length = max(matchlen, match.end() - match.start())
        center = (match.start() + match.end()) // 2
        start = max(0, center - window_length // 2)
        end = min(len(line), start + window_length)
        start = max(0, end - window_length)
        ranges.append((start, end))

    excerpts = []
    previous_end = 0
    for start, end in merge_ranges(ranges):
        if start > previous_end:
            excerpts.append("...")

        excerpt = line[start:end]
        excerpts.append(
            pattern.sub(
                lambda match: (
                    f"{red}{match.group(0)}{reset}" if red else match.group(0)
                ),
                excerpt,
            )
        )
        previous_end = end

    if previous_end < len(line):
        excerpts.append("...")

    return "".join(excerpts)


def truncate_context_line(line: str, matchlen: int) -> str:
    """Limit a context-only line to the configured character count.

    Args:
        line: Context line to display.
        matchlen: Maximum number of source characters to retain.

    Returns:
        The original line when short enough, otherwise the beginning
        of the input line plus `...`.
    """
    # Context lines do not need a centered excerpt because they contain no match.
    if len(line) > matchlen:
        line = line[:matchlen] + "..."
    return line


def search_file(
    path: Path,
    pattern: re.Pattern,
    before: int,
    after: int,
    red: str = "",
    reset: str = "",
    matchlen: int | None = None,
    display_name: str | None = None,
    link_path: Path | None = None,
) -> int:
    """Extract, search, and print a file's matching lines and context.

    Read or extraction errors are reported to standard error and treated as
    zero matches so a batch search can continue.

    Args:
        path: File to extract and search.
        pattern: Compiled regular expression to search for.
        before: Number of preceding context lines to show per match.
        after: Number of following context lines to show per match.
        red: Optional ANSI escape sequence to start match highlighting.
        reset: ANSI escape sequence to end match highlighting.
        matchlen: Optional maximum excerpt size for displayed lines.
        display_name: Optional display path, used for archive members.
        link_path: Optional real file path used as the clickable hyperlink
            target when displaying an archive member.

    Returns:
        Number of matching lines found, or zero if extraction fails or there
        are no matches.
    """
    try:
        # Convert every supported input format to the same list-of-lines representation.
        lines = extract_text(path)
    except Exception as exc:
        source_name = (
            f"{link_path}!/{display_name}"
            if link_path is not None and display_name is not None
            else display_name or path
        )
        print(
            f"Warning: cannot read {source_name}: {exc}",
            file=sys.stderr,
        )
        return 0

    # Track matching line indexes separately from the merged display ranges.
    matches = [index for index, line in enumerate(lines) if pattern.search(line)]

    if not matches:
        return 0

    # Merge overlapping context windows to avoid printing the same line twice.
    ranges = merge_ranges(
        [
            (
                max(0, index - before),
                min(len(lines) - 1, index + after),
            )
            for index in matches
        ]
    )

    match_set = set(matches)

    # Clear the progress display before printing the clickable file URI.
    print("\r\033[K", end="", file=sys.stderr)
    if link_path is not None and display_name is not None:
        output_name = f"!/{display_name}"
    else:
        output_name = display_name or ""
    print(f"\r{hyperlink(link_path or path)}{output_name}")

    for start, end in ranges:
        for index in range(start, end + 1):
            line_number = index + 1
            line = lines[index]

            if index in match_set:
                # line contains a match:
                if matchlen is not None:
                    line = truncate_match_line(line, pattern, matchlen, red, reset)
                else:
                    # Highlight every occurrence of the search pattern, if configured.
                    line = pattern.sub(
                        lambda match: (
                            f"{red}{match.group(0)}{reset}" if red else match.group(0)
                        ),
                        line,
                    )

                print(f"{line_number}: {line}")
            else:
                # line does not contain match, just a context line:
                if matchlen is not None:
                    line = truncate_context_line(line, matchlen)
                print(f"{line_number}- {line}")

    # Separator after all matches and context.
    print("---")

    return len(matches)


def show_progress(
    current: int,
    total: int,
    filename: str,
    bar_width: int = 20,
    label: str = "",
) -> None:
    """Write a single-line progress indicator to standard error.

    Args:
        current: Number of completed items.
        total: Total number of items; zero displays a zero-percent bar.
        filename: Current filename or status text shown beside the bar.
        bar_width: Number of characters used for the progress bar.
        label: Optional prefix distinguishing progress categories.
    """
    filename = str(filename)
    terminal_width = shutil.get_terminal_size(fallback=(80, 24)).columns

    # Clamp available filename space to avoid negative padding on narrow terminals.
    fraction = current / total if total else 0
    filled = int(bar_width * fraction)
    bar = "#" * filled + "." * (bar_width - filled)

    prefix = f"\r{label}[{bar}] {fraction:6.1%} "
    available = max(0, terminal_width - len(prefix) - 1)

    # Truncate long filenames.
    if len(filename) > available:
        filename = (
            "..." + filename[-(available - 3) :]
            if available > 3
            else filename[:available]
        )

    message = prefix + filename

    print(message.ljust(terminal_width - 1), end="", file=sys.stderr, flush=True)


def main() -> int:
    """Run the command-line document search and return its process status.

    The command parses search options, discovers matching documents and
    optionally archive members, prints matching lines with context, and reports
    progress on standard error.

    Returns:
        `0` when at least one matching line is found, otherwise `1`.

    Raises:
        SystemExit: For invalid command-line arguments or invalid regular
            expressions, as handled by :mod:`argparse`.
    """
    configure_logging()

    parser = argparse.ArgumentParser(
        description=(
            "Search "
            + ", ".join([i[1:] for i in SUPPORTED_EXTENSIONS])
            + " as well as any other plain-text files "
            + "using regular expressions."
        )
    )

    parser.add_argument(
        "pattern",
        help="Regular expression to search for",
    )

    parser.add_argument(
        "paths",
        nargs="+",
        help="Files or directories to search",
    )

    parser.add_argument(
        "-i",
        "--ignore-case",
        action="store_true",
        help="Ignore case distinctions",
    )

    parser.add_argument(
        "-a",
        "--include-archives",
        action="store_true",
        help="Search supported archive files after searching ordinary files",
    )

    parser.add_argument(
        "-B",
        "--before-context",
        type=int,
        default=0,
        help="Number of context lines before each match",
    )

    parser.add_argument(
        "-A",
        "--after-context",
        type=int,
        default=0,
        help="Number of context lines after each match",
    )

    parser.add_argument(
        "-C",
        "--context",
        type=int,
        default=None,
        help="Number of context lines on both sides",
    )

    parser.add_argument(
        "-l",
        "--matchlen",
        type=int,
        default=None,
        metavar="CHARS",
        help="Limit each match excerpt to CHARS, centered on the match",
    )

    parser.add_argument(
        "-e",
        "--ext",
        type=str,
        default=None,
        metavar="EXTENSIONS",
        help=(
            "Comma-separated extensions (e.g., 'py,R,java'); a leading '+' adds to "
            + "the defaults (e.g., '+py,sh' adds both options), and empty entries (e.g., 'py,') select extensionless files. "
            + "Defaults: "
            + ",".join([i[1:] for i in DEFAULT_EXTENSIONS])
            + "; "
            + ",".join([i[1:] for i in (SUPPORTED_EXTENSIONS - DEFAULT_EXTENSIONS)])
            + " must be explicitly requested."
        ),
    )

    parser.add_argument(
        "-f",
        "--filename",
        metavar="REGEX",
        help="Only search files whose name without its extension matches this regex",
    )

    parser.add_argument(
        "-c",
        "--color",
        choices=["auto", "always", "never"],
        default="auto",
        help=("Colorize matching text: " "auto (default), always, or never."),
    )
    # Parse and normalize command-line options before discovering input files.
    args = parser.parse_args()

    red, reset = configure_colors(args.color)

    before = args.before_context
    after = args.after_context

    # The combined context option overrides the individual before/after values.
    if args.context is not None:
        before = after = args.context

    if before < 0 or after < 0:
        parser.error("Context values must be non-negative")

    if args.matchlen is not None and args.matchlen < 1:
        parser.error("Match length must be a positive integer")

    # Compile both expressions up front so invalid patterns fail before searching.
    flags = re.IGNORECASE if args.ignore_case else 0

    try:
        pattern = re.compile(args.pattern, flags)
    except re.error as exc:
        parser.error(f"Invalid regular expression: {exc}")

    filename_pattern = None
    if args.filename is not None:
        try:
            filename_pattern = re.compile(args.filename, flags)
        except re.error as exc:
            parser.error(f"Invalid filename regular expression: {exc}")

    # Discover ordinary files and archives separately because their contents
    # require different extraction paths.
    files = collect_files(args.paths, args.ext, filename_pattern)
    archives = (
        collect_archives(args.paths, filename_pattern) if args.include_archives else []
    )

    if not files and not archives:
        print("No supported files found.", file=sys.stderr)
        return 1

    # Filter archive members with the same extension policy used for normal files.
    allowed_extensions = get_allowed_extensions(args.ext)
    archive_members = []
    for archive_path in archives:
        try:
            archive_type, member_names = archive_member_names(archive_path)
        except Exception as exc:
            print(
                f"\nWarning: cannot read archive {archive_path}: {exc}",
                file=sys.stderr,
            )
            continue

        compressed_stream = archive_type == "compressed"
        archive_members.extend(
            (archive_path, archive_type, member_name, compressed_stream)
            for member_name in member_names
            if is_searchable_archive_member(
                member_name,
                allowed_extensions,
                filename_pattern,
                compressed_stream,
            )
        )

    # Print number of files, and number of archives and their content-files.
    files_searched = 0
    files_total = len(files)
    if args.include_archives:
        print(
            f"Found {files_total} files, and {len(archives)} archives containing another {len(archive_members)} files..."
        )
        archive_note = ""
    else:
        archive_note = " (not searching archives)"
    print(f"Searching {files_total} files...{archive_note}")

    total_matches = 0

    # Search ordinary files first, updating progress after each completed file.
    for path in files:
        total_matches += search_file(
            path,
            pattern,
            before,
            after,
            red,
            reset,
            args.matchlen,
        )
        files_searched += 1
        show_progress(
            files_searched,
            files_total,
            path,
            bar_width=20,
            label="Files " if args.include_archives else "",
        )

    # Print newline on stderr for correct printing afterwards.
    print(file=sys.stderr)

    if args.include_archives:
        if files:
            print(file=sys.stderr)
        else:
            show_progress(0, 0, "No ordinary files", label="Files ")
            print(file=sys.stderr)
        print(f"Searching {len(archive_members)} files in archives...")

        # Extract each selected member to a temporary file before reusing search_file.
        if archive_members:
            with tempfile.TemporaryDirectory(prefix="egrepd_archive_") as temp_dir:
                temp_root = Path(temp_dir)
                for index, (archive_path, archive_type, member_name, _) in enumerate(
                    archive_members, start=1
                ):
                    member_path = temp_root / f"member{Path(member_name).suffix}"
                    display_name = f"{archive_path}!/{member_name}"
                    try:
                        with member_path.open("wb") as destination:
                            copy_archive_member(
                                archive_path,
                                archive_type,
                                member_name,
                                destination,
                            )
                        total_matches += search_file(
                            member_path,
                            pattern,
                            before,
                            after,
                            red,
                            reset,
                            args.matchlen,
                            display_name=member_name,
                            link_path=archive_path,
                        )
                    except Exception as exc:
                        print(
                            f"\nWarning: cannot read {display_name}: {exc}",
                            file=sys.stderr,
                        )

                    show_progress(
                        index,
                        len(archive_members),
                        archive_path.name,
                        bar_width=20,
                        label="Archives ",
                    )
        else:
            show_progress(0, 0, "No matching archive members", label="Archives ")

        # Print newline on stderr for correct printing afterwards.
        print(file=sys.stderr)

    # Print summary of searched files.
    if args.include_archives:
        print(
            f"Searched {files_total} files, and {len(archives)} archives containing another {len(archive_members)} files."
        )
    else:
        print(f"Searched {files_total} files (archives excluded by user).")

    return 0 if total_matches else 1


if __name__ == "__main__":
    sys.exit(main())
