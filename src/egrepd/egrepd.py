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
    ".xlsx",
    ".ods",
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
    """Suppress noisy warnings from document-processing libraries."""
    logging.getLogger("pypdf").setLevel(logging.ERROR)
    warnings.filterwarnings("ignore", module=r"pypdf(\.|$)")


def configure_colors(mode: str) -> tuple[str, str]:
    """Return ANSI color codes according to the selected mode."""

    use_color = mode == "always" or (mode == "auto" and sys.stdout.isatty())

    if use_color:
        return "\033[31m", "\033[0m"

    return "", ""


def hyperlink(path: Path) -> str:
    """Return an OSC 8 clickable file hyperlink."""
    uri = path.resolve().as_uri()
    return f"====> {uri} "


def extract_pdf_text(path: Path) -> list[str]:
    """Extract text from a PDF file."""
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


def extract_doc_text(path: Path) -> list[str]:
    """Extract text from a Word .doc document. Requires the antiword linux package."""
    try:
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
    """Extract text from a Word .docx document."""
    try:
        from docx import Document
    except ImportError as exc:
        raise RuntimeError(
            "DOCX support requires: " "python -m pip install python-docx"
        ) from exc

    document = Document(str(path))
    lines = []

    for paragraph in document.paragraphs:
        lines.extend(paragraph.text.splitlines() or [""])

    return lines


def extract_pptx_text(path: Path) -> list[str]:
    """Extract slide text and speaker notes from a PowerPoint."""
    try:
        from pptx import Presentation
    except ImportError as exc:
        raise RuntimeError(
            "PowerPoint support requires: " "python -m pip install python-pptx"
        ) from exc

    presentation = Presentation(str(path))
    lines = []

    def extract_shape(shape):
        """Extract text from shapes, tables, and grouped shapes."""
        extracted = []

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
    """Convert a legacy .ppt file to .pptx and extract its text."""
    try:
        with tempfile.TemporaryDirectory(prefix="docgrep_") as temp_dir:

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
    """Extract paragraph and heading text from an OpenDocument file."""
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
    """Extract text from an Excel .xlsx workbook, including worksheet names."""
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError(
            "Excel support requires: python -m pip install openpyxl"
        ) from exc

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
    """Extract text from a legacy Excel .xls workbook, including worksheet names."""
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError(
            "Legacy Excel support requires: python -m pip install xlrd"
        ) from exc

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
    """Extract text and cell values from an OpenDocument .ods spreadsheet."""
    try:
        from odf.opendocument import load
        from odf.table import Table, TableRow, TableCell, CoveredTableCell
        from odf.text import P, Span, S, Tab, LineBreak
    except ImportError as exc:
        raise RuntimeError("ODS support requires: python -m pip install odfpy") from exc

    document = load(str(path))
    lines = []

    def extract_paragraph_text(paragraph) -> str:
        """Extract paragraph text, including common ODF inline elements."""
        parts = []

        def visit(node):
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
        """Extract a cell's textual content or its stored value."""
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
    """Extract text lines from supported document types."""
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
    formats: list[str] | None = None,
    filename_pattern: re.Pattern[str] | None = None,
) -> list[Path]:
    """Collect files, filtering by the requested extensions and filename pattern."""

    if formats is None:
        allowed_extensions = DEFAULT_EXTENSIONS
    else:
        allowed_extensions = {
            f".{fmt.lower().lstrip('.')}" for fmt in formats if fmt.strip()
        }

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
    """Return the recognized archive suffix, preferring compound suffixes."""
    filename = filename.lower()
    return next(
        (suffix for suffix in ARCHIVE_SUFFIXES if filename.endswith(suffix)),
        None,
    )


def collect_archives(
    inputs: list[str],
    filename_pattern: re.Pattern[str] | None = None,
) -> list[Path]:
    """Collect supported archive files independently of ordinary files."""
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
    """List searchable regular-file candidates in an archive."""
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
    """Copy one archive member to a temporary file-like destination."""
    if archive_type == "zip":
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
    """Merge overlapping or adjacent context ranges."""
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
    """Display a matchlen-character excerpt around every match in a line."""
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
    """Search a file and display matches with context."""
    try:
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

    matches = [index for index, line in enumerate(lines) if pattern.search(line)]

    if not matches:
        return 0

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

    # First line: clickable file URI,
    # , but go back to beginning of line and clear line first:
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
    filename = str(filename)
    terminal_width = shutil.get_terminal_size(fallback=(80, 24)).columns

    fraction = current / total if total else 0
    filled = int(bar_width * fraction)
    bar = "#" * filled + "." * (bar_width - filled)

    prefix = f"\r{label}[{bar}] {fraction:6.1%} "
    available = max(0, terminal_width - len(prefix) - 1)

    # Truncate long filenames
    if len(filename) > available:
        filename = (
            "..." + filename[-(available - 3) :]
            if available > 3
            else filename[:available]
        )

    message = prefix + filename

    print(message.ljust(terminal_width - 1), end="", file=sys.stderr, flush=True)


def main() -> int:
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
        type=lambda value: [
            fmt.strip().lower().lstrip(".") for fmt in value.split(",") if fmt.strip()
        ],
        default=None,
        metavar="EXTENSIONS",
        help=(
            "Comma-separated file extensions to search, including custom text formats "
            + "(e.g., py,R,java). "
            + "Default: "
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
    args = parser.parse_args()

    red, reset = configure_colors(args.color)

    before = args.before_context
    after = args.after_context

    if args.context is not None:
        before = after = args.context

    if before < 0 or after < 0:
        parser.error("Context values must be non-negative")

    if args.matchlen is not None and args.matchlen < 1:
        parser.error("Match length must be a positive integer")

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

    files = collect_files(args.paths, args.ext, filename_pattern)
    archives = (
        collect_archives(args.paths, filename_pattern) if args.include_archives else []
    )

    if not files and not archives:
        print("No supported files found.", file=sys.stderr)
        return 1

    files_searched = 0
    files_total = len(files)
    print(f"Searching {files_total} files...")

    total_matches = 0

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

    allowed_extensions = (
        DEFAULT_EXTENSIONS
        if args.ext is None
        else {f".{extension.lower().lstrip('.')}" for extension in args.ext}
    )
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

    if args.include_archives:
        if files:
            print(file=sys.stderr)
        else:
            show_progress(0, 0, "No ordinary files", label="Files ")
            print(file=sys.stderr)
        print(f"Searching {len(archive_members)} files in archives...")

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

        print(file=sys.stderr)

    return 0 if total_matches else 1


if __name__ == "__main__":
    sys.exit(main())
