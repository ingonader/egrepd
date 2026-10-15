#!/usr/bin/env python3

import argparse
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
}

DEFAULT_EXTENSIONS = {".txt", ".md", ".pdf", ".docx", ".odt", ".odp", ".pptx"}

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


def search_file(
    path: Path,
    pattern: re.Pattern,
    before: int,
    after: int,
    red: str = "",
    reset: str = "",
    matchlen: int | None = None,
) -> int:
    """Search a file and display matches with context."""
    try:
        lines = extract_text(path)
    except Exception as exc:
        print(
            f"Warning: cannot read {path}: {exc}",
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

    # First line: clickable file URI.
    print(hyperlink(path))

    for start, end in ranges:
        for index in range(start, end + 1):
            line_number = index + 1
            line = lines[index]

            if index in match_set:
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
                print(f"{line_number}- {line}")

    # Separator after all matches and context.
    print("---")

    return len(matches)


def main() -> int:
    configure_logging()

    parser = argparse.ArgumentParser(
        description=(
            "Search txt, markdown, pdf, doc(x), odt, odp and ppt(x) files "
            "using regular expressions."
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
        "--ext",
        type=lambda value: [
            fmt.strip().lower().lstrip(".") for fmt in value.split(",") if fmt.strip()
        ],
        default=None,
        metavar="EXTENSIONS",
        help=(
            "Comma-separated file extensions to search, including custom text formats "
            "(e.g., py,R,java). "
            "Default: txt, md, pdf, docx, odt, odp, pptx. "
            "Legacy .doc and .ppt files must be explicitly requested."
        ),
    )

    parser.add_argument(
        "-f",
        "--filename",
        metavar="REGEX",
        help="Only search files whose name without its extension matches this regex",
    )

    parser.add_argument(
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

    if not files:
        print("No supported files found.", file=sys.stderr)
        return 1

    print(f"Searching {len(files)} files...")

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

    return 0 if total_matches else 1


if __name__ == "__main__":
    sys.exit(main())
