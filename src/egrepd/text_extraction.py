from pathlib import Path
import subprocess


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
