from pathlib import Path
import shutil
from types import SimpleNamespace

import pytest

from egrepd import text_extraction


DATA_DIR = Path(__file__).resolve().parents[1] / "data"


def sample(name: str) -> Path:
    return DATA_DIR / name


def assert_text_contains(lines: list[str], text: str) -> None:
    assert text.casefold() in "\n".join(lines).casefold()


@pytest.mark.parametrize(
    "extractor",
    [
        text_extraction.extract_pdf_text_pypdf,
        text_extraction.extract_pdf_text_pymupdf,
        text_extraction.extract_pdf_text,
    ],
)
def test_pdf_extractors_include_page_markers_and_text(extractor) -> None:
    """Verify PDF extractors preserve page markers and extract sample text.

    Args:
        extractor: PDF extraction function under test.
    """
    lines = extractor(sample("moby-dick.pdf"))

    # Check that page boundaries and extracted content are both retained.
    assert lines[0] == "[PDF page 1]"
    assert any(line.startswith("[PDF page ") for line in lines)
    assert_text_contains(lines, "Call me Ishmael")


def test_pdf_extractor_falls_back_to_pypdf(monkeypatch) -> None:
    """Verify PDF extraction falls back to pypdf when PyMuPDF fails.

    Args:
        monkeypatch: Pytest fixture for replacing extractor functions.
    """
    expected = ["[PDF page 1]", "fallback text"]

    # Force the preferred engine to fail so the fallback path is exercised.
    def fail_pymupdf(path: Path) -> list[str]:
        raise RuntimeError("PyMuPDF could not read the document")

    monkeypatch.setattr(
        text_extraction, "extract_pdf_text_pymupdf", fail_pymupdf
    )
    monkeypatch.setattr(
        text_extraction, "extract_pdf_text_pypdf", lambda path: expected
    )

    assert text_extraction.extract_pdf_text(sample("moby-dick.pdf")) == expected


@pytest.mark.parametrize(
    ("filename", "expected_text"),
    [
        ("moby-dick.docx", "Call me Ishmael"),
        ("rime-of-the-ancient-mariner.docx", "Ancient Mariner"),
    ],
)
def test_docx_extracts_paragraph_text(filename: str, expected_text: str) -> None:
    """Verify DOCX extraction returns paragraphs from the sample documents.

    Args:
        filename: Name of the DOCX sample file.
        expected_text: Text expected in the extracted paragraphs.
    """
    lines = text_extraction.extract_docx_text(sample(filename))

    # The samples contain narrative text in paragraph form.
    assert lines
    assert_text_contains(lines, expected_text)


def test_doc_extracts_text_when_antiword_is_installed() -> None:
    """Verify legacy DOC extraction when antiword is available."""
    if shutil.which("antiword") is None:
        # The extractor depends on this external system utility.
        pytest.skip("antiword is required to extract legacy .doc files")

    lines = text_extraction.extract_doc_text(sample("moby-dick.doc"))

    assert lines
    assert_text_contains(lines, "Call me Ishmael")


def test_pptx_extracts_slide_text() -> None:
    """Verify PPTX extraction includes slide markers and visible text."""
    lines = text_extraction.extract_pptx_text(sample("moby-dick.pptx"))

    # Slide markers make the extracted content traceable to its source slide.
    assert lines[0] == "[Slide 1]"
    assert_text_contains(lines, "Ishmael")


@pytest.mark.parametrize("filename", ["moby-dick.odt", "moby-dick.odp"])
def test_odf_extracts_text(filename: str) -> None:
    """Verify ODF extraction returns non-empty text from sample files.

    Args:
        filename: Name of the ODF sample file.
    """
    lines = text_extraction.extract_odf_text(sample(filename))

    # The two ODF samples differ in document type but should contain readable text.
    assert lines
    assert any(line.strip() for line in lines)


@pytest.mark.parametrize(
    ("extractor", "filename"),
    [
        (text_extraction.extract_xlsx_text, "rime-of-the-ancient-mariner.xlsx"),
        (text_extraction.extract_xls_text, "rime-of-the-ancient-mariner.xls"),
        (text_extraction.extract_ods_text, "rime-of-the-ancient-mariner.ods"),
    ],
)
def test_spreadsheet_extractors_include_sheet_names_and_text(
    extractor, filename: str
) -> None:
    """Verify spreadsheet extraction includes sheet names and cell content.

    Args:
        extractor: Spreadsheet extraction function under test.
        filename: Name of the spreadsheet sample file.
    """
    lines = extractor(sample(filename))

    # Each spreadsheet output row starts with a sheet name and uses the default delimiter.
    assert lines
    assert all("|" in line for line in lines)
    assert_text_contains(lines, "Mariner")


def test_spreadsheet_extractors_respect_custom_separator() -> None:
    """Verify spreadsheet extractors use a caller-provided separator."""
    for extractor, filename in [
        (text_extraction.extract_xlsx_text, "rime-of-the-ancient-mariner.xlsx"),
        (text_extraction.extract_xls_text, "rime-of-the-ancient-mariner.xls"),
        (text_extraction.extract_ods_text, "rime-of-the-ancient-mariner.ods"),
    ]:
        # Confirm the requested delimiter is applied to every emitted row.
        lines = extractor(sample(filename), separator="::")

        assert lines
        assert all("::" in line for line in lines)


def test_legacy_ppt_converts_to_pptx_and_extracts(monkeypatch, tmp_path: Path) -> None:
    """Verify legacy PPT conversion passes the converted file to extraction.

    Args:
        monkeypatch: Pytest fixture for replacing subprocess and extraction.
        tmp_path: Temporary directory fixture for the source presentation.
    """
    source = tmp_path / "presentation.ppt"
    source.touch()
    extracted_file: Path | None = None

    # Simulate LibreOffice creating the converted file without invoking it.
    def fake_run(command, **kwargs):
        nonlocal extracted_file
        output_dir = Path(command[command.index("--outdir") + 1])
        extracted_file = output_dir / "presentation.pptx"
        extracted_file.touch()
        return SimpleNamespace(returncode=0, stderr="", stdout="")

    def fake_extract_pptx(path: Path) -> list[str]:
        assert path == extracted_file
        return ["[Slide 1]", "converted slide"]

    monkeypatch.setattr(text_extraction.subprocess, "run", fake_run)
    monkeypatch.setattr(text_extraction, "extract_pptx_text", fake_extract_pptx)

    # Extraction should receive the temporary PPTX and return its text.
    assert text_extraction.extract_ppt_text(source) == [
        "[Slide 1]",
        "converted slide",
    ]


@pytest.mark.parametrize(
    ("filename", "expected_text"),
    [
        ("moby-dick.txt", "Call me Ishmael"),
        ("moby-dick.md", "Call me Ishmael"),
        ("moby-dick.docx", "Call me Ishmael"),
        ("moby-dick.pdf", "Call me Ishmael"),
        ("moby-dick.odt", "Ishmael"),
        ("moby-dick.odp", None),
        ("moby-dick.pptx", "Ishmael"),
        ("rime-of-the-ancient-mariner.xlsx", "Mariner"),
        ("rime-of-the-ancient-mariner.xls", "Mariner"),
        ("rime-of-the-ancient-mariner.ods", "Mariner"),
    ],
)
def test_extract_text_dispatches_supported_data_files(
    filename: str, expected_text: str | None
) -> None:
    """Verify extension-based dispatch for the supplied supported files.

    Args:
        filename: Name of a sample file to dispatch.
        expected_text: Optional text expected in the extracted result.
    """
    lines = text_extraction.extract_text(sample(filename))

    # A missing expected phrase is intentional for the presentation sample.
    assert lines
    if expected_text is not None:
        assert_text_contains(lines, expected_text)


def test_extract_text_replaces_invalid_utf8(tmp_path: Path) -> None:
    """Verify plain-text extraction replaces invalid UTF-8 bytes.

    Args:
        tmp_path: Temporary directory fixture for the test file.
    """
    path = tmp_path / "invalid.bin"
    path.write_bytes(b"first line\ninvalid: \xff")

    # Invalid bytes should be replaced instead of preventing text extraction.
    assert text_extraction.extract_text(path) == [
        "first line",
        "invalid: \N{REPLACEMENT CHARACTER}",
    ]


def test_extract_text_matches_extension_case_insensitively(tmp_path: Path) -> None:
    """Verify plain-text dispatch accepts uppercase file extensions.

    Args:
        tmp_path: Temporary directory fixture for the test file.
    """
    path = tmp_path / "sample.TXT"
    path.write_text("plain text", encoding="utf-8")

    # Uppercase suffixes should use the same plain-text path as lowercase ones.
    assert text_extraction.extract_text(path) == ["plain text"]
