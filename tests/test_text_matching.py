import re
from pathlib import Path

import pytest

from egrepd import text_matching


def test_hyperlink_formats_path_as_terminal_uri(tmp_path: Path) -> None:
    path = tmp_path / "fake document.txt"

    assert text_matching.hyperlink(path) == f"====> {path.resolve().as_uri()} "


@pytest.mark.parametrize(
    ("ranges", "expected"),
    [
        ([(5, 7), (1, 2), (3, 4), (10, 12)], [(1, 7), (10, 12)]),
        ([], []),
        ([(4, 4)], [(4, 4)]),
    ],
)
def test_merge_ranges(
    ranges: list[tuple[int, int]], expected: list[tuple[int, int]]
) -> None:
    assert text_matching.merge_ranges(ranges) == expected


def test_truncate_match_line_shows_centered_excerpts() -> None:
    line = "0123456789abcdefghij"

    assert (
        text_matching.truncate_match_line(
            line, re.compile(r"34|ef"), 6, "", ""
        )
        == "...123456...cdefgh..."
    )


def test_truncate_match_line_keeps_long_matches_and_highlights() -> None:
    assert text_matching.truncate_match_line(
        "0123456789abcdefghij", re.compile("3456789"), 4, "<", ">"
    ) == "...<3456789>..."

    assert text_matching.truncate_match_line(
        "one hit two hit", re.compile("hit"), 20, "<", ">"
    ) == "one <hit> two <hit>"


def test_truncate_match_line_returns_empty_when_no_match() -> None:
    assert text_matching.truncate_match_line(
        "nothing here", re.compile("missing"), 5, "", ""
    ) == "..."


@pytest.mark.parametrize(
    ("line", "matchlen", "expected"),
    [
        ("short", 5, "short"),
        ("long context line", 4, "long..."),
    ],
)
def test_truncate_context_line(line: str, matchlen: int, expected: str) -> None:
    assert text_matching.truncate_context_line(line, matchlen) == expected


def test_search_file_prints_matches_and_merged_context(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    path = tmp_path / "fake.txt"
    monkeypatch.setattr(
        text_matching,
        "extract_text",
        lambda source: [
            "before",
            "needle one",
            "between",
            "needle two",
            "after",
        ],
    )

    count = text_matching.search_file(
        path, re.compile("needle"), before=1, after=1
    )

    captured = capsys.readouterr()
    assert count == 2
    assert captured.out == (
        f"\r====> {path.resolve().as_uri()} \n"
        "1- before\n"
        "2: needle one\n"
        "3- between\n"
        "4: needle two\n"
        "5- after\n"
        "---\n"
    )
    assert captured.err == "\r\033[K"


def test_search_file_formats_archive_member_link(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    path = tmp_path / "archive.zip"
    monkeypatch.setattr(
        text_matching, "extract_text", lambda source: ["fake match"]
    )

    count = text_matching.search_file(
        path,
        re.compile("match"),
        before=0,
        after=0,
        red="<",
        reset=">",
        display_name="folder/member.txt",
        link_path=path,
    )

    captured = capsys.readouterr()
    assert count == 1
    assert captured.out == (
        f"\r====> {path.resolve().as_uri()} !/folder/member.txt\n"
        "1: fake <match>\n"
        "---\n"
    )


def test_search_file_without_matches_prints_nothing(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    monkeypatch.setattr(
        text_matching, "extract_text", lambda source: ["fake text"]
    )

    assert text_matching.search_file(
        tmp_path / "fake.txt", re.compile("missing"), before=0, after=0
    ) == 0
    assert capsys.readouterr().out == ""


def test_search_file_reports_extraction_error(
    monkeypatch: pytest.MonkeyPatch,
    capsys: pytest.CaptureFixture[str],
    tmp_path: Path,
) -> None:
    def fail_extraction(source: Path) -> list[str]:
        raise OSError("fake read failure")

    monkeypatch.setattr(text_matching, "extract_text", fail_extraction)
    path = tmp_path / "fake.txt"

    assert text_matching.search_file(
        path, re.compile("needle"), before=0, after=0
    ) == 0
    captured = capsys.readouterr()
    assert captured.out == ""
    assert f"Warning: cannot read {path}: fake read failure" in captured.err
