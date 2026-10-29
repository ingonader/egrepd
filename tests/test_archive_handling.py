import bz2
import gzip
import lzma
import os
import re
from io import BytesIO
from pathlib import Path

import pytest

from egrepd import archive_handling


DATA_DIR = Path(__file__).resolve().parents[1] / "data"


def sample(name: str) -> Path:
    return DATA_DIR / name


@pytest.mark.parametrize(
    ("filename", "expected"),
    [
        ("Archive.7z", ".7z"),
        ("Archive.tar", ".tar"),
        ("Archive.tar.bz2", ".tar.bz2"),
        ("Archive.tar.gz", ".tar.gz"),
        ("Archive.tar.xz", ".tar.xz"),
        ("Archive.zip", ".zip"),
        ("Archive.tar.lz", None),
        ("Archive.tar.lzma", None),
        ("Archive.tar.Z", None),
        ("Archive.tar.zst", None),
        ("moby-dick.md", None),
    ],
)
def test_archive_suffix(filename: str, expected: str | None) -> None:
    assert archive_handling.archive_suffix(filename) == expected


def test_show_candidate_id_progress(capsys: pytest.CaptureFixture[str]) -> None:
    archive_handling.show_candidate_id_progress(3, "archives")

    assert capsys.readouterr().err == "\rSearching candidates: Found 3 archives..."


@pytest.mark.parametrize(
    ("sort_spec", "expected"),
    [
        ("a", ["Archive.tar", "Archive.zip", "moby-dick.md"]),
        ("ar", ["moby-dick.md", "Archive.zip", "Archive.tar"]),
    ],
)
def test_sort_paths_by_name(
    tmp_path: Path, sort_spec: str, expected: list[str]
) -> None:
    paths = {tmp_path / name for name in expected}
    for path in paths:
        path.touch()

    assert [
        path.name for path in archive_handling.sort_paths(paths, sort_spec)
    ] == expected


def test_sort_paths_by_modification_time(tmp_path: Path) -> None:
    older = tmp_path / "older"
    newer = tmp_path / "newer"
    older.touch()
    newer.touch()
    os.utime(older, ns=(0, 1_000_000_000))
    os.utime(newer, ns=(0, 2_000_000_000))

    assert archive_handling.sort_paths({older, newer}, "t") == [newer, older]
    assert archive_handling.sort_paths({older, newer}, "tr") == [older, newer]


@pytest.mark.parametrize(
    ("suffix", "expected"),
    [
        (".md", True),
        ("md", True),
        (".MD", False),
        (".txt", True),
        (".pdf", False),
    ],
)
def test_is_allowed_extension(suffix: str, expected: bool) -> None:
    assert archive_handling.is_allowed_extension(
        suffix, [re.compile(r"md"), re.compile(r"\.txt")]
    ) is expected


def test_collect_archives_finds_supported_archives_and_filters_stems(
    capsys: pytest.CaptureFixture[str],
) -> None:
    archives = archive_handling.collect_archives(
        [str(DATA_DIR)], filename_pattern=re.compile("^Archive$")
    )

    assert {path.name for path in archives} == {
        "Archive.7z",
        "Archive.tar",
        "Archive.tar.bz2",
        "Archive.tar.gz",
        "Archive.tar.xz",
        "Archive.zip",
    }
    assert all(path.is_absolute() for path in archives)
    assert capsys.readouterr().err.endswith("\r\033[K")


def test_collect_archives_deduplicates_and_sorts_inputs(
    capsys: pytest.CaptureFixture[str],
) -> None:
    archive = sample("Archive.zip")

    assert archive_handling.collect_archives(
        [str(archive), str(DATA_DIR)], sort_spec="ar"
    )[0] == sample("Archive.zip").resolve()
    assert archive_handling.collect_archives(
        [str(archive), str(archive)]
    ) == [archive.resolve()]
    capsys.readouterr()


@pytest.mark.parametrize(
    ("filename", "container"),
    [
        ("Archive.zip", "zip"),
        ("Archive.tar", "tar"),
        ("Archive.tar.gz", "tar"),
        ("Archive.tar.bz2", "tar"),
        ("Archive.tar.xz", "tar"),
        ("Archive.7z", "7z"),
    ],
)
def test_archive_member_names_lists_sample_documents(
    filename: str, container: str
) -> None:
    archive_type, members = archive_handling.archive_member_names(sample(filename))

    assert archive_type == container
    assert "moby-dick.md" in members
    assert "rime-of-the-ancient-mariner.md" in members


def test_archive_member_names_returns_empty_for_unsupported_file() -> None:
    assert archive_handling.archive_member_names(sample("moby-dick.md")) == ("", [])


@pytest.mark.parametrize(
    ("filename", "container"),
    [
        ("Archive.zip", "zip"),
        ("Archive.tar.gz", "tar"),
        ("Archive.7z", "7z"),
    ],
)
def test_copy_archive_member_streams_sample_content(
    filename: str, container: str
) -> None:
    destination = BytesIO()

    archive_handling.copy_archive_member(
        sample(filename), container, "moby-dick.md", destination
    )

    assert destination.getvalue() == sample("moby-dick.md").read_bytes()


@pytest.mark.parametrize(
    ("suffix", "compress"),
    [
        (".gz", gzip.compress),
        (".bz2", bz2.compress),
        (".xz", lzma.compress),
    ],
)
def test_compressed_stream_lists_and_copies_sample_content(
    tmp_path: Path, suffix: str, compress
) -> None:
    source = sample("moby-dick.md").read_bytes()
    compressed_path = tmp_path / f"moby-dick.md{suffix}"
    compressed_path.write_bytes(compress(source))

    assert archive_handling.archive_member_names(compressed_path) == (
        "compressed",
        ["moby-dick.md"],
    )

    destination = BytesIO()
    archive_handling.copy_archive_member(
        compressed_path, "compressed", "moby-dick.md", destination
    )

    assert destination.getvalue() == source


@pytest.mark.parametrize(
    ("member_name", "allowed", "filename_pattern", "compressed", "expected"),
    [
        ("folder/moby-dick.md", [re.compile(r"\.md")], None, False, True),
        ("folder/moby-dick.MD", [re.compile(r"\.md")], None, False, True),
        ("folder/moby-dick.md", [re.compile(r"\.txt")], None, False, False),
        (
            "folder/moby-dick.md",
            [re.compile(r"\.md")],
            re.compile("moby"),
            False,
            True,
        ),
        (
            "folder/moby-dick.md",
            [re.compile(r"\.md")],
            re.compile("mariner"),
            False,
            False,
        ),
        ("moby-dick", [re.compile(r"\.txt")], None, True, True),
        ("moby-dick", [re.compile(r"\.md")], None, True, False),
    ],
)
def test_is_searchable_archive_member(
    member_name: str,
    allowed: list[re.Pattern[str]],
    filename_pattern: re.Pattern[str] | None,
    compressed: bool,
    expected: bool,
) -> None:
    assert bool(
        archive_handling.is_searchable_archive_member(
            member_name, allowed, filename_pattern, compressed
        )
    ) is expected
