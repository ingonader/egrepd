import bz2
import gzip
import lzma
import re
import shutil
import sys
import tarfile
import tempfile
import zipfile
from pathlib import Path
from typing import BinaryIO

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


def show_candidate_id_progress(cnt: int, candidate_name: str) -> None:
    msg = f"Searching candidates: Found {str(cnt)} {candidate_name}..."
    print(f"\r{msg}", end="", file=sys.stderr, flush=True)


def sort_paths(paths: set[Path], sort_spec: str) -> list[Path]:
    """Sort resolved candidate paths by name or file timestamp."""
    sort_key = next((option for option in sort_spec if option in "actu"), "a")
    reverse = "r" in sort_spec
    sorted_paths = sorted(paths)

    timestamp_attributes = {
        "c": "st_ctime",
        "t": "st_mtime",
        "u": "st_atime",
    }
    if sort_key in timestamp_attributes:
        timestamp_attribute = timestamp_attributes[sort_key]
        sorted_paths.sort(
            key=lambda path: getattr(path.stat(), timestamp_attribute),
            reverse=True,
        )

    if reverse:
        sorted_paths.reverse()

    return sorted_paths


def is_allowed_extension(
    suffix: str,
    allowed_extensions: list[re.Pattern[str]],
) -> bool:
    """Return whether a file suffix matches any selected extension pattern."""
    extension = suffix[1:] if suffix.startswith(".") else suffix
    return any(
        pattern.fullmatch(suffix) or pattern.fullmatch(extension)
        for pattern in allowed_extensions
    )


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
    sort_spec: str = "a",
) -> list[Path]:
    """Find supported archives at the given paths, independently of documents.

    Directories are searched recursively, and the optional filename expression
    is matched against the archive name with its complete archive suffix removed.

    Args:
        inputs: File and directory paths to inspect.
        filename_pattern: Optional regular expression matched against archive
            filename stems.
        sort_spec: Sort key (`a`, `c`, `t`, or `u`) and optional reverse flag
            (`r`).

    Returns:
        Resolved matching archive paths in the requested order.
    """
    # Deduplicate overlapping input paths before returning stable output.
    archives = set()
    archive_cnt = 0
    show_candidate_id_progress(archive_cnt, "archives")

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
                archive_cnt = len(archives)
                show_candidate_id_progress(archive_cnt, "archives")

    # Clear the progress display:
    print("\r\033[K", end="", file=sys.stderr)
    return sort_paths(archives, sort_spec)


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
    allowed_extensions: list[re.Pattern[str]],
    filename_pattern: re.Pattern[str] | None,
    compressed_stream: bool = False,
) -> bool:
    """Check whether an archive member passes extension and filename filters.

    A compressed stream has no internal filename, so it is treated as a text
    file for extension filtering.

    Args:
        member_name: Member path as stored in the archive.
        allowed_extensions: Regular expressions permitted to match the suffix.
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

    return is_allowed_extension(suffix, allowed_extensions) and (
        filename_pattern is None or filename_pattern.search(member_path.stem)
    )
