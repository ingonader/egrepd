#!/usr/bin/env python3

import argparse
import re
import sys
from pathlib import Path
import tempfile
import bz2
import gzip
import lzma
import xml.etree.ElementTree as ET
import zipfile
import shutil
import tarfile

import logging
import warnings
from typing import BinaryIO

from .text_extraction import extract_text

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


def get_allowed_extensions(
    extension_spec: str | None,
) -> list[re.Pattern[str]]:
    """Compile extension patterns, including additive and extensionless forms.

    Args:
        extension_spec: Comma-separated regular expressions matched against
            file suffixes. A leading `+` includes the default extensions, and
            empty entries select files without an extension. `None` selects
            only the defaults.

    Returns:
        Compiled regular expressions for matching file suffixes, including the
        empty suffix when extensionless files are selected.
    """
    include_defaults = extension_spec is None or extension_spec.startswith("+")
    allowed_extensions = (
        [
            re.compile(re.escape(extension), re.IGNORECASE)
            for extension in DEFAULT_EXTENSIONS
        ]
        if include_defaults
        else []
    )

    if extension_spec is not None:
        if include_defaults:
            extension_spec = extension_spec[1:]
        for extension in extension_spec.split(","):
            extension = extension.strip()
            allowed_extensions.append(
                re.compile(extension if extension else r"^$", re.IGNORECASE)
            )

    return allowed_extensions


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


def show_candidate_id_progress(cnt: int, candidate_name: str) -> None:
    msg = f"Searching candidates: Found {str(cnt)} {candidate_name}..."
    print("\r\033[K", end="", file=sys.stderr)
    print(f"{msg}", end="", file=sys.stderr, flush=True)


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
        formats: Comma-separated extension regular expressions. A leading `+`
            adds to the default extensions, and empty entries select files
            without an extension. `None` selects the default extensions.
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
            if is_allowed_extension(path.suffix, allowed_extensions) and (
                filename_pattern is None or filename_pattern.search(path.stem)
            ):
                files.add(path.resolve())
                file_cnt = len(files)
                show_candidate_id_progress(file_cnt, "files")
            continue

        if path.is_dir():
            for candidate in path.rglob("*"):
                if (
                    candidate.is_file()
                    and is_allowed_extension(candidate.suffix, allowed_extensions)
                    and (
                        filename_pattern is None
                        or filename_pattern.search(candidate.stem)
                    )
                ):
                    files.add(candidate.resolve())
                    file_cnt = len(files)
                    show_candidate_id_progress(file_cnt, "files")

    # Clear the progress display:
    print("\r\033[K", end="", file=sys.stderr)
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


def show_search_progress(
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


def _run() -> int:
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
            "Comma-separated extension regexes (e.g., 'py,R,java' or 'docx?|xlsx?'); "
            + "a leading '+' adds to the defaults (e.g., '+py,sh'), and empty entries "
            + "(e.g., 'py,') select extensionless files. Patterns match the suffix "
            + "with or without its leading dot. "
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

    # for interactive debugging:
    # args = parser.parse_args(["-a", "ishmael", str(Path(os.getcwd()).parent.parent)])

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

    try:
        allowed_extensions = get_allowed_extensions(args.ext)
    except re.error as exc:
        parser.error(f"Invalid extension regular expression: {exc}")

    # Discover ordinary files and archives separately because their contents
    # require different extraction paths.
    files = collect_files(args.paths, args.ext, filename_pattern)
    files_unique_extensions = {path.suffix for path in files}
    archives = (
        collect_archives(args.paths, filename_pattern) if args.include_archives else []
    )

    if not files and not archives:
        print("No supported files found.", file=sys.stderr)
        return 1

    # Filter archive members with the same extension policy used for normal files.
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

    # get filenames from archives, and their extentions:
    archive_members_filenames = [Path(i[2]) for i in archive_members]
    archive_members_unique_extensions = {i.suffix for i in archive_members_filenames}
    unique_extensions = files_unique_extensions | archive_members_unique_extensions
    unique_extensions_str = ",".join([i[1:] for i in unique_extensions])

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
    print(f"Matched extentions in candidate files: {unique_extensions_str}")
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
        show_search_progress(
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
            show_search_progress(0, 0, "No ordinary files", label="Files ")
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

                    show_search_progress(
                        index,
                        len(archive_members),
                        archive_path.name,
                        bar_width=20,
                        label="Archives ",
                    )
        else:
            show_search_progress(0, 0, "No matching archive members", label="Archives ")

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


def main() -> int:
    """Run the command-line application, reporting user interruptions cleanly."""
    try:
        return _run()
    except KeyboardInterrupt:
        print("Interrupted by user", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
