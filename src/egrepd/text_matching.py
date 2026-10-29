import re
import sys
from pathlib import Path

from .text_extraction import extract_text


def hyperlink(path: Path) -> str:
    """Format a file path as a terminal OSC 8 hyperlink."""
    uri = path.resolve().as_uri()
    return f"====> {uri} "


def merge_ranges(
    ranges: list[tuple[int, int]],
) -> list[tuple[int, int]]:
    """Combine overlapping or adjacent inclusive integer ranges."""
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
