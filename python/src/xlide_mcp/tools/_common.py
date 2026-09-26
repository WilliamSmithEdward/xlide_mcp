"""Shared pieces every tool group uses: annotations, truncation, and diffs.

The annotation helpers exist so the read/write split is declared once per tool
rather than spelled out at each registration. A client that gates writes behind a
confirmation reads `destructive_hint`, and a tool that silently lied about it
would be confirmed by nobody.
"""

from __future__ import annotations

import difflib
import json
from typing import Any

from mcp.types import ToolAnnotations

# A tool result is model context, and a 40,000-line module read in full leaves no
# room for the work. Reads that can be unbounded say how much was cut and how to
# ask for the rest, rather than truncating in silence.
MAX_RESULT_CHARS = 120_000
MAX_LIST_ITEMS = 2_000

ALLOW_PROTECTED_DESCRIPTION = (
    "Permit saving a password-protected VBA project. Set true only after the user agrees."
)
ALLOW_SIGNATURE_DESCRIPTION = (
    "Permit saving a change that removes the VBA project's digital signature, where "
    "applicable. Set true only after the user agrees."
)


def read_only(title: str) -> ToolAnnotations:
    return ToolAnnotations(
        title=title,
        read_only_hint=True,
        destructive_hint=False,
        idempotent_hint=True,
        open_world_hint=False,
    )


def writes(title: str, *, destructive: bool = False, idempotent: bool = True) -> ToolAnnotations:
    return ToolAnnotations(
        title=title,
        read_only_hint=False,
        destructive_hint=destructive,
        idempotent_hint=idempotent,
        open_world_hint=False,
    )


def truncate(text: str, limit: int = MAX_RESULT_CHARS, *, hint: str = "") -> tuple[str, bool]:
    """Cut a long result and say so. Returns the text and whether it was cut."""
    if len(text) <= limit:
        return text, False
    kept = text[:limit]
    tail = f"\n\n[cut after {limit:,} characters of {len(text):,}."
    tail += f" {hint}]" if hint else "]"
    return kept + tail, True


# A diff longer than this is not one an agent reads; it is one it summarizes.
MAX_DIFF_LINES = 400


def unified_diff(
    before: str,
    after: str,
    *,
    label: str,
    from_label: str = "before",
    to_label: str = "after",
    narrower: str = "",
) -> tuple[str, bool]:
    """A diff of two sources, and whether it was cut.

    Every diff this server produces comes from here: the one a write reports and
    the one a revision comparison reports are the same text for the same change,
    which is the only reason an agent can tell them apart by context rather than
    by shape. Line endings are normalized first, because a host rewriting CRLF is
    not somebody's edit.
    """
    lines = list(
        difflib.unified_diff(
            _lines(before),
            _lines(after),
            fromfile=f"{label} ({from_label})",
            tofile=f"{label} ({to_label})",
            lineterm="",
            n=3,
        )
    )
    if not lines:
        return "(no change)", False
    if len(lines) <= MAX_DIFF_LINES:
        return "\n".join(lines), False
    withheld = len(lines) - MAX_DIFF_LINES
    kept = lines[:MAX_DIFF_LINES]
    tail = f"... {withheld} more diff lines."
    kept.append(f"{tail} {narrower}".strip())
    return "\n".join(kept), True


def _lines(text: str) -> list[str]:
    return text.replace("\r\n", "\n").replace("\r", "\n").splitlines()


def change_summary(before: str, after: str) -> dict[str, int]:
    """How much a write moved, without shipping the whole diff."""
    before_lines = before.splitlines()
    after_lines = after.splitlines()
    added = removed = 0
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(
        None, before_lines, after_lines, autojunk=False
    ).get_opcodes():
        if tag in {"replace", "delete"}:
            removed += i2 - i1
        if tag in {"replace", "insert"}:
            added += j2 - j1
    return {
        "lines_before": len(before_lines),
        "lines_after": len(after_lines),
        "lines_added": added,
        "lines_removed": removed,
    }


def limited(items: list[Any], limit: int = MAX_LIST_ITEMS) -> tuple[list[Any], int]:
    """The first `limit` items and how many there were, for a bounded listing."""
    if len(items) <= limit:
        return items, len(items)
    return items[:limit], len(items)


# How many of a thing one result carries. A legacy project with 400 modules is
# ordinary, and a data-entry form with 400 controls is not absurd; measured, that
# form came back as 112 KB in a single call, which is most of what an agent has
# to think with. These are generous enough that no ordinary file meets them.
MAX_ITEMS: dict[str, int] = {
    "modules": 300,
    "controls": 300,
    "shapes": 300,
    "queries": 300,
    "forms": 300,
    "files": 2_000,
    "plan": 500,
    "relationships": 300,
    "tables": 500,
}


# The size one listing may occupy. A count alone is the wrong measure: 300
# modules is 34 KB and 300 form controls with their properties is 84 KB, because
# what an item costs depends entirely on what an item is.
MAX_LISTING_CHARS = 40_000


def bound(items: list[Any], what: str, narrower: str = "") -> tuple[list[Any], str]:
    """Cut a listing to its ceiling and say what was cut and how to see the rest.

    Returns the items to send and a note, empty when nothing was cut. Every
    listing in this server goes through here rather than choosing its own limit,
    so a result that stops short says so the same way wherever it came from: a
    silent truncation reads as a complete answer, and an agent acts on it.

    Two ceilings, whichever comes first: a count, and a size. The size is what
    actually matters, because a tool result is model context and an item's cost
    depends on what the item is.
    """
    limit = min(MAX_ITEMS.get(what, MAX_LIST_ITEMS), len(items))
    budget = MAX_LISTING_CHARS
    kept = 0
    for item in items[:limit]:
        budget -= len(json.dumps(item, default=str))
        if budget < 0:
            break
        kept += 1
    kept = max(kept, 1) if items else 0

    if kept >= len(items):
        return items, ""
    withheld = len(items) - kept
    note = f"{len(items)} {what} in all; the first {kept} are here and {withheld} are not."
    return items[:kept], f"{note} {narrower}".strip()


def page(
    items: list[Any], what: str, offset: int, max_results: int
) -> tuple[list[Any], int | None]:
    """A size-bounded page and the offset of the next one, if any."""
    shown, _ = bound(items[offset : offset + max_results], what)
    following = offset + len(shown)
    return shown, following if following < len(items) else None
