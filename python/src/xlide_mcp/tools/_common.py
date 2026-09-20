"""Shared pieces every tool group uses: annotations, truncation, and diffs.

The annotation helpers exist so the read/write split is declared once per tool
rather than spelled out at each registration. A client that gates writes behind a
confirmation reads `destructive_hint`, and a tool that silently lied about it
would be confirmed by nobody.
"""

from __future__ import annotations

import difflib
from typing import Any

from mcp.types import ToolAnnotations

# A tool result is model context, and a 40,000-line module read in full leaves no
# room for the work. Reads that can be unbounded say how much was cut and how to
# ask for the rest, rather than truncating in silence.
MAX_RESULT_CHARS = 120_000
MAX_LIST_ITEMS = 2_000


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


def unified_diff(before: str, after: str, *, label: str) -> str:
    """A diff of two module sources, for reporting a write that already happened."""
    diff = difflib.unified_diff(
        before.splitlines(),
        after.splitlines(),
        fromfile=f"{label} (before)",
        tofile=f"{label} (after)",
        lineterm="",
        n=2,
    )
    body = "\n".join(diff)
    return body or "(no change)"


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
