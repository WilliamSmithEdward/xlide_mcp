"""An Office file rendered as text, so that a diff of one reads like source.

Everything an Office file holds that is code lives in a binary container, which
makes the file opaque to every tool that reads diffs: git, a review, an editor's
side-by-side view. One commit that changed a line of VBA and one that replaced
the whole project are the same three words, "Binary files differ".

This module is the one place that decides what "the text of this file" means. Two
callers depend on that being one answer:

* `xlide-mcp --textconv`, a git textconv driver, so `git diff` and `git show`
  render a workbook as its VBA and M.
* `xlide_git_changes`, which diffs two revisions section by section.

They would otherwise each grow their own idea of what is in a file, and the first
change to one would silently stop agreeing with the other.

What it does not cover is stated in the rendered text rather than left for the
reader to discover: cell values are not here. A diff that quietly showed nothing
for a workbook whose numbers all changed would be worse than no diff at all.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path

from .hosts import HostInfo, container, host_info

# Said in the rendered text and in every tool result built from it, because a
# reader who does not know the scope reads silence as "nothing changed".
COVERAGE = (
    "VBA modules and Power Query only. Worksheet cell values are not compared."
)

# A heading is a VBA comment, so that a dump opened in an editor still colours as
# code, and so that a rename shows as one changed line rather than a moved block.
RULE = "====="


@dataclass(frozen=True)
class Section:
    """One named piece of text inside a file: a module, or a query."""

    name: str
    kind: str
    """module | query"""

    source: str

    @property
    def heading(self) -> str:
        return f"' {RULE} {self.kind} {self.name} {RULE}"


def sections(path: Path, info: HostInfo | None = None) -> list[Section]:
    """Every piece of code in the file, in an order that does not move.

    Ordering is by kind then name, folded for case, so that two revisions line up
    section for section. A listing that came back in container order would diff as
    a rewrite the first time a host reordered its streams.
    """
    resolved = info or host_info(path)
    found: list[Section] = []
    found.extend(_modules(path, resolved))
    found.extend(_queries(path, resolved))
    return sorted(found, key=lambda s: (s.kind, s.name.casefold()))


def render(path: Path, info: HostInfo | None = None) -> str:
    """The file as one text document, for a diff to work on.

    Never raises for a file it cannot read. This runs as a git textconv driver,
    where a non-zero exit aborts the whole diff, and where returning nothing for
    an unreadable file would make two different files compare equal.
    """
    try:
        resolved = info or host_info(path)
    except Exception:
        return _opaque(path, "not an Office file this server reads")
    try:
        found = sections(path, resolved)
    except Exception as exc:
        return _opaque(path, f"could not be read: {exc}")

    lines = [f"' xlide-mcp: {COVERAGE}", ""]
    for section in found:
        lines.append(section.heading)
        lines.append(_normalize(section.source))
        lines.append("")
    inventory = _sheet_inventory(path, resolved)
    if inventory:
        lines.append(f"' {RULE} sheets {RULE}")
        lines.extend(inventory)
        lines.append("")
    if not found and not inventory:
        lines.append("' (no VBA and no Power Query in this file)")
    return "\n".join(lines).rstrip() + "\n"


# ------------------------------------------------------------------ the parts


def _modules(path: Path, info: HostInfo) -> list[Section]:
    if not info.readable:
        return []
    from . import project as project_layer

    with container(path, info) as handle:
        return [
            Section(name=module.name, kind="module", source=module.body)
            for module in project_layer.read_modules(handle, info)
        ]


def _queries(path: Path, info: HostInfo) -> list[Section]:
    if not info.supports_power_query:
        return []
    import pyopenvba

    try:
        with pyopenvba.PowerQueryWorkbook(path) as book:
            return [
                Section(name=query.name, kind="query", source=query.formula or "")
                for query in book.queries()
            ]
    except pyopenvba.PyOpenVBAError:
        # A workbook with no query part is the ordinary case, not a failure.
        return []


def _sheet_inventory(path: Path, info: HostInfo) -> list[str]:
    """Sheet names and used ranges, as comments.

    Not the cell values: this catches a sheet added, removed, renamed or grown,
    which is the structural change a code review of a workbook cares about. The
    binary formats keep their grid somewhere the package reader cannot see, so
    they contribute nothing here rather than a wrong answer.
    """
    if not info.supports_sheets:
        return []
    from .xlsx import Workbook, XlsxError

    try:
        sheets = Workbook(path).sheets()
    except (XlsxError, OSError):
        return []
    width = max((len(sheet.name) for sheet in sheets), default=0)
    return [
        f"' {sheet.name.ljust(width)}  {sheet.used_range or '(empty)'}"
        + ("  (hidden)" if sheet.hidden else "")
        for sheet in sheets
    ]


def _opaque(path: Path, why: str) -> str:
    """A file this cannot read, described so that two different ones differ.

    The digest is the point. Without it every unreadable file renders identically
    and git reports no change between two that share nothing.
    """
    try:
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
    except OSError as exc:
        digest = f"unreadable: {exc}"
    return f"' xlide-mcp: {why}.\n' sha256 {digest}\n"


def _normalize(source: str) -> str:
    """One line-ending style, so a host rewriting CRLF is not somebody's edit."""
    return source.replace("\r\n", "\n").replace("\r", "\n").rstrip()
