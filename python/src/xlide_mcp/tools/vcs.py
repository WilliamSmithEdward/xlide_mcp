"""What changed inside an Office file since a git revision.

`git diff` cannot answer this on its own. An Office file is one binary blob to
git, so a commit that changed three lines of VBA and a commit that replaced the
whole project look identical in a diff and in a review. The code is in there, and
it is readable: this extracts the blob at a revision, renders both versions
through `textual`, and diffs them section by section.

That makes it the tool to call before committing, and the one that makes a code
review of a workbook possible at all. For the same thing in git's own diff, so
that `git diff` and a side-by-side view show it too, `xlide-mcp --textconv` is a
textconv driver built on the same renderer.

It shells out to git rather than binding a library. The repository the user works
in is the one git already knows about, with their config, their worktree layout
and their credential setup; a second implementation of that would be a second set
of answers.
"""

from __future__ import annotations

import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import resolve_path
from ..textual import COVERAGE, Section, sections
from ._common import MAX_ITEMS, bound, change_summary, read_only, unified_diff

# A git call that has not answered by now is a repository problem, not slow work.
GIT_TIMEOUT = 30.0


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_git_changes",
        title="What changed inside the file",
        annotations=read_only("What changed inside the file"),
        description=(
            "Lists what changed inside an Office file since a git revision, one entry per VBA "
            "module and Power Query, with a unified diff. git diff cannot show this: the file "
            "is binary, so a commit that changed one line and one that replaced the whole "
            "project look the same. Call this before committing, and to review what an agent "
            "or a colleague changed. Worksheet cell values are not compared. Defaults to HEAD; "
            "pass any revision git understands. Needs the file to be inside a git repository "
            "and tracked in that revision."
        ),
    )
    def git_changes(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        revision: Annotated[
            str,
            Field(
                default="HEAD",
                description="Any revision git understands: HEAD, a branch, a tag, a SHA.",
            ),
        ] = "HEAD",
        section: Annotated[
            str,
            Field(
                default="",
                description=(
                    "Only this one, named as a module or a query is named. "
                    "Empty reports every one."
                ),
            ),
        ] = "",
        include_diff: Annotated[
            bool,
            Field(
                default=True,
                description=(
                    "Include the unified diff. Off gives just what changed and by how much."
                ),
            ),
        ] = True,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        if info.host == "vb6":
            raise ToolError(
                "A Visual Basic 6 project's modules are text files that git diffs directly, "
                "so `git diff` already shows what changed in them. This tool exists for the "
                "hosts whose code is inside a binary container."
            )
        repository = _repository_for(path)
        relative = _relative_to_repository(path, repository)
        blob = _blob_at(repository, revision, relative, path.name)

        current = _by_key(sections(path, info))
        with _temporary_copy(blob, path.suffix) as older_path:
            previous = _by_key(sections(older_path, info))

        wanted = section.strip().casefold()
        entries: list[dict[str, Any]] = []
        for key in sorted(set(current) | set(previous)):
            before = previous.get(key)
            after = current.get(key)
            present = after or before
            assert present is not None
            if wanted and present.name.casefold() != wanted:
                continue
            entries.append(_entry(present, before, after, include_diff))

        if wanted and not entries:
            raise ToolError(
                f"Nothing named {section!r} in either version. Now: "
                + (_inventory(current.values()) or "(nothing)")
            )

        changed = [e for e in entries if e["status"] != "unchanged"]
        # What a reviewer came for is what changed, so the unchanged sections are
        # the ones dropped when a project is too big to report whole.
        crowded = len(entries) > MAX_ITEMS["modules"]
        shown, note = bound(
            changed if crowded else entries, "modules", "Ask for one with section."
        )
        result: dict[str, Any] = {
            "path": str(path),
            "repository": str(repository),
            "revision": revision,
            "changed": len(changed),
            "changes": shown,
            "verdict": "no changes" if not changed else f"{len(changed)} changed",
            "covers": COVERAGE,
        }
        if crowded:
            result["note"] = (
                f"{len(entries)} sections in this file; only the ones that changed are "
                "listed. " + note
            ).strip()
        elif note:
            result["note"] = note
        return result


def _by_key(found: list[Section]) -> dict[tuple[str, str], Section]:
    """Sections by kind and folded name, so the two revisions line up.

    Folded, because a rename that only changes case is a rename in neither VBA nor
    Power Query, and reporting it as one section added and another removed would
    send a reviewer looking for a change nobody made.
    """
    return {(section.kind, section.name.casefold()): section for section in found}


@dataclass(frozen=True)
class _Change:
    status: str
    text: str
    truncated: bool
    added: int
    removed: int


def _entry(
    present: Section, before: Section | None, after: Section | None, include_diff: bool
) -> dict[str, Any]:
    change = _change(present, before, after)
    entry: dict[str, Any] = {
        "name": present.name,
        "kind": present.kind,
        "status": change.status,
        "lines_added": change.added,
        "lines_removed": change.removed,
    }
    if include_diff and change.text:
        entry["diff"] = change.text
        if change.truncated:
            entry["diff_truncated"] = True
    return entry


def _change(present: Section, before: Section | None, after: Section | None) -> _Change:
    old = before.source if before is not None else ""
    new = after.source if after is not None else ""
    if before is None:
        status = "added"
    elif after is None:
        status = "removed"
    elif _same(old, new):
        return _Change("unchanged", "", False, 0, 0)
    else:
        status = "modified"

    label = f"{present.kind} {present.name}"
    text, truncated = unified_diff(
        old,
        new,
        label=label,
        from_label="at the revision",
        to_label="now",
        narrower="Ask for this one on its own.",
    )
    summary = change_summary(old, new)
    return _Change(status, text, truncated, summary["lines_added"], summary["lines_removed"])


def _same(left: str, right: str) -> bool:
    """Line-ending style is not somebody's edit."""
    return left.replace("\r\n", "\n").rstrip() == right.replace("\r\n", "\n").rstrip()


def _inventory(found: Any) -> str:
    return ", ".join(sorted(f"{s.name} ({s.kind})" for s in found))


# ------------------------------------------------------------------- the repo


def _git(repository: Path, *arguments: str, binary: bool = False) -> Any:
    try:
        return subprocess.run(
            ["git", "-C", str(repository), *arguments],
            capture_output=True,
            timeout=GIT_TIMEOUT,
            check=False,
            text=not binary,
        )
    except FileNotFoundError as exc:
        raise ToolError(
            "git is not on the PATH, so changes inside the file cannot be read. "
            "Install git, or read the modules directly with xlide_read_module."
        ) from exc
    except subprocess.TimeoutExpired as exc:
        raise ToolError(f"git did not answer within {GIT_TIMEOUT:g} seconds.") from exc


def _repository_for(path: Path) -> Path:
    result = _git(path.parent, "rev-parse", "--show-toplevel")
    if result.returncode != 0:
        raise ToolError(
            f"{path.name} is not inside a git repository, so there is no revision to compare "
            "against. xlide_read_module reads the current source."
        )
    return Path(result.stdout.strip())


def _relative_to_repository(path: Path, repository: Path) -> str:
    try:
        # git wants forward slashes whatever the platform.
        return path.relative_to(repository).as_posix()
    except ValueError as exc:
        raise ToolError(f"{path} is not inside {repository}.") from exc


def _blob_at(repository: Path, revision: str, relative: str, name: str) -> bytes:
    result = _git(repository, "cat-file", "blob", f"{revision}:{relative}", binary=True)
    if result.returncode != 0:
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        if "exists on disk, but not in" in detail or "does not exist" in detail:
            raise ToolError(
                f"{name} is not tracked at {revision}. If it is new, everything in it is new; "
                "there is nothing to compare against."
            )
        if "unknown revision" in detail or "bad revision" in detail:
            raise ToolError(f"git does not know the revision {revision!r}. {detail}")
        raise ToolError(f"git could not read {name} at {revision}: {detail}")
    return result.stdout


class _temporary_copy:
    """The revision's bytes as a file on disk, with the extension that opens it.

    pyOpenVBA opens a path, and the container reader is chosen by extension, so a
    blob has to land somewhere named before its VBA can be read.
    """

    def __init__(self, data: bytes, suffix: str) -> None:
        self.data = data
        self.suffix = suffix
        self._path: Path | None = None

    def __enter__(self) -> Path:
        handle = tempfile.NamedTemporaryFile(suffix=self.suffix, delete=False)
        try:
            handle.write(self.data)
        finally:
            handle.close()
        self._path = Path(handle.name)
        return self._path

    def __exit__(self, *exc_info: object) -> None:
        if self._path is not None:
            self._path.unlink(missing_ok=True)
            self._path = None
