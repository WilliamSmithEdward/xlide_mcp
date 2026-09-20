"""What changed inside an Office file since a git revision.

`git diff` cannot answer this. An Office file is one binary blob to git, so a
commit that changed three lines of VBA and a commit that replaced the whole
project look identical in a diff and in a review. The code is in there, and it is
readable: this extracts the blob at a revision, opens both versions' VBA projects,
and diffs them module by module.

That makes it the tool to call before committing, and the one that makes a code
review of a workbook possible at all.

It shells out to git rather than binding a library. The repository the user works
in is the one git already knows about, with their config, their worktree layout
and their credential setup; a second implementation of that would be a second set
of answers.
"""

from __future__ import annotations

import difflib
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import project as project_layer
from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import resolve_path
from ._common import MAX_ITEMS, bound, read_only, truncate

# A git call that has not answered by now is a repository problem, not slow work.
GIT_TIMEOUT = 30.0

# A diff longer than this is not one an agent reads; it is one it summarizes.
MAX_DIFF_LINES = 400


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_git_changes",
        title="What changed inside the file",
        annotations=read_only("What changed inside the file"),
        description=(
            "Lists what changed in an Office file's VBA since a git revision, one entry per "
            "module, with a unified diff. git diff cannot show this: the file is binary, so a "
            "commit that changed one line and one that replaced the whole project look the "
            "same. Call this before committing, and to review what an agent or a colleague "
            "changed. Defaults to HEAD; pass any revision git understands. Needs the file to "
            "be inside a git repository and tracked in that revision."
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
        module_name: Annotated[
            str,
            Field(default="", description="Only this module. Empty reports every one."),
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

        with project_layer.open_project(path, info) as handle:
            current = {m.name: m for m in project_layer.read_modules(handle, info)}

        with (
            _temporary_copy(blob, path.suffix) as older_path,
            project_layer.open_project(older_path, info) as handle,
        ):
            previous = {m.name: m for m in project_layer.read_modules(handle, info)}

        wanted = module_name.strip().casefold()
        entries: list[dict[str, Any]] = []
        for name in sorted(set(current) | set(previous), key=str.casefold):
            if wanted and name.casefold() != wanted:
                continue
            before = previous.get(name)
            after = current.get(name)
            entry = _entry(name, before, after, include_diff)
            entries.append(entry)

        changed = [e for e in entries if e["status"] != "unchanged"]
        # What a reviewer came for is what changed, so the unchanged modules are
        # the ones dropped when a project is too big to report whole.
        reportable = changed if len(entries) > MAX_ITEMS["modules"] else entries
        shown, note = bound(
            reportable, "modules", "Ask for one module with module_name."
        )
        result: dict[str, Any] = {
            "path": str(path),
            "repository": str(repository),
            "revision": revision,
            "modules_changed": len(changed),
            "modules": shown,
            "verdict": "no VBA changes" if not changed else f"{len(changed)} modules changed",
        }
        if len(entries) > MAX_ITEMS["modules"]:
            result["note"] = (
                f"{len(entries)} modules in the project; only the ones that changed are "
                "listed. " + note
            ).strip()
        elif note:
            result["note"] = note
        if wanted and not entries:
            raise ToolError(
                f"No module named {module_name!r} in either version. Modules now: "
                + (", ".join(sorted(current)) or "(none)")
            )
        return result


@dataclass(frozen=True)
class _Diff:
    status: str
    text: str
    added: int
    removed: int


def _entry(name: str, before: Any, after: Any, include_diff: bool) -> dict[str, Any]:
    if before is None:
        diff = _diff("", after.body, name, "added")
    elif after is None:
        diff = _diff(before.body, "", name, "removed")
    elif _same(before.body, after.body):
        diff = _Diff("unchanged", "", 0, 0)
    else:
        diff = _diff(before.body, after.body, name, "modified")

    entry: dict[str, Any] = {
        "module": name,
        "status": diff.status,
        "lines_added": diff.added,
        "lines_removed": diff.removed,
    }
    if include_diff and diff.text:
        text, was_cut = truncate(
            diff.text, hint="Ask for one module, or set include_diff false."
        )
        entry["diff"] = text
        if was_cut:
            entry["diff_truncated"] = True
    return entry


def _diff(before: str, after: str, name: str, status: str) -> _Diff:
    before_lines = before.replace("\r\n", "\n").splitlines()
    after_lines = after.replace("\r\n", "\n").splitlines()
    lines = list(
        difflib.unified_diff(
            before_lines,
            after_lines,
            fromfile=f"{name} (at the revision)",
            tofile=f"{name} (now)",
            lineterm="",
            n=3,
        )
    )
    added = sum(1 for line in lines if line.startswith("+") and not line.startswith("+++"))
    removed = sum(1 for line in lines if line.startswith("-") and not line.startswith("---"))
    if len(lines) > MAX_DIFF_LINES:
        withheld = len(lines) - MAX_DIFF_LINES
        lines = lines[:MAX_DIFF_LINES]
        lines.append(f"... {withheld} more diff lines. Ask for this module on its own.")
    return _Diff(status=status, text="\n".join(lines), added=added, removed=removed)


def _same(left: str, right: str) -> bool:
    """Line-ending style is not somebody's edit."""
    return left.replace("\r\n", "\n").rstrip() == right.replace("\r\n", "\n").rstrip()


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
