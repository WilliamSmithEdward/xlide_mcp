"""Turning a caller's path string into a path this server is allowed to touch.

Every tool that takes a path routes through `resolve_path`. It resolves symlinks
before the containment check, so a link inside a root that points outside one does
not smuggle the target in, and it refuses the traversal and null-byte forms before
the filesystem sees them.

A relative path resolves against the first root rather than the process working
directory. A server started with --root and then run from somewhere else would
otherwise answer differently depending on where it was launched from, which is a
question the caller cannot see the answer to.
"""

from __future__ import annotations

from pathlib import Path

from .config import Settings
from .errors import ToolError


def resolve_path(
    raw: str,
    settings: Settings,
    *,
    must_exist: bool = True,
    must_be_file: bool = True,
) -> Path:
    """Resolve a caller-supplied path, or refuse it with the reason."""
    text = (raw or "").strip()
    if not text:
        raise ToolError("No path given. Pass the absolute path to the file.")
    if "\x00" in text:
        raise ToolError("Path contains a null byte.")

    candidate = Path(text).expanduser()
    roots = settings.effective_roots()
    if not candidate.is_absolute():
        candidate = roots[0] / candidate

    try:
        resolved = candidate.resolve()
    except OSError as exc:
        raise ToolError(f"Path cannot be resolved: {text} ({exc.strerror or exc}).") from exc

    if not settings.allow_outside_roots and not _inside_any(resolved, roots):
        listed = ", ".join(str(r) for r in roots)
        raise ToolError(
            f"Path is outside this server's workspace: {resolved}. "
            f"Allowed roots: {listed}. Ask the user to start the server with "
            "--root pointing at the folder they mean."
        )

    if must_exist and not resolved.exists():
        raise ToolError(
            f"No such file: {resolved}. Call xlide_list_projects to see what is in the workspace."
        )
    if must_exist and must_be_file and not resolved.is_file():
        raise ToolError(f"Not a file: {resolved}.")
    return resolved


def resolve_directory(raw: str, settings: Settings, *, must_exist: bool = True) -> Path:
    """Resolve a path that names a folder."""
    resolved = resolve_path(raw, settings, must_exist=must_exist, must_be_file=False)
    if must_exist and not resolved.is_dir():
        raise ToolError(f"Not a folder: {resolved}.")
    return resolved


def require_writable(settings: Settings, operation: str) -> None:
    """Refuse a mutating tool when the server was started read-only."""
    if settings.read_only:
        raise ToolError(
            f"{operation} was refused: this server is running read-only "
            "(XLIDE_MCP_READ_ONLY). Ask the user to restart it without that setting."
        )


def display(path: Path, settings: Settings) -> str:
    """A path as it is worth showing: relative to its root when that is shorter."""
    for root in settings.effective_roots():
        try:
            relative = path.relative_to(root)
        except ValueError:
            continue
        return str(relative)
    return str(path)


def _inside_any(path: Path, roots: tuple[Path, ...]) -> bool:
    return any(path == root or root in path.parents for root in roots)
