"""Server settings, and the workspace roots that bound every path a tool accepts.

The roots are the security boundary. An MCP server reads and writes files on
behalf of a model that can be steered by whatever it reads, so a path argument is
untrusted input: it is resolved, symlinks and all, and refused unless it lands
inside a root. Nothing here is a judgement about intent, only about reach.

Defaults come from the process environment so a client that can only set env vars
configures the server as fully as one that passes arguments.
"""

from __future__ import annotations

import os
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field, replace
from pathlib import Path

ENV_ROOTS = "XLIDE_MCP_ROOTS"
ENV_READ_ONLY = "XLIDE_MCP_READ_ONLY"
ENV_ALLOW_OUTSIDE_ROOTS = "XLIDE_MCP_ALLOW_OUTSIDE_ROOTS"
ENV_TIMEOUT = "XLIDE_MCP_TIMEOUT"

# A run that outlives this is a wedged Office application, not slow work. The
# harness holds its own deadline; this is the ceiling a caller may ask for.
MAX_TIMEOUT_SECONDS = 900.0
DEFAULT_TIMEOUT_SECONDS = 120.0


def _truthy(value: str | None) -> bool:
    return (value or "").strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class Settings:
    """What the server is allowed to reach, and for how long."""

    roots: tuple[Path, ...] = field(default_factory=tuple)
    """Directories a path argument may resolve inside. Empty means cwd."""

    read_only: bool = False
    """Refuse every tool that changes a file. Reads and analysis still work."""

    allow_outside_roots: bool = False
    """Accept absolute paths anywhere. Off by default, and a deliberate choice."""

    default_timeout: float = DEFAULT_TIMEOUT_SECONDS
    """Deadline applied to a run that does not name its own."""

    @classmethod
    def from_environment(cls, env: dict[str, str] | None = None) -> Settings:
        source = os.environ if env is None else env
        raw_roots = source.get(ENV_ROOTS, "")
        roots = _normalize_roots(p for p in raw_roots.split(os.pathsep) if p.strip())
        timeout = DEFAULT_TIMEOUT_SECONDS
        raw_timeout = source.get(ENV_TIMEOUT, "").strip()
        if raw_timeout:
            try:
                timeout = max(1.0, min(MAX_TIMEOUT_SECONDS, float(raw_timeout)))
            except ValueError:
                timeout = DEFAULT_TIMEOUT_SECONDS
        return cls(
            roots=roots,
            read_only=_truthy(source.get(ENV_READ_ONLY)),
            allow_outside_roots=_truthy(source.get(ENV_ALLOW_OUTSIDE_ROOTS)),
            default_timeout=timeout,
        )

    def with_roots(self, roots: Iterable[str | Path]) -> Settings:
        return replace(self, roots=_normalize_roots(roots))

    def effective_roots(self) -> tuple[Path, ...]:
        """The roots in force. No configured root means the working directory."""
        if self.roots:
            return self.roots
        return (Path.cwd().resolve(),)

    def describe(self) -> str:
        roots = ", ".join(str(r) for r in self.effective_roots())
        parts = [f"roots: {roots}"]
        if self.read_only:
            parts.append("read-only")
        if self.allow_outside_roots:
            parts.append("paths outside the roots allowed")
        return "; ".join(parts)


def _normalize_roots(roots: Iterable[str | Path]) -> tuple[Path, ...]:
    """Resolve each root once, drop duplicates, keep the order given."""
    seen: dict[Path, None] = {}
    for root in roots:
        text = str(root).strip()
        if not text:
            continue
        resolved = Path(text).expanduser()
        try:
            resolved = resolved.resolve()
        except OSError:
            resolved = resolved.absolute()
        seen.setdefault(resolved, None)
    return tuple(seen)


def clamp_timeout(requested: float | None, settings: Settings) -> float:
    """A caller's deadline, held between one second and the ceiling."""
    if requested is None:
        return settings.default_timeout
    return max(1.0, min(MAX_TIMEOUT_SECONDS, float(requested)))


def roots_from_argv(values: Sequence[str] | None) -> tuple[Path, ...]:
    """--root may repeat, and each value may itself hold os.pathsep-separated paths."""
    if not values:
        return ()
    flattened: list[str] = []
    for value in values:
        flattened.extend(part for part in value.split(os.pathsep) if part.strip())
    return _normalize_roots(flattened)
