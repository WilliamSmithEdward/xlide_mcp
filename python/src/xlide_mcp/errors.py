"""The one error type every tool raises, and the helpers that build its message.

A tool failure reaches the calling agent as text and nothing else, so the message
is the whole of what the agent has to work with. Each one names what was refused
and what to do instead; an error that only says something went wrong costs a turn
and teaches nothing.
"""

from __future__ import annotations

from mcp.server.mcpserver.exceptions import ToolError

__all__ = [
    "ToolError",
    "needs_package",
    "needs_windows",
    "not_found",
    "unsupported",
]


def not_found(what: str, *, tried: str, remedy: str) -> ToolError:
    """Something the caller named does not exist."""
    return ToolError(f"{what} not found: {tried}. {remedy}")


def unsupported(what: str, *, supported: str) -> ToolError:
    """Something the caller named is outside what this server handles."""
    return ToolError(f"{what} is not supported. Supported: {supported}.")


def needs_windows(operation: str) -> ToolError:
    """An operation that drives a desktop application, off Windows."""
    return ToolError(
        f"{operation} needs Windows with the desktop Office application installed. "
        "Every tool that only reads or writes the file itself works on any platform; "
        "call xlide_doctor for what this machine can do."
    )


def needs_package(operation: str, package: str, extra: str) -> ToolError:
    """An optional dependency is missing. Name the package and the exact command."""
    return ToolError(
        f"{operation} needs the {package} package, which is not installed. "
        f"Ask the user before installing anything, then run: pip install 'xlide-mcp[{extra}]'"
    )
