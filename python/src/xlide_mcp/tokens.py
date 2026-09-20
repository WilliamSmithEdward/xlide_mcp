"""Content tokens: the guard that stops one agent's write discarding another's.

`xlide_write_module` takes a module's whole source, so the blast radius of a stale
read is the entire module rather than the lines being changed. An agent reads,
reasons for a while, and writes; in between, the user can have edited the module in
the VBE, another agent can have written it, or a test run can have replaced it.

A token is a hash of the source as it was read. Passing it back makes the write
conditional: same token, the write proceeds; different token, it is refused and the
caller re-reads. The token ignores line-ending style, because the VBA project stores
CRLF and callers hold LF, and a round trip through an agent is not someone's edit.

Ported from XLIDE's moduleContentToken.ts so a token means the same thing in both
products and an agent moving between them does not have to learn two guards.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass

TOKEN_PREFIX = "xlide1:"


def content_token(source: str) -> str:
    """A stable token for a module's source, across line-ending style."""
    normalized = source.replace("\r\n", "\n")
    digest = hashlib.sha256(normalized.encode("utf-8")).hexdigest()
    return TOKEN_PREFIX + digest[:32]


@dataclass(frozen=True)
class StaleWrite:
    expected: str
    actual: str
    message: str


def check_content_token(
    current_source: str,
    expected_token: str | None,
    module_name: str,
) -> StaleWrite | None:
    """None when the write may proceed, including when no token was supplied.

    The guard is opt-in, which is the same call XLIDE made: a caller that has just
    created a module has nothing to compare against, and refusing every untokened
    write would make the first write of a new module impossible.
    """
    if not expected_token:
        return None
    actual = content_token(current_source)
    if actual == expected_token:
        return None
    return StaleWrite(
        expected=expected_token,
        actual=actual,
        message=(
            f'Module "{module_name}" changed since it was read, so the write was refused '
            "to avoid discarding that change. Read the module again, reapply your edit to "
            f"the current source, and write with the new content_token ({actual})."
        ),
    )
