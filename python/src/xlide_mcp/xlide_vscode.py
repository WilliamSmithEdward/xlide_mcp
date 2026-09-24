"""Telling XLIDE for VS Code what a tool just changed, when it is running.

Inside VS Code with the XLIDE extension, a module an agent writes through this
server should look like one an agent wrote through XLIDE's own tools: marked in
the project tree, with a diff of before and after, and Keep and Revert on it.
XLIDE already keeps that review for its own writes, as a before, an after and
whether the module existed. It cannot see this server's writes except as a file
that changed on disk, which carries no before, so this server sends what it has.

The protocol is in docs/xlide-vscode-bridge.md. In short: XLIDE listens on a
loopback port with a token, and says where, either in XLIDE_VSCODE_API when it
starts this server itself, or in a discovery file per VS Code window, the same
arrangement xlide_vbide uses. After a write this server posts the change to
every window that answers. Nothing here waits on XLIDE or depends on it: a
window that does not answer within a second is skipped, and with no XLIDE
running the whole of this is a directory listing.

What is sent is the module text before and after, the file's path and the
module's name: nothing XLIDE could not read from the file itself, sent to a
port on this machine that only a process knowing the token can use.
"""

from __future__ import annotations

import json
import os
import sys
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from ._version import __version__

ENVIRONMENT = "XLIDE_VSCODE_API"
DISCOVERY_DIRECTORY = "xlide_vscode"
DISCOVERY_GLOB = "xlide-api-*.json"
PROTOCOL = 1

# A notice is best effort. A window that takes longer than this is skipped
# rather than made to hold up the tool call it describes.
TIMEOUT = 1.0
# How long a found or missing window is trusted before looking again.
_CACHE_SECONDS = 10.0
_LOOPBACK = frozenset({"127.0.0.1", "localhost", "::1"})
# urllib's default opener sends 127.0.0.1 through HTTP_PROXY or the system proxy
# like any other host, which would hand the proxy the token and the module text.
_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


@dataclass(frozen=True)
class Window:
    base: str
    """http://127.0.0.1:{port}/{token}, to which a route is appended."""
    folders: tuple[str, ...] = ()
    version: str = ""


_cache: tuple[float, list[Window]] | None = None


def windows() -> list[Window]:
    """Every XLIDE window that answers, found at most every few seconds."""
    global _cache
    now = time.monotonic()
    if _cache is not None and now - _cache[0] < _CACHE_SECONDS:
        return _cache[1]
    try:
        found = [window for window in _candidates() if _hello(window)]
    except OSError:
        # The discovery folder can be unreadable; that is no window, not a failed tool.
        found = []
    _cache = (now, found)
    return found


def forget() -> None:
    """Drop what was found, so the next notice looks again. For tests."""
    global _cache
    _cache = None


def module_written(
    path: Path,
    module: str,
    *,
    before: str,
    before_existed: bool,
    after: str,
    after_exists: bool = True,
    kind: str = "",
    tool: str,
) -> dict[str, Any] | None:
    """A module this server wrote, created or deleted, for XLIDE's review.

    Returns what to put in the tool's result, or None when no window is running.
    """
    body = {
        "file": str(path),
        "module": module,
        "before": before,
        "beforeExisted": before_existed,
        "after": after,
        "afterExists": after_exists,
        "kind": kind,
        "tool": tool,
    }
    return _tell("agent-edit", body)


def module_renamed(path: Path, old: str, new: str, *, tool: str) -> dict[str, Any] | None:
    return _tell("module-renamed", {"file": str(path), "from": old, "to": new, "tool": tool})


def file_changed(path: Path, what: str, *, tool: str = "") -> dict[str, Any] | None:
    """Anything else this server wrote, so the tree shows it without waiting on a watcher."""
    return _tell("file-changed", {"file": str(path), "what": what, "tool": tool})


# ------------------------------------------------------------------ the parts


def _tell(route: str, body: dict[str, Any]) -> dict[str, Any] | None:
    reached = windows()
    if not reached:
        return None
    payload = {**body, "server": "xlide_mcp", "serverVersion": __version__, "protocol": PROTOCOL}
    answers: list[dict[str, Any]] = []
    for window in reached:
        answer = _post(window, route, payload)
        if answer is not None:
            answers.append(answer)
    if not answers:
        return {"notified": False, "note": "XLIDE is running but did not take the notice."}
    shown = any(answer.get("shown") for answer in answers)
    result: dict[str, Any] = {"notified": True, "shown": shown}
    review = next((a.get("review") for a in answers if a.get("review")), None)
    if review:
        result["review"] = review
    if route == "agent-edit" and shown:
        result["note"] = (
            "XLIDE marks this module in its project tree as an agent edit, with the diff, "
            "and the user can keep it or revert it there."
        )
    return result


def _candidates() -> list[Window]:
    found: list[Window] = []
    configured = os.environ.get(ENVIRONMENT, "").strip()
    if configured and _loopback(configured):
        found.append(Window(base=configured.rstrip("/")))
    directory = _discovery_directory()
    if directory is None or not directory.is_dir():
        return found
    for entry in sorted(directory.glob(DISCOVERY_GLOB)):
        try:
            record = json.loads(entry.read_text(encoding="utf-8"))
            port = int(record["port"])
            token = str(record["token"])
        except (OSError, ValueError, KeyError, TypeError):
            continue
        if not token or not 0 < port < 65536:
            continue
        base = f"http://127.0.0.1:{port}/{urllib.parse.quote(token, safe='')}"
        if any(window.base == base for window in found):
            continue
        found.append(
            Window(
                base=base,
                folders=tuple(str(f) for f in record.get("workspaceFolders", []) or []),
                version=str(record.get("version", "")),
            )
        )
    return found


def _discovery_directory() -> Path | None:
    if sys.platform == "win32":
        root = os.environ.get("LOCALAPPDATA")
        return Path(root) / DISCOVERY_DIRECTORY if root else None
    state = os.environ.get("XDG_STATE_HOME") or str(Path.home() / ".local" / "state")
    return Path(state) / DISCOVERY_DIRECTORY


def _loopback(url: str) -> bool:
    """Only ever a port on this machine: module source goes nowhere else."""
    try:
        parsed = urllib.parse.urlparse(url)
    except ValueError:
        return False
    return parsed.scheme == "http" and (parsed.hostname or "") in _LOOPBACK


def _hello(window: Window) -> bool:
    answer = _request(window, "hello", None)
    return bool(answer) and answer.get("product") == "xlide_vscode"


def _post(window: Window, route: str, payload: dict[str, Any]) -> dict[str, Any] | None:
    return _request(window, route, payload)


def _request(window: Window, route: str, payload: dict[str, Any] | None) -> dict[str, Any] | None:
    if not _loopback(window.base):
        return None
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    request = urllib.request.Request(
        f"{window.base}/{route}",
        data=data,
        method="POST" if data is not None else "GET",
        headers={"Content-Type": "application/json"} if data is not None else {},
    )
    try:
        with _OPENER.open(request, timeout=TIMEOUT) as response:
            answer = json.loads(response.read().decode("utf-8") or "{}")
    except Exception:
        # A notice goes after the write it describes, which has already landed.
        # A discovery file outlives its window, and whatever holds that port now
        # can answer anything at all, so any failure here is only "not told".
        return None
    return answer if isinstance(answer, dict) else None
