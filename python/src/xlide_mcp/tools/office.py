"""The file in the user's own Office application: is it open, open it, close it.

The one tool group that acts inside applications the user is running, and only
on the document named. It exists because a write fails while the file is open -
Word and PowerPoint lock a document even open read-only - and the user asked for
a way through that does not start with "close it yourself" (xlide_mcp#4, #5).

The rules it keeps, each of which is a refusal rather than a best effort:

* Nothing holding unsaved work is closed unless the call says to save it or to
  discard it, and the descriptions say whose decision that is.
* An application is never quit unless this server started that instance and it
  has nothing else open.
* A process is ended only when end_process says so, only the process Windows
  names as holding the file, and only if it is that file's own application.

A file this server closed is remembered with where the user was in it - the
sheet, selection and scroll in Excel, the selection in Word, the slide in
PowerPoint - and put back there when it is opened again.
"""

from __future__ import annotations

import sys
import time
from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import locks, office_apps
from ..config import Settings
from ..errors import ToolError
from ..hosts import HostInfo, host_info
from ..paths import resolve_path
from ._common import read_only, writes

# How long each operation may take before the application is taken to be waiting
# on a dialog. Opening includes starting the application, which is the slow part.
FIND_TIMEOUT = 20.0
OPEN_TIMEOUT = 90.0
CLOSE_TIMEOUT = 30.0

# Instances this server started, the only ones it may quit: process id to the
# time it started, since Windows gives a finished process's id to a later one.
_STARTED: dict[int, int | None] = {}
# Where the user was in a file this server closed, for the next open of it.
_PLACES: dict[str, dict[str, Any]] = {}

_APPLICATION_HOSTS = frozenset({"excel", "word", "powerpoint", "access"})


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_is_open",
        title="Is the file open in Office",
        annotations=read_only("Is the file open in Office"),
        description=(
            "Says whether a file is open in its Office application, in any running instance: "
            "read-only or for editing, with unsaved changes or not, and which processes hold "
            "it locked. Call it when a write reports the file is locked, before closing "
            "anything, or when the user asks. Windows only. Changes nothing."
        ),
    )
    def is_open(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
    ) -> dict[str, Any]:
        path, info = _office_file(file_path, settings)
        held = locks.holders(path)
        result: dict[str, Any] = {"path": str(path), "application": info.title}
        try:
            copies = office_apps.run("find", _request(path, info), FIND_TIMEOUT)["documents"]
        except office_apps.NotAnswering as exc:
            result.update({"open": "unknown", "not_answering": str(exc)})
            copies = []
        else:
            result["open"] = bool(copies)
        result["copies"] = copies
        result["locked"] = locks.is_locked(path)
        result["held_by"] = [holder.summary() for holder in held or []]
        result["note"] = _state_note(path, info, copies, result["locked"], held or [])
        return result

    @server.tool(
        name="xlide_open_in_app",
        title="Open in Office",
        annotations=writes("Open in Office", destructive=False, idempotent=False),
        description=(
            "Opens a file in its Office application on the user's screen, for them to see or "
            "work in: after a change they should look at, or to put back a copy "
            "xlide_close_in_app closed. Already open, it is brought forward instead. "
            "read_only opens a copy that cannot be saved over the file, and new_instance a "
            "separate Excel or Word the user's other files are not in. bring_to_front=false "
            "reopens it behind whatever the user is doing. Macros follow the user's Trust "
            "Center settings, as for a file they open themselves. Windows only."
        ),
    )
    def open_in_app(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        read_only: Annotated[
            bool, Field(default=False, description="Open a copy that cannot be saved over it.")
        ] = False,
        new_instance: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "Open it in a new instance of the application rather than the one "
                    "running. PowerPoint has only one; Access always gets its own."
                ),
            ),
        ] = False,
        bring_to_front: Annotated[
            bool,
            Field(default=True, description="Bring its window forward. False leaves it behind."),
        ] = True,
    ) -> dict[str, Any]:
        path, info = _office_file(file_path, settings)
        if read_only and info.host == "access":
            raise ToolError(
                "Access has no read-only open through automation, so a database opens for "
                "editing or not at all. Call again without read_only if that is what the user "
                "wants."
            )
        if not read_only and locks.is_locked(path):
            held = locks.holders(path) or []
            here = office_apps.run("find", _request(path, info), FIND_TIMEOUT)["documents"]
            if not here:
                named = locks.describe(held) or f"another {info.title}"
                raise ToolError(
                    f"{path.name} is open for editing in {named}, so a second copy opened for "
                    f"editing would stop on {info.title}'s file-in-use prompt. Open it with "
                    "read_only=true, or ask the user to close the other copy."
                )
        request = _request(path, info)
        request.update(
            {
                "read_only": read_only,
                "new_instance": new_instance,
                "bring_to_front": bring_to_front,
                "place": _PLACES.pop(office_apps.normal_key(path), None),
            }
        )
        answer = office_apps.run("open", request, OPEN_TIMEOUT)
        if answer.get("started_instance") and answer.get("pid"):
            pid = int(answer["pid"])
            _STARTED[pid] = locks.started_at(pid)
        result: dict[str, Any] = {"path": str(path), "application": info.title, **answer}
        if request["place"] and not answer.get("already_open"):
            result["restored_place"] = request["place"]
        if answer.get("already_open"):
            result["note"] = (
                f"It was already open in {answer.get('application', info.title)}"
                + (", read-only" if answer.get("read_only") else "")
                + ", so that copy was brought forward rather than a second one opened."
            )
        elif new_instance and not answer.get("started_instance"):
            result["note"] = (
                f"{info.title} runs one instance, so it opened in the one already running."
            )
        return result

    @server.tool(
        name="xlide_close_in_app",
        title="Close in Office",
        annotations=writes("Close in Office", destructive=True, idempotent=False),
        description=(
            "Closes a file in the Office application the user has it open in, which frees it "
            "for a write. Only that file: the application stays, unless this server started "
            "the instance and nothing else is open in it. A copy holding unsaved work is left "
            "open and reported, unless save_changes saves it first or discard_changes closes "
            "it without saving; losing the user's work is their decision, so ask them. When "
            "the application does not answer, or the file stays locked, end_process ends the "
            "process Windows says holds the file, and every unsaved document in that process "
            "is lost with it: ask the user first. Where the user was in the file is "
            "remembered, and xlide_open_in_app puts them back there. Windows only."
        ),
    )
    def close_in_app(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        save_changes: Annotated[
            bool,
            Field(
                default=False,
                description="Save unsaved changes before closing. Not for a read-only copy.",
            ),
        ] = False,
        discard_changes: Annotated[
            bool,
            Field(
                default=False,
                description="Close without saving, losing unsaved changes. Ask the user first.",
            ),
        ] = False,
        end_process: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "If the application does not answer, or the file stays locked, end the "
                    "process holding it. Loses unsaved work in every document open in it. Ask "
                    "the user first."
                ),
            ),
        ] = False,
    ) -> dict[str, Any]:
        path, info = _office_file(file_path, settings)
        if save_changes and discard_changes:
            raise ToolError("save_changes and discard_changes cannot both be true.")
        request = _request(path, info)
        request.update(
            {
                "save": save_changes,
                "discard": discard_changes,
                "quit_if_empty": _still_started(),
            }
        )
        result: dict[str, Any] = {"path": str(path), "application": info.title}
        try:
            answer = office_apps.run("close", request, CLOSE_TIMEOUT)
        except office_apps.NotAnswering as exc:
            if not end_process:
                raise ToolError(
                    f"{exc} If it has hung, end_process=true ends the process that holds the "
                    "file, and every unsaved document open in it is lost: ask the user first."
                ) from exc
            answer = {"closed": False, "open": True, "not_answering": True}
            result["not_answering"] = str(exc)

        refused = answer.get("refused")
        if refused:
            raise ToolError(_refusal(path, info, refused, answer))
        for copy in answer.get("copies", []):
            if copy.get("place"):
                _PLACES[office_apps.normal_key(path)] = copy["place"]
            if copy.get("quit_instance") and copy.get("pid"):
                _STARTED.pop(int(copy["pid"]), None)
        result["closed"] = bool(answer.get("closed"))
        result["copies"] = answer.get("copies", [])
        failed = [copy for copy in result["copies"] if copy.get("error")]
        if failed:
            result["not_closed"] = [
                f"{copy.get('application', info.title)}: {copy['error']}" for copy in failed
            ]

        locked = _wait_for_release(path)
        if locked and end_process:
            result["ended_processes"] = _end_holders(path, info)
            locked = _wait_for_release(path)
        result["locked"] = locked
        if locked:
            held = locks.holders(path) or []
            result["held_by"] = [holder.summary() for holder in held]
            result["note"] = (
                f"{path.name} is still locked"
                + (f" by {locks.describe(held)}" if held else "")
                + ". "
                + (
                    "end_process ended only the file's own application; ask the user about "
                    "the rest."
                    if end_process
                    else "end_process=true ends the process holding it if it is "
                    f"{info.title}, losing unsaved work in everything open in it: ask the user."
                )
            )
        elif not answer.get("open") and not answer.get("closed") and not end_process:
            result["note"] = f"{path.name} was not open in {info.title}. Nothing was closed."
        elif failed:
            result["note"] = (
                f"Not every copy closed; not_closed says why. {path.name} is not locked, so a "
                "write goes ahead, and a copy still open shows the old content until it is "
                "reopened."
            )
        else:
            result["note"] = (
                f"{path.name} is free to write. xlide_open_in_app puts it back, where the user "
                "was in it; with read_only=true and bring_to_front=false it goes back as a "
                "read-only copy behind what they are doing."
            )
        return result


# ------------------------------------------------------------------ the parts


def _office_file(raw: str, settings: Settings) -> tuple[Path, HostInfo]:
    path = resolve_path(raw, settings)
    info = host_info(path)
    if info.host not in _APPLICATION_HOSTS:
        raise ToolError(
            f"{path.name} is a {info.title} file, which has no Office application to open it in."
        )
    if sys.platform != "win32":
        raise ToolError(
            "Opening and closing a file in its Office application needs Windows with the "
            "application installed. Every tool that reads or writes the file itself works here."
        )
    return path, info


def _request(path: Path, info: HostInfo) -> dict[str, Any]:
    return {"path": str(path), "host": info.host, "application": info.title}


def _state_note(
    path: Path,
    info: HostInfo,
    copies: list[dict[str, Any]],
    locked: bool | None,
    held: list[locks.Holder],
) -> str:
    if not copies and not locked:
        blocker = locks.write_blocker(path)
        if blocker:
            return f"{path.name} is not open in {info.title}, but it {blocker}"
        return (
            f"{path.name} is not open in {info.title} and nothing holds it. A write can go ahead."
        )
    if not copies:
        holder = locks.describe(held) or "a process"
        return (
            f"{path.name} is not open as a document here, but {holder} holds it, so a write "
            "will fail until that lets go."
        )
    unsaved = [copy for copy in copies if copy.get("unsaved")]
    if unsaved:
        return (
            "A copy holds unsaved changes. Closing it would lose them, so that is the user's "
            "decision: xlide_close_in_app with save_changes or discard_changes."
        )
    if all(copy.get("read_only") for copy in copies) and not locked:
        return (
            f"It is open read-only, which in {info.title} does not lock the file: a write goes "
            "ahead, and the open copy shows the old content until it is reopened."
        )
    return (
        "It is open with nothing unsaved, and locked. xlide_close_in_app closes it without "
        "losing anything; ask the user if it is open for editing, since they may be working "
        "in it."
    )


def _refusal(path: Path, info: HostInfo, reason: str, state: dict[str, Any]) -> str:
    where = state.get("application", info.title)
    if reason == "access_unknown":
        return (
            f"{path.name} is open in {where}, which did not say which of its objects hold "
            "unsaved design, so none of them can be saved or discarded from here. Nothing was "
            f"closed. Ask the user to save or close them in {where}."
        )
    if reason == "read_only_save":
        return (
            f"{path.name} is open read-only in {where}, and a read-only copy cannot be saved "
            f"over the file. The user can Save As in {where}, or discard_changes=true closes "
            "it without saving."
        )
    mode = "read-only" if state.get("read_only") else "for editing"
    return (
        f"{path.name} is open {mode} in {where} with unsaved changes, and closing it would "
        "lose them. Nothing was closed. Ask the user: save_changes=true saves them first"
        + (" (not possible for a read-only copy)" if state.get("read_only") else "")
        + ", discard_changes=true closes it without saving."
    )


def _wait_for_release(path: Path) -> bool:
    """Whether the file is still locked, after giving a closing application a moment."""
    for _attempt in range(6):
        if not locks.is_locked(path):
            return False
        time.sleep(0.25)
    return bool(locks.is_locked(path))


def _end_holders(path: Path, info: HostInfo) -> list[dict[str, Any]]:
    """End the processes Windows says hold the file, if each is the file's own application.

    A sync client, a backup agent or another program holding the file is named
    and left alone: ending one of those is not what "close the workbook" meant.
    """
    own = locks.IMAGES.get(info.title, "").casefold()
    ended: list[dict[str, Any]] = []
    for holder in locks.holders(path) or []:
        if not own or holder.image.casefold() != own:
            continue
        if _terminate(holder.pid):
            _STARTED.pop(holder.pid, None)
            ended.append(holder.summary())
    return ended


def _still_started() -> list[int]:
    """The instances this server started that are still the same processes.

    A process id that now belongs to a later process, one the user may have
    started, is forgotten rather than quit.
    """
    for pid, started in list(_STARTED.items()):
        if started is None or locks.started_at(pid) != started:
            del _STARTED[pid]
    return sorted(_STARTED)


def _terminate(pid: int) -> bool:
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    process_terminate = 0x0001
    handle = kernel32.OpenProcess(process_terminate, False, pid)
    if not handle:
        return False
    try:
        return bool(kernel32.TerminateProcess(handle, 1))
    finally:
        kernel32.CloseHandle(handle)
