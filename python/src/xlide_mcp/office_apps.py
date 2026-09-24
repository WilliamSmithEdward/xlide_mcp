"""A file open in the user's own Office application: find it, open it, close it.

Everything else in this server works on the file on disk and never touches an
application the user is running. These three operations exist because the user
asked for them (xlide_mcp#4, #5): a write fails while Word has the document open,
even read-only, and "ask the user to close it" is a round trip the agent can
save them, with their agreement.

How a document is found. Excel, Word and PowerPoint register every open document
in the Running Object Table under its full path, hidden and read-only copies
included, measured on Office 16. So the table finds a file in any running
instance, where GetActiveObject reaches only the one instance the application
registered. Access registers its application rather than its database, and is
reached through that.

Why a separate process. Every call here goes into an application the user is
driving, and one with a dialog open does not answer until somebody clicks. Run
in this process, that would hang the server with it. So each operation runs in a
short-lived worker, `python -m xlide_mcp.office_apps`, which reads one JSON
request on stdin, writes one JSON answer on stdout and exits; the server holds a
deadline over it and ends it, not the application, when that passes.

What the worker never does: kill an application, quit one it did not start, or
close a document holding unsaved work unless the request says to save it or to
discard it. Ending a process that holds a file is xlide_close_in_app's last
resort, done in the server with the process named, and only when asked.

Opening applies the user's own macro security. Office runs every macro in a file
it was told to open through automation, and that is not what double-clicking the
file does, so AutomationSecurity is set to follow the Trust Center for the open.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from . import locks
from .errors import ToolError

PROGIDS = {
    "excel": "Excel.Application",
    "word": "Word.Application",
    "powerpoint": "PowerPoint.Application",
    "access": "Access.Application",
}

# The COM answers of an application that is busy rather than gone:
# RPC_E_CALL_REJECTED and RPC_E_SERVERCALL_RETRYLATER.
_BUSY = frozenset({-2147418111, -2147417846})

# msoAutomationSecurityByUI: macros in a file opened through automation are
# treated as the Trust Center says, as they are for a file the user opens.
_SECURITY_BY_UI = 2

# DoCmd.Close's Save argument: acSaveYes and acSaveNo.
_ACCESS_SAVE = {True: 1, False: 2}

_PACKAGE_PARENT = str(Path(__file__).resolve().parent.parent)


def run(operation: str, request: dict[str, Any], timeout: float) -> dict[str, Any]:
    """Run one operation in a worker process and return its answer.

    Raises a ToolError naming the application when it does not answer in time:
    that is almost always a dialog waiting for the user.
    """
    if sys.platform != "win32":
        raise ToolError(
            "Opening and closing a file in its Office application needs Windows with the "
            "application installed."
        )
    payload = json.dumps({"operation": operation, **request})
    try:
        finished = subprocess.run(
            [sys.executable, "-m", "xlide_mcp.office_apps"],
            input=payload,
            capture_output=True,
            text=True,
            timeout=timeout,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            # -m puts the working directory first on the import path, and the
            # server's is wherever it was launched: an xlide_mcp folder there
            # would run instead. This one holds the package the server runs.
            cwd=_PACKAGE_PARENT,
        )
    except subprocess.TimeoutExpired as exc:
        raise NotAnswering(request.get("application", "The application"), timeout) from exc
    lines = [line for line in finished.stdout.splitlines() if line.startswith("{")]
    if not lines:
        detail = (finished.stderr or "").strip().splitlines()[-1:] or ["no answer"]
        raise ToolError(f"The Office worker failed: {detail[0]}")
    answer = json.loads(lines[-1])
    if answer.get("error"):
        raise ToolError(answer["error"])
    return answer


class NotAnswering(ToolError):
    """The application did not answer before the deadline."""

    def __init__(self, application: str, timeout: float) -> None:
        super().__init__(
            f"{application} did not answer within {timeout:g} seconds. Usually a dialog is "
            "open in it, waiting for the user; sometimes it is still busy starting or saving. "
            "Ask the user to look at it, then try again."
        )


# --------------------------------------------------------------- the worker


def main() -> None:
    """One request in, one answer out. Runs in the worker process."""
    request = json.loads(sys.stdin.read() or "{}")
    try:
        answer = _dispatch(request)
    except ToolError as exc:
        answer = {"error": str(exc)}
    except Exception as exc:  # the worker reports rather than dying mute
        answer = {"error": f"{type(exc).__name__}: {exc}"}
    sys.stdout.write(json.dumps(answer, default=str) + "\n")
    sys.stdout.flush()


def _dispatch(request: dict[str, Any]) -> dict[str, Any]:
    try:
        import pythoncom
    except ImportError as exc:
        raise ToolError(
            "This needs pywin32, which comes with xlide-mcp[live]. Ask the user before "
            "installing it."
        ) from exc
    pythoncom.CoInitialize()
    try:
        operation = request.get("operation")
        if operation == "find":
            return {"documents": [_describe(doc, request) for doc in _find(request)]}
        if operation == "open":
            return _open(request)
        if operation == "close":
            return _close(request)
        raise ToolError(f"Unknown operation {operation!r}.")
    finally:
        pythoncom.CoUninitialize()


def _find(request: dict[str, Any]) -> list[Any]:
    """Every open copy of the file, in any running instance."""
    if request["host"] == "access":
        return _find_access(request["path"])
    import pythoncom
    import win32com.client

    table = pythoncom.GetRunningObjectTable()
    context = pythoncom.CreateBindCtx(0)
    found: list[Any] = []
    for moniker in table.EnumRunning():
        try:
            name = moniker.GetDisplayName(context, None)
        except pythoncom.com_error:
            continue
        if not _same_file(name, request["path"]):
            continue
        try:
            running = table.GetObject(moniker)
            found.append(
                win32com.client.Dispatch(running.QueryInterface(pythoncom.IID_IDispatch))
            )
        except pythoncom.com_error:
            continue
    return found


def _find_access(path: str) -> list[Any]:
    """Access registers its application, not its database: one instance is reachable."""
    import win32com.client

    try:
        app = win32com.client.GetActiveObject(PROGIDS["access"])
    except Exception:
        return []
    try:
        current = str(app.CurrentProject.FullName or "")
    except Exception:
        return []
    return [app.CurrentProject] if _same_file(current, path) else []


def _describe(doc: Any, request: dict[str, Any]) -> dict[str, Any]:
    host = request["host"]
    app = doc.Application
    entry: dict[str, Any] = {"application": str(_attempt(lambda: app.Name, "Office"))}
    if host == "access":
        entry.update({"read_only": False, "unsaved": _access_unsaved(app)})
    else:
        entry["read_only"] = _truthy(_attempt(lambda: doc.ReadOnly, False))
        entry["unsaved"] = _unsaved(doc, host)
    pid = _pid(app, doc, host)
    if pid:
        entry["pid"] = pid
    entry["visible"] = _truthy(_attempt(lambda: app.Visible, True))
    others, others_unsaved = _others(app, doc, host)
    entry["other_documents"] = others
    if others_unsaved:
        entry["other_documents_unsaved"] = others_unsaved
    return entry


def _open(request: dict[str, Any]) -> dict[str, Any]:
    import win32com.client

    host = request["host"]
    path = request["path"]
    read_only = bool(request.get("read_only"))
    existing = _find(request)
    if existing:
        doc = existing[0]
        if request.get("bring_to_front", True):
            _front(doc.Application, doc, host)
        return {"already_open": True, **_describe(doc, request)}

    started = False
    app = None
    # Access keeps one database per instance, so attaching would close the one
    # the user has open. Everything else attaches unless a new instance is asked.
    if host != "access" and not request.get("new_instance"):
        try:
            app = win32com.client.GetActiveObject(PROGIDS[host])
        except Exception:
            app = None
    if app is None:
        # PowerPoint runs one instance, and DispatchEx hands back the user's own
        # when it is running, which is not one this server started.
        already = host == "powerpoint" and bool(locks.running(locks.IMAGES["PowerPoint"]))
        app = win32com.client.DispatchEx(PROGIDS[host])
        started = not already

    previous_security = _attempt(lambda: app.AutomationSecurity, None)
    _attempt(lambda: setattr(app, "AutomationSecurity", _SECURITY_BY_UI), None)
    try:
        if host == "excel":
            app.Visible = True
            if started:
                # An instance automation started closes when its last reference
                # goes; one the user controls stays, which is what was asked for.
                app.UserControl = True
            doc = _busy_retry(lambda: app.Workbooks.Open(path, 0, read_only))
        elif host == "word":
            app.Visible = True
            doc = _busy_retry(lambda: app.Documents.Open(path, False, read_only, False))
        elif host == "powerpoint":
            # ReadOnly, Untitled and WithWindow are MsoTriState: -1 is true.
            flag = -1 if read_only else 0
            doc = _busy_retry(lambda: app.Presentations.Open(path, flag, 0, -1))
        else:
            app.Visible = True
            app.UserControl = True
            _busy_retry(lambda: app.OpenCurrentDatabase(path, False))
            doc = app.CurrentProject
    finally:
        if not started and previous_security is not None:
            _attempt(lambda: setattr(app, "AutomationSecurity", previous_security), None)

    place = request.get("place") or {}
    if place and host != "access":
        _restore_place(app, doc, host, place)
    if request.get("bring_to_front", True):
        _front(app, doc, host)
    answer = {"opened": True, "started_instance": started, **_describe(doc, request)}
    return answer


def _close(request: dict[str, Any]) -> dict[str, Any]:
    host = request["host"]
    documents = _find(request)
    if not documents:
        return {"closed": False, "open": False}
    save = bool(request.get("save"))
    discard = bool(request.get("discard"))

    # Every copy is checked before any is closed. A refusal after one copy had
    # closed would report that nothing was, and lose where the user was in it.
    states = [_describe(doc, request) for doc in documents]
    for doc, state in zip(documents, states, strict=True):
        if not state["unsaved"]:
            continue
        if not (save or discard):
            return {"closed": False, "open": True, "refused": "unsaved", **state}
        if save and state["read_only"]:
            return {"closed": False, "open": True, "refused": "read_only_save", **state}
        if host == "access" and _access_pending(doc.Application) is None:
            return {"closed": False, "open": True, "refused": "access_unknown", **state}

    copies: list[dict[str, Any]] = []
    quit_if_empty = set(request.get("quit_if_empty", []))
    for doc, state in zip(documents, states, strict=True):
        app = doc.Application
        keep = save and bool(state["unsaved"])
        place = _capture_place(app, doc, host) if host != "access" else {}
        try:
            if host == "excel":
                doc.Close(keep)
            elif host == "word":
                doc.Close(-1 if keep else 0)
            elif host == "powerpoint":
                if keep:
                    doc.Save()
                doc.Close()
            else:
                _close_access(app, keep)
        except Exception as exc:
            # Reported with the copy rather than raised, so the copies closed
            # before it are still reported closed, with where the user was.
            state.update({"closed": False, "error": f"{type(exc).__name__}: {exc}"})
            copies.append(state)
            continue
        state["closed"] = True
        state["saved_changes"] = keep
        state["discarded_changes"] = bool(state["unsaved"]) and not keep
        if place:
            state["place"] = place
        if state.get("pid") in quit_if_empty:
            remaining, _unsaved_count = _counts(app, host)
            if remaining == 0:
                _attempt(app.Quit, None)
                state["quit_instance"] = True
        copies.append(state)
    closed = [copy for copy in copies if copy.get("closed")]
    return {"closed": bool(closed), "open": len(closed) < len(copies), "copies": copies}


def _close_access(app: Any, keep: bool) -> None:
    """Close the database, saving or discarding each object's unsaved design first.

    Closed one by one with the answer the call gave, so nothing is left for
    Access to decide, or to ask the user about in a dialog nobody here can answer.
    """
    for kind, name in _access_pending(app) or []:
        app.DoCmd.Close(kind, name, _ACCESS_SAVE[keep])
    app.CloseCurrentDatabase()


# ----------------------------------------------------------------- the parts


def _normal(path: str) -> str:
    return os.path.normcase(os.path.abspath(path)) if path else ""


def _same_file(name: str, path: str) -> bool:
    """Whether a name an application registered is the file at `path`.

    The application registers the path it opened, and the server holds the file's
    resolved path, which for a mapped drive is the UNC path and for a SUBST drive
    or a junction is the target. So a name that differs as text is compared as a
    file, but only when the file names agree: an alias keeps its file name, and a
    name on an unreachable share is not waited on for nothing.
    """
    if not name or not path:
        return False
    if _normal(name) == _normal(path):
        return True
    same_name = os.path.basename(name).casefold() == os.path.basename(path).casefold()
    if not same_name or not os.path.isabs(name):
        return False
    try:
        return os.path.samefile(name, path)
    except (OSError, ValueError):
        return False


def _truthy(value: Any) -> bool:
    """A Boolean, or PowerPoint's MsoTriState, where only msoTrue (-1) is true."""
    if isinstance(value, bool):
        return value
    try:
        return int(value) != 0
    except (TypeError, ValueError):
        return bool(value)


def _unsaved(doc: Any, host: str) -> bool:
    """Whether closing would lose work. Anything that cannot be told counts as unsaved."""
    try:
        saved = doc.Saved
    except Exception:
        return True
    if host == "powerpoint":
        return int(saved) != -1
    return not bool(saved)


def _access_unsaved(app: Any) -> bool:
    """Whether any object holds unsaved design. Anything that cannot be told counts."""
    pending = _access_pending(app)
    return pending is None or bool(pending)


def _access_pending(app: Any) -> list[tuple[int, str]] | None:
    """The objects holding unsaved design, as (AcObjectType, name), or None if unreadable.

    A loaded object with its dirty bit set (acSysCmdGetObjectState is 10, dirty
    is 2), and a form, report or module that is open but was never saved, which
    the All collections do not list at all.
    """
    try:
        project, data = app.CurrentProject, app.CurrentData
        groups = {
            0: data.AllTables, 1: data.AllQueries, 2: project.AllForms,
            3: project.AllReports, 4: project.AllMacros, 5: project.AllModules,
        }
        pending: list[tuple[int, str]] = []
        for kind, objects in groups.items():
            for index in range(objects.Count):
                item = objects.Item(index)
                if item.IsLoaded and int(app.SysCmd(10, kind, item.Name)) & 2:
                    pending.append((kind, str(item.Name)))
        for kind, open_now in ((2, app.Forms), (3, app.Reports), (5, app.Modules)):
            saved = {
                str(groups[kind].Item(index).Name).casefold()
                for index in range(groups[kind].Count)
            }
            for index in range(open_now.Count):
                name = str(open_now.Item(index).Name)
                if name.casefold() not in saved:
                    pending.append((kind, name))
    except Exception:
        return None
    return pending


def _counts(app: Any, host: str) -> tuple[int, int]:
    """How many documents an instance holds, and how many of them are unsaved."""
    collection = {"excel": "Workbooks", "word": "Documents", "powerpoint": "Presentations"}
    name = collection.get(host)
    if name is None:
        return 0, 0
    try:
        documents = getattr(app, name)
        total = int(documents.Count)
        unsaved = sum(1 for index in range(1, total + 1) if _unsaved(documents.Item(index), host))
    except Exception:
        return 0, 0
    return total, unsaved


def _others(app: Any, doc: Any, host: str) -> tuple[int, int]:
    total, unsaved = _counts(app, host)
    own_unsaved = 1 if host != "access" and _unsaved(doc, host) else 0
    return max(total - 1, 0), max(unsaved - own_unsaved, 0)


def _pid(app: Any, doc: Any, host: str) -> int:
    hwnd = 0
    if host == "excel":
        hwnd = _attempt(lambda: int(app.Hwnd), 0)
    elif host == "powerpoint":
        hwnd = _attempt(lambda: int(app.HWND), 0)
    elif host == "word":
        hwnd = _attempt(lambda: int(doc.ActiveWindow.Hwnd), 0)
    elif host == "access":
        hwnd = _attempt(lambda: int(app.hWndAccessApp()), 0)
    if not hwnd:
        return 0
    import ctypes
    from ctypes import wintypes

    pid = wintypes.DWORD()
    ctypes.windll.user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(pid))
    return int(pid.value)


def _front(app: Any, doc: Any, host: str) -> None:
    """Bring the document's window forward, as far as Windows lets a background process."""
    if host != "access":
        _attempt(doc.Activate, None)
    hwnd = 0
    if host == "excel":
        hwnd = _attempt(lambda: int(app.Hwnd), 0)
    elif host == "word":
        hwnd = _attempt(lambda: int(app.ActiveWindow.Hwnd), 0)
    elif host == "powerpoint":
        hwnd = _attempt(lambda: int(app.HWND), 0)
    elif host == "access":
        hwnd = _attempt(lambda: int(app.hWndAccessApp()), 0)
    if hwnd:
        import ctypes

        user32 = ctypes.windll.user32
        user32.ShowWindow(hwnd, 9)  # SW_RESTORE
        user32.SetForegroundWindow(hwnd)


def _capture_place(app: Any, doc: Any, host: str) -> dict[str, Any]:
    """Where the user was in the document, so a reopen can put them back there."""
    place: dict[str, Any] = {}
    if host == "excel":
        window = _attempt(lambda: doc.Windows(1), None)
        if window is not None:
            place["sheet"] = _attempt(lambda: str(window.ActiveSheet.Name), None)
            place["scroll_row"] = _attempt(lambda: int(window.ScrollRow), None)
            place["scroll_column"] = _attempt(lambda: int(window.ScrollColumn), None)
            place["selection"] = _attempt(lambda: str(window.RangeSelection.Address()), None)
            place["active_cell"] = _attempt(lambda: str(window.ActiveCell.Address()), None)
    elif host == "word":
        window = _attempt(lambda: doc.ActiveWindow, None)
        if window is not None:
            place["start"] = _attempt(lambda: int(window.Selection.Start), None)
            place["end"] = _attempt(lambda: int(window.Selection.End), None)
            place["scrolled"] = _attempt(lambda: int(window.VerticalPercentScrolled), None)
    elif host == "powerpoint":
        place["slide"] = _attempt(lambda: int(doc.Windows(1).View.Slide.SlideIndex), None)
    return {key: value for key, value in place.items() if value is not None}


def _restore_place(app: Any, doc: Any, host: str, place: dict[str, Any]) -> None:
    """Each step on its own: a sheet that has gone does not stop the scroll."""
    if host == "excel":
        if "sheet" in place:
            _attempt(lambda: doc.Sheets(place["sheet"]).Activate(), None)
        if "selection" in place:
            _attempt(lambda: doc.ActiveSheet.Range(place["selection"]).Select(), None)
        if "active_cell" in place:
            _attempt(lambda: doc.ActiveSheet.Range(place["active_cell"]).Activate(), None)
        window = _attempt(lambda: doc.Windows(1), None)
        if window is not None:
            if "scroll_row" in place:
                _attempt(lambda: setattr(window, "ScrollRow", place["scroll_row"]), None)
            if "scroll_column" in place:
                _attempt(lambda: setattr(window, "ScrollColumn", place["scroll_column"]), None)
    elif host == "word":
        window = _attempt(lambda: doc.ActiveWindow, None)
        if window is not None and "start" in place:
            end = place.get("end", place["start"])
            _attempt(lambda: window.Selection.SetRange(place["start"], end), None)
        if window is not None and "scrolled" in place:
            _attempt(lambda: setattr(window, "VerticalPercentScrolled", place["scrolled"]), None)
    elif host == "powerpoint" and "slide" in place:
        _attempt(lambda: doc.Windows(1).View.GotoSlide(place["slide"]), None)


def _busy_retry(call: Any) -> Any:
    """Retry a call an application rejected because it was busy, briefly."""
    import pythoncom

    for _attempt_number in range(12):
        try:
            return call()
        except pythoncom.com_error as exc:
            if exc.args and exc.args[0] in _BUSY:
                time.sleep(0.25)
                continue
            raise
    return call()


def _attempt(call: Any, fallback: Any) -> Any:
    """A property an application may refuse to answer, with what to use when it does."""
    try:
        return call()
    except Exception:
        return fallback


def normal_key(path: Path) -> str:
    """How the server keys a remembered place: the path as Windows compares it."""
    return _normal(str(path))


if __name__ == "__main__":
    main()
