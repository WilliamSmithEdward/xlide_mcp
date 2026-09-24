"""Who holds a file open, and what to tell the user about it.

A write fails while an Office application has the file open, and "most likely
open in Excel" is a guess that sends the user to the wrong window when the holder
is a second Excel behind a Read-Only title bar, Word, a sync client or a backup
agent. Windows keeps the real answer. Restart Manager lists the processes holding
a file, the same list Explorer's "the file is open in ..." comes from, and it
only lists them: it asks nothing of any application, so a busy or wedged one
cannot hang the lookup.

What holds what, measured on Office build 16.0.20326 (xlide_mcp#4):

* Excel holds a workbook only while it is open for editing. A read-only copy
  holds no handle, so a locked workbook is open for editing somewhere.
* Word and PowerPoint hold a document open read-only as well as for editing.

Windows only. Elsewhere there is nobody to ask, and the message says so rather
than inventing a holder.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

# RmGetList answers this until the array it was given is big enough.
_ERROR_MORE_DATA = 234
# CreateFileW's answers for a file somebody else has open without sharing.
_SHARING_VIOLATIONS = frozenset({32, 33})


@dataclass(frozen=True)
class Holder:
    pid: int
    application: str
    """The name Restart Manager gives, such as "Microsoft Excel"."""
    image: str = ""
    """The executable, such as EXCEL.EXE, when this process may read it."""

    def describe(self) -> str:
        name = self.application or self.image or "a program"
        detail = [self.image] if self.image and self.image != name else []
        detail.append(f"process {self.pid}")
        return f"{name} ({', '.join(detail)})"

    def summary(self) -> dict[str, Any]:
        entry: dict[str, Any] = {"pid": self.pid, "application": self.application}
        if self.image:
            entry["image"] = self.image
        return entry


def describe(holders: list[Holder]) -> str:
    """'Microsoft Excel (EXCEL.EXE, process 4242)', or several joined with 'and'."""
    named = [holder.describe() for holder in holders]
    if len(named) <= 1:
        return "".join(named)
    return ", ".join(named[:-1]) + " and " + named[-1]


def holders(path: str | Path) -> list[Holder] | None:
    """The processes holding a file open, [] for none, or None when Windows cannot say."""
    if sys.platform != "win32":
        return None
    try:
        return _restart_manager(str(Path(path)))
    except (OSError, AttributeError, ValueError):
        return None


def is_locked(path: str | Path) -> bool | None:
    """Whether another process has the file open without sharing it for writing.

    Opened for writing while sharing everything, and closed again at once. That
    fails with a sharing violation exactly when some holder did not share writes,
    which is when a save would fail too; a holder that shares them, such as an
    indexer, is no lock, and the probe itself shuts nobody out. None where it
    cannot be told, which includes a file marked read-only.
    """
    if sys.platform != "win32":
        return None
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    create = kernel32.CreateFileW
    create.argtypes = [
        wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p,
        wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE,
    ]
    create.restype = wintypes.HANDLE
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

    generic_write = 0x40000000
    share_everything = 0x1 | 0x2 | 0x4  # read, write and delete
    open_existing = 3
    handle = create(str(path), generic_write, share_everything, None, open_existing, 0x80, None)
    if handle in (None, wintypes.HANDLE(-1).value):
        error = ctypes.get_last_error()
        if error in _SHARING_VIOLATIONS:
            return True
        return None if error != 2 else False
    kernel32.CloseHandle(handle)
    return False


# The executable each application runs as, for telling it from any other holder.
IMAGES = {
    "Excel": "EXCEL.EXE",
    "Word": "WINWORD.EXE",
    "PowerPoint": "POWERPNT.EXE",
    "Access": "MSACCESS.EXE",
}


def lock_message(path: str | Path, application: str, *, reading: bool = False) -> str:
    """Why this file could not be written, or read, who holds it, and what frees it.

    Called on a PermissionError, which is a lock only when something holds the
    file: a file marked read-only or hidden refuses a save with the same error,
    and "open in Excel" would send the user looking for a window that is not there.
    """
    target = Path(path)
    found = holders(target)
    outcome = "It could not be read." if reading else "Nothing was written."
    if found:
        held = f"{target.name} is locked, held open by {describe(found)}. {outcome}"
        if reading:
            # The advice about each application's windows is about what blocks a
            # save; Excel, for one, does not stop a read.
            return f"{held} Ask the user to close it in the program named."
        own = IMAGES.get(application, "").casefold()
        if own and any(holder.image.casefold() == own for holder in found):
            return f"{held} {advice(application)}"
        # Not the application at all: a sync client, a backup agent, a viewer, or a
        # script. Advice about Excel's windows would send the user to the wrong one.
        return (
            f"{held} It is not {application} that has it. Ask the user to close it in the "
            "program named; a sync client or a backup agent lets go once it has finished."
        )
    blocker = None if reading else write_blocker(target)
    if blocker:
        return f"{target.name} {blocker} {outcome}"
    if found is None:
        return (
            f"{target.name} is locked, most likely open in {application}. {outcome} "
            f"{advice(application)}"
        )
    return (
        f"{target.name} could not be {'read' if reading else 'written'}, and Windows says no "
        f"program has it open, so it is not a lock. {outcome} The user's account may not be "
        "allowed to change it: check its permissions, or ask the user to."
    )


def write_blocker(path: str | Path) -> str | None:
    """What about the file itself refuses a save, or None.

    Every save here writes the file anew, and Windows refuses that for a file
    marked read-only, and for a hidden or system file too.
    """
    try:
        attributes = int(getattr(Path(path).stat(), "st_file_attributes", 0))
    except OSError:
        return None
    if attributes & 0x1:  # FILE_ATTRIBUTE_READONLY
        return (
            "is marked read-only in its properties, so it cannot be written. The user can "
            "clear Read-only in the file's Properties."
        )
    if attributes & (0x2 | 0x4):  # FILE_ATTRIBUTE_HIDDEN, FILE_ATTRIBUTE_SYSTEM
        return (
            "is a hidden file, and Windows refuses to replace a hidden file, which is how a "
            "save writes one. The user can clear Hidden in the file's Properties."
        )
    return None


def started_at(pid: int) -> int | None:
    """When a process started, which tells it from a later one Windows gives the same id."""
    if sys.platform != "win32":
        return None
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

    handle = kernel32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not handle:
        return None
    try:
        times = [wintypes.FILETIME() for _ in range(4)]
        if not kernel32.GetProcessTimes(handle, *(ctypes.byref(time) for time in times)):
            return None
        created = times[0]
        return (int(created.dwHighDateTime) << 32) | int(created.dwLowDateTime)
    finally:
        kernel32.CloseHandle(handle)


def running(image: str) -> set[int]:
    """The processes running an executable, such as POWERPNT.EXE, by id."""
    if sys.platform != "win32" or not image:
        return set()
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.K32EnumProcesses.argtypes = [
        ctypes.POINTER(wintypes.DWORD), wintypes.DWORD, ctypes.POINTER(wintypes.DWORD),
    ]
    size = 1024
    while True:
        ids = (wintypes.DWORD * size)()
        returned = wintypes.DWORD()
        if not kernel32.K32EnumProcesses(ids, ctypes.sizeof(ids), ctypes.byref(returned)):
            return set()
        count = returned.value // ctypes.sizeof(wintypes.DWORD)
        if count < size:
            break
        size *= 2
    wanted = image.casefold()
    return {
        int(ids[index])
        for index in range(count)
        if ids[index] and _image_name(int(ids[index])).casefold() == wanted
    }


def advice(application: str) -> str:
    """What frees the file, which differs by application."""
    if application == "Excel":
        return (
            "Excel holds a workbook only while it is open for editing, so it is open for "
            "editing somewhere: usually the user's own Excel, sometimes a second copy behind a "
            "window that says Read-Only. Ask the user to save and close it there. "
            "xlide_is_open says where it is open and whether it has unsaved changes."
        )
    if application in {"Word", "PowerPoint"}:
        return (
            f"{application} holds a file open read-only as well as for editing, and closing it "
            "there frees it either way. If it is only open read-only with nothing unsaved, "
            "closing it loses nothing: with the user's agreement, xlide_close_in_app closes "
            "it and xlide_open_in_app with read_only=true puts it back after the write. "
            "xlide_is_open says which it is."
        )
    return f"Ask the user to close it in {application}, then try again."


# ----------------------------------------------------------- Restart Manager


def _restart_manager(path: str) -> list[Holder] | None:
    import ctypes
    from ctypes import wintypes

    class UniqueProcess(ctypes.Structure):
        _fields_ = [("dwProcessId", wintypes.DWORD), ("ProcessStartTime", wintypes.FILETIME)]

    # RestartManager.h, 10.0.26100 SDK: CCH_RM_MAX_APP_NAME and CCH_RM_MAX_SVC_NAME
    # plus a terminator each. 668 bytes, every field 4-aligned.
    class ProcessInfo(ctypes.Structure):
        _fields_ = [
            ("Process", UniqueProcess),
            ("strAppName", ctypes.c_wchar * 256),
            ("strServiceShortName", ctypes.c_wchar * 64),
            ("ApplicationType", ctypes.c_int),
            ("AppStatus", wintypes.ULONG),
            ("TSSessionId", wintypes.DWORD),
            ("bRestartable", wintypes.BOOL),
        ]

    manager = ctypes.WinDLL("rstrtmgr.dll")
    manager.RmStartSession.argtypes = [
        ctypes.POINTER(wintypes.DWORD), wintypes.DWORD, wintypes.LPWSTR,
    ]
    manager.RmRegisterResources.argtypes = [
        wintypes.DWORD, wintypes.UINT, ctypes.POINTER(wintypes.LPCWSTR),
        wintypes.UINT, ctypes.c_void_p, wintypes.UINT, ctypes.c_void_p,
    ]
    manager.RmGetList.argtypes = [
        wintypes.DWORD, ctypes.POINTER(wintypes.UINT), ctypes.POINTER(wintypes.UINT),
        ctypes.c_void_p, ctypes.POINTER(wintypes.DWORD),
    ]
    manager.RmEndSession.argtypes = [wintypes.DWORD]

    session = wintypes.DWORD()
    key = ctypes.create_unicode_buffer(33)  # CCH_RM_SESSION_KEY + 1
    if manager.RmStartSession(ctypes.byref(session), 0, key) != 0:
        return None
    try:
        files = (wintypes.LPCWSTR * 1)(path)
        if manager.RmRegisterResources(session, 1, files, 0, None, 0, None) != 0:
            return None
        needed = wintypes.UINT(0)
        count = wintypes.UINT(0)
        reasons = wintypes.DWORD(0)
        found = (ProcessInfo * 0)()
        for _attempt in range(5):
            status = manager.RmGetList(
                session, ctypes.byref(needed), ctypes.byref(count),
                ctypes.cast(found, ctypes.c_void_p), ctypes.byref(reasons),
            )
            if status == 0:
                break
            if status != _ERROR_MORE_DATA:
                return None
            # The list can grow between calls, so the array is sized to what
            # was needed and asked again.
            found = (ProcessInfo * needed.value)()
            count.value = needed.value
        else:
            return None
        return [
            Holder(
                pid=int(found[index].Process.dwProcessId),
                application=found[index].strAppName.strip(),
                image=_image_name(int(found[index].Process.dwProcessId)),
            )
            for index in range(count.value)
        ]
    finally:
        manager.RmEndSession(session)


def _image_name(pid: int) -> str:
    """The executable a process runs, or '' when this process may not ask."""
    import ctypes
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD),
    ]
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

    query_limited_information = 0x1000
    handle = kernel32.OpenProcess(query_limited_information, False, pid)
    if not handle:
        return ""
    try:
        size = wintypes.DWORD(32768)
        buffer = ctypes.create_unicode_buffer(size.value)
        if not kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
            return ""
        return Path(buffer.value).name
    finally:
        kernel32.CloseHandle(handle)
