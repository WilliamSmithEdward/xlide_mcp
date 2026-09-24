"""Naming what holds a locked file, and saying what frees it.

A write refusal is the whole of what an agent has to relay, and "most likely open
in Excel" sends a user to the wrong window whenever the holder is something else.
On Windows the holder is asked of Restart Manager; the tests hold a file open
from a child process, without sharing, the way an application holds a document.
"""

from __future__ import annotations

import subprocess
import sys
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any

import pytest
from conftest import SAMPLE_MODULE, ToolFailure

from xlide_mcp import locks

windows = pytest.mark.skipif(sys.platform != "win32", reason="Restart Manager is Windows only")

# Opened for reading and, by default, shared for reading only, which is how an
# Office application holds a document: anyone may read it, nobody may write it.
_HOLD = """
import ctypes, sys, time
from ctypes import wintypes
create = ctypes.WinDLL("kernel32", use_last_error=True).CreateFileW
create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p,
                   wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
create.restype = wintypes.HANDLE
handle = create(sys.argv[1], 0x80000000, int(sys.argv[2]), None, 3, 0x80, None)
print("held" if handle not in (None, wintypes.HANDLE(-1).value) else "failed", flush=True)
time.sleep(60)
"""
SHARE_READ = 1
SHARE_EVERYTHING = 7


@pytest.fixture
def hold() -> Iterator[Callable[..., subprocess.Popen[str]]]:
    """Hold a file open from a child process, as an application would, until the test ends."""
    children: list[subprocess.Popen[str]] = []

    def start(path: Path, share: int = SHARE_READ) -> subprocess.Popen[str]:
        child = subprocess.Popen(
            [sys.executable, "-c", _HOLD, str(path), str(share)],
            stdout=subprocess.PIPE,
            text=True,
        )
        children.append(child)
        assert child.stdout is not None
        assert child.stdout.readline().strip() == "held"
        return child

    yield start
    for child in children:
        child.kill()
        child.wait()


def test_several_holders_read_as_a_sentence() -> None:
    found = [
        locks.Holder(1, "Microsoft Excel", "EXCEL.EXE"),
        locks.Holder(2, "OneDrive", "OneDrive.exe"),
        locks.Holder(3, "", "backup.exe"),
    ]
    assert locks.describe(found[:1]) == "Microsoft Excel (EXCEL.EXE, process 1)"
    assert locks.describe(found) == (
        "Microsoft Excel (EXCEL.EXE, process 1), OneDrive (OneDrive.exe, process 2) "
        "and backup.exe (process 3)"
    )


def test_the_advice_follows_the_application_that_holds_it(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Excel locks only what is open for editing, Word and PowerPoint lock a
    read-only copy too, and a holder that is neither needs neither answer."""
    path = tmp_path / "Report.docm"
    word = locks.Holder(4242, "Microsoft Word", "WINWORD.EXE")
    monkeypatch.setattr(locks, "holders", lambda _: [word])
    message = locks.lock_message(path, "Word")
    assert "held open by Microsoft Word (WINWORD.EXE, process 4242)" in message
    assert "read-only as well as for editing" in message
    assert "xlide_close_in_app" in message

    monkeypatch.setattr(locks, "holders", lambda _: [locks.Holder(7, "OneDrive", "OneDrive.exe")])
    message = locks.lock_message(tmp_path / "Book.xlsm", "Excel")
    assert "It is not Excel that has it" in message
    assert "Read-Only" not in message

    monkeypatch.setattr(locks, "holders", lambda _: None)
    message = locks.lock_message(tmp_path / "Book.xlsm", "Excel")
    assert "most likely open in Excel" in message
    assert "open for editing somewhere" in message


def test_a_read_refusal_names_the_holder_without_the_advice_about_saving(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """The advice about Excel's windows is about what stops a save, and Excel does not
    stop a read, so it would send the user to close the wrong thing."""
    excel = locks.Holder(4242, "Microsoft Excel", "EXCEL.EXE")
    monkeypatch.setattr(locks, "holders", lambda _: [excel])
    message = locks.lock_message(tmp_path / "Book.xlsm", "Excel", reading=True)
    assert "held open by Microsoft Excel" in message
    assert "It could not be read" in message
    assert "open for editing somewhere" not in message


def test_a_refusal_nothing_holds_is_not_called_a_lock(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """Windows answering that nothing holds the file is not Windows being unable to say."""
    path = tmp_path / "Book.xlsm"
    path.write_bytes(b"x")
    monkeypatch.setattr(locks, "holders", lambda _: [])
    message = locks.lock_message(path, "Excel")
    assert "no program has it open, so it is not a lock" in message
    assert "open in Excel" not in message


@windows
def test_a_holder_that_shares_writes_is_no_lock(
    hold: Callable[..., subprocess.Popen[str]], tmp_path: Path
) -> None:
    """An indexer or a viewer can hold a file and still share it for writing, and then a
    save goes ahead. Probing with no sharing at all called that a lock, and for that
    instant shut everyone else out of the file as well."""
    path = tmp_path / "Shared.xlsm"
    path.write_bytes(b"x")
    hold(path, SHARE_EVERYTHING)
    assert locks.is_locked(path) is False
    path.write_bytes(b"y")


@windows
def test_a_file_marked_read_only_is_not_called_open_in_excel(
    call: Callable[..., Any], workbook: Path
) -> None:
    import os
    import stat

    os.chmod(workbook, stat.S_IREAD)
    try:
        with pytest.raises(ToolFailure) as refusal:
            call(
                "xlide_write_module",
                file_path=str(workbook),
                module_name="Helpers",
                source=SAMPLE_MODULE + "\n' changed\n",
            )
    finally:
        os.chmod(workbook, stat.S_IREAD | stat.S_IWRITE)
    assert "is marked read-only in its properties" in refusal.value.message
    assert "open in Excel" not in refusal.value.message


@windows
def test_a_hidden_file_says_why_it_cannot_be_saved(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    """Windows refuses to replace a hidden file, which is how every save here writes."""
    import ctypes

    hidden, normal = 0x2, 0x80
    assert ctypes.windll.kernel32.SetFileAttributesW(str(plain_workbook), hidden)
    try:
        with pytest.raises(ToolFailure) as refusal:
            call(
                "xlide_write_cells",
                file_path=str(plain_workbook),
                sheet="Sheet1",
                start_cell="A1",
                data=[[1]],
            )
    finally:
        ctypes.windll.kernel32.SetFileAttributesW(str(plain_workbook), normal)
    assert "is a hidden file" in refusal.value.message


@windows
def test_a_held_file_names_its_holder(
    hold: Callable[[Path], subprocess.Popen[str]], tmp_path: Path
) -> None:
    path = tmp_path / "Held.xlsm"
    path.write_bytes(b"x")
    assert locks.is_locked(path) is False
    assert locks.holders(path) == []

    hold(path)
    assert locks.is_locked(path) is True
    found = locks.holders(path)
    assert found, "Restart Manager lists the child holding the file"
    assert all(holder.image.lower().startswith("python") for holder in found)


@windows
def test_a_write_refused_by_a_lock_names_the_holder_and_writes_nothing(
    call: Callable[..., Any],
    hold: Callable[[Path], subprocess.Popen[str]],
    workbook: Path,
) -> None:
    before = workbook.read_bytes()
    hold(workbook)
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_write_module",
            file_path=str(workbook),
            module_name="Helpers",
            source=SAMPLE_MODULE + "\n' changed\n",
        )
    assert "Budget.xlsm is locked, held open by" in refusal.value.message
    assert "(python" in refusal.value.message.lower()
    assert "Nothing was written" in refusal.value.message
    assert workbook.read_bytes() == before


@windows
def test_a_cell_write_refused_by_a_lock_names_the_holder(
    call: Callable[..., Any],
    hold: Callable[[Path], subprocess.Popen[str]],
    plain_workbook: Path,
) -> None:
    hold(plain_workbook)
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_write_cells",
            file_path=str(plain_workbook),
            sheet="Sheet1",
            start_cell="A1",
            data=[[1]],
        )
    assert "Data.xlsx is locked, held open by" in refusal.value.message
