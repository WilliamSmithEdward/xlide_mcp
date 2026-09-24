"""Opening and closing a file in the user's Office application.

The live cases start their own instances, behind the user's windows, and close
what they opened; nothing here touches an application or a file the user has.
The rest pin the refusals, which is where the care is: unsaved work is never
closed unasked, and a process is ended only when it is the file's own
application.
"""

from __future__ import annotations

import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import SAMPLE_MODULE, ToolFailure

from xlide_mcp import locks

windows = pytest.mark.skipif(sys.platform != "win32", reason="Office runs on Windows")


def test_a_file_with_no_office_application_is_refused(
    call: Callable[..., Any], workspace: Path
) -> None:
    project = workspace / "Legacy.vbp"
    project.write_text('Type=Exe\nName="Legacy"\n', encoding="latin-1")
    for tool in ("xlide_is_open", "xlide_open_in_app", "xlide_close_in_app"):
        with pytest.raises(ToolFailure) as refusal:
            call(tool, file_path=str(project))
        assert "no Office application" in refusal.value.message


@windows
def test_saving_and_discarding_at_once_is_refused(
    call: Callable[..., Any], workbook: Path
) -> None:
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_close_in_app",
            file_path=str(workbook),
            save_changes=True,
            discard_changes=True,
        )
    assert "cannot both be true" in refusal.value.message


@windows
def test_unsaved_work_is_left_open_and_the_choice_named(
    call: Callable[..., Any], workbook: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from xlide_mcp import office_apps

    answer = {
        "closed": False,
        "open": True,
        "refused": "unsaved",
        "application": "Microsoft Word",
        "read_only": True,
        "unsaved": True,
    }
    monkeypatch.setattr(office_apps, "run", lambda *_args: answer)
    with pytest.raises(ToolFailure) as refusal:
        call("xlide_close_in_app", file_path=str(workbook))
    message = refusal.value.message
    assert "open read-only in Microsoft Word with unsaved changes" in message
    assert "Nothing was closed" in message
    assert "discard_changes=true" in message
    assert "not possible for a read-only copy" in message


@windows
def test_only_the_files_own_application_is_ended(
    workbook: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A sync client or a backup agent holding the workbook is named and left
    alone: ending one is not what closing a workbook means."""
    from xlide_mcp.hosts import host_info
    from xlide_mcp.tools import office

    monkeypatch.setattr(
        locks,
        "holders",
        lambda _: [
            locks.Holder(101, "Microsoft Excel", "EXCEL.EXE"),
            locks.Holder(202, "OneDrive", "OneDrive.exe"),
        ],
    )
    ended: list[int] = []
    monkeypatch.setattr(office, "_terminate", lambda pid: ended.append(pid) or True)
    result = office._end_holders(workbook, host_info(workbook))
    assert ended == [101]
    assert [entry["pid"] for entry in result] == [101]


class _Copy:
    """A document the worker found: it records being closed, or refuses to be."""

    def __init__(self, fails: bool = False) -> None:
        self.Application = self
        self.closed = False
        self.fails = fails

    def Close(self, *_args: Any) -> None:
        if self.fails:
            raise RuntimeError("the application refused")
        self.closed = True


def test_every_copy_is_checked_before_any_is_closed(monkeypatch: pytest.MonkeyPatch) -> None:
    """The same workbook open read-only in one Excel and with unsaved work in another.
    Closing copy by copy closed the first, then refused on the second and reported
    that nothing was closed, losing where the user had been in the first."""
    from xlide_mcp import office_apps

    clean, unsaved = _Copy(), _Copy()
    states = {id(clean): {"read_only": True, "unsaved": False},
              id(unsaved): {"read_only": False, "unsaved": True}}
    monkeypatch.setattr(office_apps, "_find", lambda _request: [clean, unsaved])
    monkeypatch.setattr(office_apps, "_describe", lambda doc, _request: dict(states[id(doc)]))
    request = {"host": "excel", "path": "Book.xlsm"}

    answer = office_apps._close(request)
    assert answer["refused"] == "unsaved"
    assert not clean.closed and not unsaved.closed

    answer = office_apps._close({**request, "discard": True})
    assert answer["closed"] is True and answer["open"] is False
    assert clean.closed and unsaved.closed
    assert [copy["discarded_changes"] for copy in answer["copies"]] == [False, True]


def test_a_copy_that_will_not_close_is_reported_with_the_rest(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from xlide_mcp import office_apps

    first, stuck = _Copy(), _Copy(fails=True)
    monkeypatch.setattr(office_apps, "_find", lambda _request: [first, stuck])
    monkeypatch.setattr(
        office_apps, "_describe", lambda _doc, _request: {"read_only": False, "unsaved": False}
    )
    answer = office_apps._close({"host": "excel", "path": "Book.xlsm"})
    assert answer["closed"] is True and answer["open"] is True
    assert [copy["closed"] for copy in answer["copies"]] == [True, False]
    assert "the application refused" in answer["copies"][1]["error"]


class _Objects:
    def __init__(self, *names: str, loaded: tuple[str, ...] = ()) -> None:
        self.items = [
            type("AccessObject", (), {"Name": name, "IsLoaded": name in loaded})()
            for name in names
        ]

    @property
    def Count(self) -> int:
        return len(self.items)

    def Item(self, index: int) -> Any:
        return self.items[index]


class _Access:
    """An Access holding a saved form with unsaved design, and a form never saved."""

    def __init__(self) -> None:
        self.calls: list[tuple[Any, ...]] = []
        forms = _Objects("Orders", "Customers", loaded=("Orders",))
        self.CurrentProject = type("Project", (), {
            "AllForms": forms, "AllReports": _Objects(), "AllMacros": _Objects(),
            "AllModules": _Objects("Helpers"),
        })()
        self.CurrentData = type("Data", (), {"AllTables": _Objects(), "AllQueries": _Objects()})()
        self.Forms = _Objects("Orders", "Form1")
        self.Reports = _Objects()
        self.Modules = _Objects()
        access = self
        self.DoCmd = type("DoCmd", (), {
            "Close": staticmethod(lambda *args: access.calls.append(("Close", *args))),
        })()

    def SysCmd(self, _action: int, kind: int, name: str) -> int:
        return 2 if (kind, name) == (2, "Orders") else 1

    def CloseCurrentDatabase(self) -> None:
        self.calls.append(("CloseCurrentDatabase",))


def test_access_closes_each_unsaved_object_with_the_answer_the_call_gave() -> None:
    """CloseCurrentDatabase alone left Access to decide, or to ask the user, about each
    object with unsaved design, and the result reported the call's flags as if they
    had been applied. A form open but never saved is in no All collection."""
    from xlide_mcp import office_apps

    access = _Access()
    assert office_apps._access_pending(access) == [(2, "Orders"), (2, "Form1")]
    office_apps._close_access(access, keep=False)
    assert access.calls == [
        ("Close", 2, "Orders", 2), ("Close", 2, "Form1", 2), ("CloseCurrentDatabase",)
    ]
    access.calls.clear()
    office_apps._close_access(access, keep=True)
    assert access.calls[:2] == [("Close", 2, "Orders", 1), ("Close", 2, "Form1", 1)]


@windows
def test_a_name_registered_through_a_junction_is_the_same_file(tmp_path: Path) -> None:
    """Office registers the path it opened; the server holds the resolved one. Through
    a mapped drive, a SUBST drive or a junction the two differ as text."""
    import subprocess

    from xlide_mcp import office_apps

    target = tmp_path / "real"
    target.mkdir()
    book = target / "Book.xlsm"
    book.write_bytes(b"x")
    link = tmp_path / "linked"
    subprocess.run(
        ["cmd", "/c", "mklink", "/J", str(link), str(target)], check=True, capture_output=True
    )
    assert office_apps._same_file(str(link / "Book.xlsm"), str(book))
    assert office_apps._same_file(str(link / "BOOK.XLSM"), str(book))
    assert not office_apps._same_file(str(link / "Other.xlsm"), str(book))
    assert not office_apps._same_file("!{00020819-0000-0000-C000-000000000046}", str(book))


@windows
def test_the_worker_runs_this_package_whatever_the_working_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """-m puts the working directory first on the import path, so an xlide_mcp folder
    wherever the server was launched ran in place of the worker."""
    pytest.importorskip("pythoncom")
    from xlide_mcp import office_apps

    planted = tmp_path / "xlide_mcp"
    planted.mkdir()
    (planted / "__init__.py").write_text("", encoding="utf-8")
    (planted / "office_apps.py").write_text(
        "print('{\"documents\": [{\"planted\": true}]}')\n", encoding="utf-8"
    )
    monkeypatch.chdir(tmp_path)
    request = {"path": str(tmp_path / "Nowhere.xlsm"), "host": "excel", "application": "Excel"}
    assert office_apps.run("find", request, 20)["documents"] == []


@windows
def test_an_instance_is_quit_only_while_its_process_id_is_the_same_process(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Windows gives a finished process's id to a later one, which may be an Excel the
    user started. Only the process that started at the recorded time is quit."""
    import os

    from xlide_mcp.tools import office

    pid = os.getpid()
    started = locks.started_at(pid)
    assert started is not None
    monkeypatch.setattr(office, "_STARTED", {pid: started})
    assert office._still_started() == [pid]
    monkeypatch.setattr(office, "_STARTED", {pid: started + 1})
    assert office._still_started() == []
    assert office._STARTED == {}


# ------------------------------------------------------------------- live


@windows
@pytest.mark.live
def test_a_workbook_opens_closes_and_comes_back_where_it_was(
    call: Callable[..., Any], workbook: Path
) -> None:
    assert call("xlide_is_open", file_path=str(workbook))["open"] is False

    opened = call(
        "xlide_open_in_app", file_path=str(workbook), new_instance=True, bring_to_front=False
    )
    try:
        assert opened["opened"] is True
        assert opened["started_instance"] is True
        state = call("xlide_is_open", file_path=str(workbook))
        assert state["open"] is True
        assert state["locked"] is True
        assert state["copies"][0]["read_only"] is False
        assert any(holder["image"] == "EXCEL.EXE" for holder in state["held_by"])

        with pytest.raises(ToolFailure) as refusal:
            call(
                "xlide_write_module",
                file_path=str(workbook),
                module_name="Helpers",
                source=SAMPLE_MODULE + "\n' changed\n",
            )
        assert "held open by Microsoft Excel (EXCEL.EXE" in refusal.value.message

        _make_dirty(workbook)
        with pytest.raises(ToolFailure) as unsaved:
            call("xlide_close_in_app", file_path=str(workbook))
        assert "with unsaved changes" in unsaved.value.message
    finally:
        closed = call("xlide_close_in_app", file_path=str(workbook), discard_changes=True)
    assert closed["closed"] is True
    assert closed["copies"][0]["discarded_changes"] is True
    assert closed["copies"][0]["quit_instance"] is True
    assert closed["locked"] is False

    reopened = call(
        "xlide_open_in_app", file_path=str(workbook), new_instance=True, bring_to_front=False
    )
    try:
        assert reopened["restored_place"]["sheet"] == "Sheet1"
    finally:
        call("xlide_close_in_app", file_path=str(workbook))


@windows
@pytest.mark.live
def test_a_read_only_word_copy_moves_aside_for_a_write(
    call: Callable[..., Any], workspace: Path
) -> None:
    """Word locks a document open read-only as well, so a write fails while the
    user only has it open to look at. Closing that copy loses nothing, and it
    goes back read-only, behind, once the write is done."""
    import pyopenvba

    document = workspace / "Report.docm"
    with pyopenvba.WordFile.create_new(document) as word:
        word.save()

    call(
        "xlide_open_in_app",
        file_path=str(document),
        read_only=True,
        new_instance=True,
        bring_to_front=False,
    )
    try:
        with pytest.raises(ToolFailure) as refusal:
            call(
                "xlide_write_module",
                file_path=str(document),
                module_name="Helpers",
                source=SAMPLE_MODULE,
            )
        assert "read-only as well as for editing" in refusal.value.message
    finally:
        closed = call("xlide_close_in_app", file_path=str(document))
    assert closed["copies"][0]["read_only"] is True
    assert closed["copies"][0]["discarded_changes"] is False

    written = call(
        "xlide_write_module", file_path=str(document), module_name="Helpers", source=SAMPLE_MODULE
    )
    assert written["saved"] is True
    back = call(
        "xlide_open_in_app",
        file_path=str(document),
        read_only=True,
        new_instance=True,
        bring_to_front=False,
    )
    try:
        assert back["read_only"] is True
    finally:
        call("xlide_close_in_app", file_path=str(document))


def _make_dirty(path: Path) -> None:
    """Type into the open workbook through the running object, as the user would."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        table = pythoncom.GetRunningObjectTable()
        context = pythoncom.CreateBindCtx(0)
        for moniker in table.EnumRunning():
            if moniker.GetDisplayName(context, None).lower() == str(path).lower():
                book = win32com.client.Dispatch(
                    table.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch)
                )
                book.Worksheets(1).Range("B3").Value = 42
                return
    finally:
        pythoncom.CoUninitialize()
    raise AssertionError(f"{path.name} is not in the running object table")
