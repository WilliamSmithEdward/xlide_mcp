"""Running VBA: macros, source and tests, in an application this server owns.

Three failure modes make Office automation different from calling a function, and
pyVBAharness is here because it handles all three. A VBA runtime error opens a
modal dialog and waits forever. `Application.Run` takes no timeout. And an Office
application started over COM is not a child process, so killing the caller leaves
EXCEL.EXE running with the document open.

The harness starts an instance it owns, never attaches to one the user is running,
and holds a deadline over every call. That ownership rule is the reason these
tools can enforce a timeout at all: the process terminated on expiry is one this
server created. It is also why nothing here will ever touch an application the
user has open.

Every run opens its document read-only. Running a macro that writes to a workbook
and expecting the workbook to change is a reasonable thing to want and a dangerous
default, so it is `read_only=false`, said out loud by the caller.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from ..config import Settings, clamp_timeout
from ..errors import ToolError, needs_package, needs_windows
from ..hosts import HostInfo, require_openable, require_readable
from ..paths import require_writable, resolve_path
from ._common import read_only, writes

# How long to wait for the previous session to let go before giving up. An
# application takes a couple of seconds to exit after its session closes, and a
# caller making two tool calls in a row should not have to know that.
SESSION_LOCK_WAIT_SECONDS = 60.0

_SESSION_CLASSES = {
    "excel": "ExcelSession",
    "word": "WordSession",
    "powerpoint": "PowerPointSession",
    "access": "AccessSession",
}


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_run_macro",
        title="Run a macro",
        annotations=writes("Run a macro", destructive=True, idempotent=False),
        description=(
            "Runs a procedure that already exists in an Office file, in an application this "
            "server starts and owns, under a deadline. Windows only, with the application "
            "installed. The document opens read-only unless read_only=false, and closes "
            "without saving unless save=true. Returns the procedure's return value, anything "
            "it logged, and, on a VBA error, the error number and message. The failing line "
            "and call stack come from instrumentation applied to injected source, so a "
            "procedure already in the document does not carry them; run the same code through "
            "xlide_run_vba when you need them. A run that exceeds its timeout reports "
            "'timeout' and the application is terminated. Ask the user before running a macro "
            "that changes data."
        ),
    )
    def run_macro(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        procedure: Annotated[
            str,
            Field(description="Procedure to call, as 'Proc' or 'Module.Proc'."),
        ],
        args: Annotated[
            list[Any] | None,
            Field(default=None, description="Arguments, in order. Scalars only."),
        ] = None,
        timeout: Annotated[
            float,
            Field(default=0, ge=0, description="Seconds before the run is terminated."),
        ] = 0,
        read_only: Annotated[
            bool,
            Field(
                default=True,
                description="Open the document read-only. False lets the macro change it.",
            ),
        ] = True,
        save: Annotated[
            bool,
            Field(
                default=False,
                description="Save the document after the run. Needs read_only=false.",
            ),
        ] = False,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        if not read_only or save:
            require_writable(settings, "xlide_run_macro with read_only=false")
        if save and read_only:
            raise ToolError("save=true needs read_only=false; a read-only document cannot save.")
        deadline = clamp_timeout(timeout or None, settings)

        with _session(info) as session:
            _open_document(session, path, info, read_only)
            result = session.run_macro(procedure, *(args or []), timeout=deadline)
            answer = _run_result(result, deadline)
            if save and answer["outcome"] == "passed":
                session.save_as(str(path))
                answer["saved"] = True
        answer["path"] = str(path)
        answer["procedure"] = procedure
        return answer

    @server.tool(
        name="xlide_run_vba",
        title="Run VBA source",
        annotations=writes("Run VBA source", destructive=True, idempotent=False),
        description=(
            "Injects VBA source into a document and calls one procedure from it, under a "
            "deadline. Windows only. Use it to check what a piece of code actually does before "
            "writing it into a file, or to read something out of a document that no existing "
            "macro exposes. With no file_path it runs in a new empty document. Errors come "
            "back as data with the failing line and call stack, never as a dialog. Call "
            "PyVbaLog \"text\" from the code to return output; Debug.Print and MsgBox do not "
            "work under automation."
        ),
    )
    def run_vba(
        source: Annotated[str, Field(description="VBA source to inject.")],
        procedure: Annotated[
            str,
            Field(
                default="",
                description="Procedure to call. Empty calls the first one in the source.",
            ),
        ] = "",
        args: Annotated[
            list[Any] | None,
            Field(default=None, description="Arguments, in order. Scalars only."),
        ] = None,
        host: Annotated[
            str,
            Field(
                default="excel",
                description="'excel', 'word', 'powerpoint' or 'access'. Ignored with file_path.",
            ),
        ] = "excel",
        file_path: Annotated[
            str,
            Field(
                default="",
                description="Run against this document instead of a new empty one.",
            ),
        ] = "",
        timeout: Annotated[
            float, Field(default=0, ge=0, description="Seconds before the run is terminated.")
        ] = 0,
        read_only: Annotated[
            bool, Field(default=True, description="Open the document read-only.")
        ] = True,
    ) -> dict[str, Any]:
        if not source.strip():
            raise ToolError("No source given.")
        deadline = clamp_timeout(timeout or None, settings)

        info: HostInfo | None = None
        path: Path | None = None
        if file_path.strip():
            path = resolve_path(file_path, settings)
            # Injected source runs in the session, not out of the document, so a
            # file with no VBA project of its own is a perfectly good target.
            info = require_openable(path)
            if not read_only:
                require_writable(settings, "xlide_run_vba with read_only=false")
        else:
            info = _host_only(host)

        with _session(info) as session:
            if path is not None:
                _open_document(session, path, info, read_only)
            kwargs: dict[str, Any] = {"timeout": deadline}
            if procedure.strip():
                kwargs["proc"] = procedure.strip()
            if args:
                kwargs["args"] = tuple(args)
            result = session.run_vba(source, **kwargs)
            answer = _run_result(result, deadline)
        answer["host"] = info.host
        if path is not None:
            answer["path"] = str(path)
        return answer

    @server.tool(
        name="xlide_run_tests",
        title="Run VBA tests",
        annotations=writes("Run VBA tests", destructive=False, idempotent=False),
        description=(
            "Runs the VBA tests in an Office file: every zero-argument procedure whose name "
            "starts with Test. Windows only, with the application installed. Each test is run "
            "on its own, so one that hangs is reported as a timeout and the rest still run. "
            "Returns pass or fail per test with the assertion message, failing line, call "
            "stack and any logged output. Use PyVbaAssert and PyVbaAssertEqual in the tests; "
            "the harness supplies them."
        ),
    )
    def run_tests(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        module_name: Annotated[
            str,
            Field(
                default="",
                description="Run only this module's tests. Empty runs every module's.",
            ),
        ] = "",
        timeout: Annotated[
            float, Field(default=0, ge=0, description="Seconds allowed for each test.")
        ] = 0,
    ) -> dict[str, Any]:
        from .. import project as project_layer

        path = resolve_path(file_path, settings)
        info = require_readable(path)
        deadline = clamp_timeout(timeout or None, settings)

        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
        wanted = module_name.strip().casefold()
        candidates = [
            module
            for module in modules
            if (not wanted or module.name.casefold() == wanted) and _has_tests(module.body)
        ]
        if not candidates:
            scope = f"module {module_name}" if wanted else "this file"
            raise ToolError(
                f"No test procedures found in {scope}. A test is a Public Sub with no "
                "parameters whose name starts with Test. Write one, or name the module that "
                "holds them."
            )

        cases: list[dict[str, Any]] = []
        with _session(info) as session:
            _open_document(session, path, info, read_only=True)
            for module in candidates:
                for case in session.run_tests(module.body, timeout=deadline):
                    cases.append(_test_case(module.name, case))

        passed = sum(1 for case in cases if case["passed"])
        return {
            "path": str(path),
            "modules": [module.name for module in candidates],
            "total": len(cases),
            "passed": passed,
            "failed": len(cases) - passed,
            "verdict": "all passed" if passed == len(cases) else "failures",
            "cases": cases,
        }

    @server.tool(
        name="xlide_compile_check",
        title="Compile check",
        annotations=read_only("Compile check"),
        description=(
            "Asks the real VBA editor to compile a file's project, and reports whether it "
            "accepted it. Windows only. This is the compiler's own verdict, where xlide_analyze "
            "is a static analyzer's; run the analyzer first, because it is free and needs no "
            "Office, and use this when the last word matters. Excel is made visible for the "
            "check, because a hidden instance does not surface the compile-error dialog and "
            "would report a rejection as a pass."
        ),
    )
    def compile_check(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        timeout: Annotated[
            float, Field(default=0, ge=0, description="Seconds to wait for the verdict.")
        ] = 0,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        if info.host == "access":
            raise ToolError(
                "Access has no read-only automation mode, so a compile check there would "
                "write to the database, and a read-only check that writes is not one. "
                "xlide_analyze answers the same question without opening Access at all, and "
                "Access recompiles the project itself the next time it opens the file."
            )
        deadline = clamp_timeout(timeout or None, settings)
        with _session(info) as session:
            _open_document(session, path, info, read_only=True)
            result = session.compile_project(watch_seconds=min(deadline, 60))
        outcome = getattr(result, "outcome", "unknown")
        answer: dict[str, Any] = {
            "path": str(path),
            "outcome": outcome,
            "accepted": outcome == "accepted",
        }
        dialog = getattr(result, "dialog", None)
        if dialog is not None and getattr(dialog, "message", ""):
            answer["message"] = dialog.message
        if outcome == "infrastructure-failure":
            answer["note"] = (
                "The check could not complete, so the result is unknown. This is not a "
                "verdict on the code."
            )
        return answer


def office_report() -> dict[str, Any]:
    """What xlide_doctor says about running VBA on this machine."""
    if sys.platform != "win32":
        return {
            "available": False,
            "reason": (
                "Running VBA needs Windows with the desktop application installed. "
                "Every file and analysis tool still works here."
            ),
        }
    try:
        from pyvbaharness import apps
    except ImportError:
        return {
            "available": False,
            "reason": (
                "pyvbaharness is not installed. Ask the user before installing, then: "
                "pip install 'xlide-mcp[live]'"
            ),
        }
    found: dict[str, Any] = {}
    office_version = ""
    for key, detail in apps.APPS.items():
        current = _registry(_HKLM, rf"SOFTWARE\Classes\{detail.progid}\CurVer")
        entry: dict[str, Any] = {"installed": current is not None}
        if current is None:
            entry["note"] = f"{detail.label} is not installed on this machine."
        else:
            entry["progid"] = current
            if not detail.multi_instance:
                entry["note"] = "Single-instance: one session at a time, and no pool."
            if not detail.can_hide:
                entry["runs_on_screen"] = True
            tail = str(current).rsplit(".", 1)[-1]
            if not office_version and tail.isdigit():
                office_version = tail + ".0"
        found[key] = entry

    # Excel, Word and PowerPoint block the VBA project object model by default,
    # and module injection goes through it. The setting is per application and
    # per Windows user, so it is read here rather than assumed.
    for key, detail in apps.APPS.items():
        if not detail.needs_vbom or not found[key]["installed"]:
            continue
        if not office_version:
            found[key]["vba_project_access"] = "unknown (Office version not determined)"
            continue
        access = _registry(
            _HKCU,
            rf"Software\Microsoft\Office\{office_version}\{detail.security_key}\Security",
            "AccessVBOM",
        )
        found[key]["vba_project_access"] = "enabled" if access == 1 else "disabled"
        if access != 1:
            found[key]["remedy"] = (
                f"In {detail.label}: File > Options > Trust Center > Trust Center Settings > "
                "Macro Settings > tick 'Trust access to the VBA project object model'. "
                "Ask the user to do it; do not change the setting for them."
            )

    report: dict[str, Any] = {
        "available": any(entry["installed"] for entry in found.values()),
        "applications": found,
    }
    # Break on All Errors stops in the debugger even for handled errors, so every
    # managed run would come back modal-blocked instead of vba-error.
    for vba_version in ("7.1", "7.0", "6.0"):
        mode = _registry(_HKCU, rf"Software\Microsoft\VBA\{vba_version}\Common", "BreakOnAllErrors")
        if mode is None:
            continue
        if mode == 1:
            report["vbe_error_trapping"] = "break on all errors"
            report["warning"] = (
                "The VBA editor is set to Break on All Errors, so handled errors stop in the "
                "debugger and runs report modal-blocked rather than vba-error. Ask the user to "
                "set Tools > Options > General > Error Trapping to 'Break on Unhandled Errors'."
            )
        else:
            report["vbe_error_trapping"] = "break on unhandled errors"
        break
    return report


_HKLM = "HKEY_LOCAL_MACHINE"
_HKCU = "HKEY_CURRENT_USER"


def _registry(root: str, path: str, name: str | None = None) -> Any:
    """One registry value, or None. Windows only; absent is not an error."""
    if sys.platform != "win32":
        return None
    import winreg

    try:
        with winreg.OpenKey(getattr(winreg, root), path) as key:
            value, _kind = winreg.QueryValueEx(key, name)
            return value
    except OSError:
        return None


def _host_only(host: str) -> HostInfo:
    wanted = (host or "excel").strip().lower()
    if wanted not in _SESSION_CLASSES:
        raise ToolError(
            f"{host!r} is not a host. Use 'excel', 'word', 'powerpoint' or 'access'."
        )
    return HostInfo(host=wanted, extension="", readable=True)  # type: ignore[arg-type]


def _session(info: HostInfo) -> Any:
    """An owned application session, or a refusal naming exactly what is missing."""
    if sys.platform != "win32":
        raise needs_windows(f"Running VBA in {info.title}")
    try:
        import pyvbaharness
    except ImportError as exc:
        raise needs_package(f"Running VBA in {info.title}", "pyvbaharness", "live") from exc

    session_class = getattr(pyvbaharness, _SESSION_CLASSES[info.host])
    # Office automation is sequential by contract: one session per application
    # per machine. Two tool calls in a row are the ordinary case, though, and the
    # first application takes a couple of seconds to exit after its session
    # closes - measured, three grid calls in succession collided on the second.
    # Waiting is what the caller meant; failing because the previous call has not
    # finished tidying up is not.
    try:
        config = pyvbaharness.HarnessConfig(lock_wait_s=SESSION_LOCK_WAIT_SECONDS)
    except TypeError:  # pragma: no cover - an older harness without the setting
        config = None
    try:
        return session_class(config) if config is not None else session_class()
    except pyvbaharness.HarnessError as exc:
        raise ToolError(_harness_refusal(exc, info)) from exc


def _harness_refusal(exc: Exception, info: HostInfo) -> str:
    text = str(exc)
    lowered = text.lower()
    if "trust" in lowered and "project" in lowered:
        return (
            f"{info.title} blocks access to the VBA project object model, which this needs. "
            "Ask the user to enable it: File > Options > Trust Center > Trust Center Settings "
            "> Macro Settings > Trust access to the VBA project object model. It is per "
            f"application and per Windows user. ({text})"
        )
    if "already" in lowered or "running" in lowered:
        return (
            f"{info.title} is already running, and this server will not take ownership of a "
            "process it did not create, because a timeout terminates the process it owns. "
            f"Ask the user to close {info.title} and try again. ({text})"
        )
    return f"{info.title} could not be started: {text}"


def _open_document(session: Any, path: Path, info: HostInfo, read_only: bool) -> None:
    # Access has no read-only automation mode at all: running anything there
    # writes to the database as it goes. Refusing before the call names the
    # decision the user has to make, rather than passing through a message about
    # a parameter the caller never saw.
    if info.host == "access" and read_only:
        raise ToolError(
            f"Access cannot open {path.name} read-only for automation: running code there "
            "writes to the database as it goes, and injecting a module changes the file "
            "immediately rather than at a save. Pass read_only=false to allow that, and ask "
            "the user first, because it cannot be undone."
        )
    try:
        session.open_document(str(path), read_only=read_only)
    except Exception as exc:
        raise ToolError(f"{info.title} could not open {path.name}: {exc}") from exc


def _run_result(result: Any, deadline: float) -> dict[str, Any]:
    outcome = getattr(result, "outcome", "runner-error")
    answer: dict[str, Any] = {
        "outcome": outcome,
        "passed": outcome == "passed",
        "value": _plain(getattr(result, "value", None)),
        "output": list(getattr(result, "output", []) or []),
    }
    error = getattr(result, "error", None)
    if error is not None:
        answer["error"] = {
            "number": getattr(error, "number", None),
            "source": getattr(error, "source", ""),
            "description": getattr(error, "description", ""),
            "line": getattr(error, "line", None),
            "stack": [
                {"procedure": frame[0], "line": frame[1]}
                for frame in (getattr(error, "stack", []) or [])
            ],
        }
        if answer["error"]["line"] is None:
            # The failing line comes from line-number instrumentation, which is
            # applied to source as it is injected. A procedure that was already in
            # the document was never instrumented, so there is no line to report,
            # and saying why beats leaving an agent to wonder what it did wrong.
            answer["error"]["line_note"] = (
                "No failing line: the procedure was already in the document, and line "
                "numbers come from instrumenting source as it is injected. To get the line "
                "and the call stack, run the same code through xlide_run_vba."
            )
    if outcome == "timeout":
        answer["note"] = (
            f"The run passed its {deadline:g} second deadline and the application was "
            "terminated. Nothing is known about how far the code got."
        )
    elif outcome == "modal-blocked":
        answer["note"] = (
            "A dialog appeared that needs a person to answer it, so the run was stopped and "
            "the application terminated. Usually a MsgBox, or a prompt the code turned back "
            "on. Neither the harness nor this server answers a dialog with a real choice in it."
        )
    elif outcome == "runner-error":
        answer["note"] = (
            "The harness or a COM call failed around the run. This says nothing about whether "
            "the VBA is correct."
        )
    dialogs = getattr(result, "dialogs", None)
    if dialogs:
        answer["dialogs"] = [str(d) for d in dialogs]
    return answer


def _test_case(module: str, case: Any) -> dict[str, Any]:
    result = getattr(case, "result", None)
    entry: dict[str, Any] = {
        "module": module,
        "name": getattr(case, "name", "?"),
        "passed": bool(getattr(case, "passed", False)),
        "outcome": getattr(result, "outcome", "unknown"),
    }
    error = getattr(result, "error", None)
    if error is not None:
        entry["message"] = getattr(error, "description", "")
        entry["line"] = getattr(error, "line", None)
        entry["stack"] = [
            {"procedure": frame[0], "line": frame[1]}
            for frame in (getattr(error, "stack", []) or [])
        ]
    output = list(getattr(result, "output", []) or [])
    if output:
        entry["output"] = output
    return entry


def _has_tests(source: str) -> bool:
    import re

    return bool(
        re.search(
            r"^\s*(?:Public\s+|Private\s+|Friend\s+)?Sub\s+Test\w*\s*\(\s*\)",
            source,
            re.IGNORECASE | re.MULTILINE,
        )
    )


def _plain(value: Any) -> Any:
    """A VBA return value as JSON. A COM object is a marker, not a failure."""
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (list, tuple)):
        return [_plain(item) for item in value]
    return str(value)
