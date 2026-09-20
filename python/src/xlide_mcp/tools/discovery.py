"""Finding a file, and learning everything about it in one call.

`xlide_project_info` is deliberately one round trip. An agent that has to make six
calls to find out what a workbook holds spends most of a turn on orientation, and
the six answers are all cheap to produce together: the VBA project is opened once
and the package is read once.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import project as project_layer
from ..config import Settings
from ..errors import ToolError
from ..hosts import CREATABLE, NOT_READABLE, host_info, iter_office_files, require_readable
from ..paths import require_writable, resolve_path
from ._common import limited, read_only, writes


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_list_projects",
        title="List Office files",
        annotations=read_only("List Office files"),
        description=(
            "Finds the Office files in this server's workspace and returns their absolute "
            "paths, host application and whether their VBA can be opened. Call this first "
            "when the user has not named a file. Covers Excel (.xlsm, .xlsb, .xlam, .xls, "
            "and .xlsx for Power Query and sheets), Word (.docm, .dotm, .doc), PowerPoint "
            "(.pptm, .potm), Access (.accdb, .mdb), and Visual Basic 6 projects (.vbp, "
            "whose modules are the files its manifest names). Files whose extension is "
            "recognized but not openable are listed with the reason, so a template or add-in "
            "is not silently missing."
        ),
    )
    def list_projects(
        subfolder: Annotated[
            str,
            Field(
                default="",
                description=(
                    "Limit the search to this folder. Empty searches every workspace root."
                ),
            ),
        ] = "",
    ) -> dict[str, Any]:
        roots = (
            [resolve_path(subfolder, settings, must_exist=True, must_be_file=False)]
            if subfolder.strip()
            else list(settings.effective_roots())
        )
        found: list[dict[str, Any]] = []
        for root in roots:
            if not root.is_dir():
                continue
            for path in iter_office_files(root, include_unreadable=True):
                info = host_info(path)
                entry: dict[str, Any] = {
                    "path": str(path),
                    "name": path.name,
                    "host": info.host,
                    "readable": info.readable,
                }
                if not info.readable:
                    entry["reason"] = info.reason
                found.append(entry)
        shown, total = limited(found)
        result: dict[str, Any] = {
            "roots": [str(r) for r in roots],
            "count": total,
            "files": shown,
        }
        if total > len(shown):
            result["note"] = f"{total - len(shown)} more; narrow the search with subfolder."
        if not found:
            result["note"] = (
                "No Office files under "
                + ", ".join(str(r) for r in roots)
                + ". Ask the user for the path, or for the folder the server should watch."
            )
        return result

    @server.tool(
        name="xlide_project_info",
        title="Project summary",
        annotations=read_only("Project summary"),
        description=(
            "Everything about one Office file in a single call: its VBA modules with kinds "
            "and line counts, its UserForms, its Power Query queries, its worksheets with "
            "used ranges and named ranges, and whether the VBA project is password-protected "
            "or digitally signed. Call this once per file before working on it. Each module "
            "carries a content_token for a guarded write."
        ),
    )
    def project_info(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = host_info(path)
        result: dict[str, Any] = {
            "path": str(path),
            "name": path.name,
            "host": info.host,
            "host_application": info.title,
            "vba_readable": info.readable,
        }
        if not info.readable:
            result["reason"] = info.reason
        else:
            with project_layer.open_project(path, info) as handle:
                modules = project_layer.read_modules(handle, info)
                status = project_layer.project_status(handle, info)
                result["project_name"] = status.project_name
                result["modules"] = [m.summary() for m in modules]
                result["password_protected"] = status.password_protected
                result["digitally_signed"] = status.digitally_signed
                result["forms"] = _form_names(handle)
            cautions = status.warnings()
            if cautions:
                result["cautions"] = cautions

        if info.supports_power_query:
            result["power_query"] = _query_summary(path)
        if info.supports_sheets:
            result.update(_sheet_summary(path))
        elif info.host == "excel":
            result["sheets_note"] = (
                f"Worksheet cells are not read from {info.extension}; its grid is not OOXML. "
                "The VBA project is fully readable."
            )
        return result

    @server.tool(
        name="xlide_validate_project",
        title="Validate VBA project",
        annotations=read_only("Validate VBA project"),
        description=(
            "Checks a file's VBA project for structural problems: records that disagree with "
            "each other, a module the directory names but the container does not hold, and "
            "the like. This is about the container, not about the code; use xlide_analyze for "
            "the code. Worth calling before risky work on an old or repaired file."
        ),
    )
    def validate_project(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            try:
                problems = list(handle.validate())
            except AttributeError:
                return {
                    "path": str(path),
                    "supported": False,
                    "note": f"Structural validation is not implemented for {info.title} files.",
                }
        return {
            "path": str(path),
            "supported": True,
            "problem_count": len(problems),
            "problems": problems,
            "verdict": "clean" if not problems else "problems found",
        }

    @server.tool(
        name="xlide_create_project",
        title="Create Office file",
        annotations=writes("Create Office file", destructive=False, idempotent=False),
        description=(
            "Creates a new Office file with an empty VBA project at the given absolute path, "
            "from a template the application itself authored, so it opens with no repair "
            "prompt. Extensions: .xlsm, .xlsb, .xlam, .docm, .pptm, .accdb, and .xlsx for a "
            "workbook with Power Query and no macros. It never overwrites an existing file."
        ),
    )
    def create_project(
        file_path: Annotated[
            str, Field(description="Absolute path to create. The extension picks the format.")
        ],
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_create_project")
        path = resolve_path(file_path, settings, must_exist=False)
        if path.exists():
            raise ToolError(
                f"{path} already exists. This tool never overwrites a file; "
                "pick another path, or work on the existing one."
            )
        extension = path.suffix.lower()
        if extension not in CREATABLE:
            hint = NOT_READABLE.get(extension, "")
            raise ToolError(
                f"Cannot create {extension or 'a file with no extension'}. "
                f"Creatable: {', '.join(sorted(CREATABLE))}." + (f" {hint}" if hint else "")
            )
        if not path.parent.is_dir():
            raise ToolError(f"The folder does not exist: {path.parent}.")

        if extension == ".xlsx":
            import pyopenvba

            with pyopenvba.PowerQueryWorkbook.create_new(path) as book:
                book.save()
            return {
                "path": str(path),
                "created": True,
                "host": "excel",
                "note": "A .xlsx has no VBA project. It holds Power Query and sheets.",
            }

        info = host_info(path)
        from ..hosts import container_class

        handle = container_class(info).create_new(path)
        try:
            handle.save()
        finally:
            handle.close()
        return {
            "path": str(path),
            "created": True,
            "host": info.host,
            "note": (
                "Empty VBA project. Add code with xlide_write_module, which creates a module "
                "that does not exist yet."
            ),
        }

    @server.tool(
        name="xlide_doctor",
        title="What this machine can do",
        annotations=read_only("What this machine can do"),
        description=(
            "Reports what this server can reach: its workspace roots, whether it is read-only, "
            "which optional layers are installed, and, on Windows, which Office applications "
            "are available to run macros and tests. Call it when a tool reports something "
            "missing, or before promising the user a test run."
        ),
    )
    def doctor() -> dict[str, Any]:
        import platform
        import sys

        report: dict[str, Any] = {
            "server_version": _version(),
            "python": sys.version.split()[0],
            "platform": platform.platform(),
            "workspace_roots": [str(r) for r in settings.effective_roots()],
            "read_only": settings.read_only,
            "paths_outside_roots_allowed": settings.allow_outside_roots,
            "layers": {
                "files": _probe("pyopenvba"),
                "analysis": _probe("pyvbaanalysis"),
                "execution": _probe("pyvbaharness"),
            },
        }
        from .execution import office_report
        from .live import live_report

        report["execution"] = office_report()
        report["live_sessions"] = live_report()
        return report


def _version() -> str:
    from ..server import __version__

    return __version__


def _probe(module: str) -> dict[str, Any]:
    try:
        imported = __import__(module)
    except ImportError as exc:
        return {"available": False, "reason": str(exc)}
    return {"available": True, "version": getattr(imported, "__version__", "unknown")}


def _form_names(handle: Any) -> list[str]:
    try:
        return [form.name for form in handle.forms()]
    except Exception:
        return []


def _query_summary(path: Path) -> dict[str, Any]:
    import pyopenvba

    try:
        with pyopenvba.PowerQueryWorkbook(path) as book:
            names = list(book.query_names())
    except pyopenvba.PowerQueryError:
        return {"count": 0, "queries": []}
    except Exception as exc:
        return {"count": 0, "queries": [], "note": str(exc)}
    return {"count": len(names), "queries": names}


def _sheet_summary(path: Path) -> dict[str, Any]:
    from ..xlsx import Workbook, XlsxError

    try:
        book = Workbook(path)
        sheets = [
            {"name": s.name, "used_range": s.used_range or "(empty)", "hidden": s.hidden}
            for s in book.sheets()
        ]
        named = [{"name": n.name, "refers_to": n.refers_to} for n in book.named_ranges()]
    except XlsxError as exc:
        return {"sheets_note": str(exc)}
    out: dict[str, Any] = {"sheets": sheets}
    if named:
        out["named_ranges"] = named
    return out
