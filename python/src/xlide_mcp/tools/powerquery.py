"""Power Query: the M code in a workbook, which is code the same way VBA is.

Excel keeps Get and Transform queries in one custom XML part, base64-encoded,
outside the VBA project. That is why a plain .xlsx with no macros still has
queries in it, and why an agent asked to explain what a workbook does has to look
here as well as at the modules.

Loading a query onto a sheet is deliberately not exposed as a write. It takes a
connection, a query table, a table and the sheet's reference to it, and the column
names have to be given because knowing them means running the query. Getting that
wrong produces a workbook that opens and then fails on refresh, which is a worse
outcome than not offering it.
"""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from ..config import Settings
from ..errors import ToolError
from ..hosts import host_info
from ..paths import require_writable, resolve_path
from ._common import bound, read_only, truncate, writes


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_list_queries",
        title="List Power Query",
        annotations=read_only("List Power Query"),
        description=(
            "Lists the Power Query queries in an Excel workbook: name, group, description, "
            "where each loads, and its applied step names. Queries live outside the VBA "
            "project, so a plain .xlsx has them too. Call this when asked what a workbook "
            "does; a workbook with no macros can still be doing most of its work here."
        ),
    )
    def list_queries(
        file_path: Annotated[str, Field(description="Absolute path to the Excel workbook.")],
    ) -> dict[str, Any]:
        path = _query_path(file_path, settings)
        with _open(path) as book:
            queries = [
                {
                    "name": query.name,
                    "group": _group_name(query),
                    "description": query.description or "",
                    "load_target": str(query.load_target or ""),
                    "steps": _steps(query),
                    "is_function": bool(query.is_function),
                }
                for query in book.queries()
            ]
            groups = [group.name for group in book.groups()]
        shown, note = bound(queries, "queries", "Read one with xlide_read_query.")
        result: dict[str, Any] = {
            "path": str(path),
            "count": len(queries),
            "queries": shown,
            "groups": groups,
        }
        if note:
            result["note"] = note
        return result

    @server.tool(
        name="xlide_read_query",
        title="Read Power Query",
        annotations=read_only("Read Power Query"),
        description=(
            "Reads one query's M formula, with its description, group, load target and refresh "
            "settings. The formula is the whole let ... in expression as the Advanced Editor "
            "shows it."
        ),
    )
    def read_query(
        file_path: Annotated[str, Field(description="Absolute path to the Excel workbook.")],
        query_name: Annotated[str, Field(description="Query name, matched without case.")],
    ) -> dict[str, Any]:
        path = _query_path(file_path, settings)
        import pyopenvba

        with _open(path) as book:
            query = _find_query(book, query_name)
            formula, was_cut = truncate(query.formula or "")
            result: dict[str, Any] = {
                "path": str(path),
                "query": query.name,
                "group": _group_name(query),
                "description": getattr(query, "description", "") or "",
                "load_target": str(getattr(query, "load_target", "") or ""),
                "steps": _steps(query),
                "truncated": was_cut,
                "formula": formula,
            }
            # `refresh` is a property that raises for a query loading nowhere,
            # so getattr with a default does not shield the call: a query with no
            # sheet behind it has no connection, and therefore no refresh settings.
            try:
                refresh = query.refresh
            except pyopenvba.PowerQueryError:
                refresh = None
                result["refresh_note"] = (
                    "This query loads nowhere, so it has no connection and no refresh "
                    "settings. Those exist once a query is loaded onto a sheet."
                )
            if refresh is not None:
                result["refresh"] = {
                    "on_open": getattr(refresh, "on_open", None),
                    "interval_minutes": getattr(refresh, "interval_minutes", None),
                    "background": getattr(refresh, "background", None),
                    "keep_data": getattr(refresh, "keep_data", None),
                    "in_refresh_all": getattr(refresh, "in_refresh_all", None),
                    "enabled": getattr(refresh, "enabled", None),
                }
        return result

    @server.tool(
        name="xlide_write_query",
        title="Write Power Query",
        annotations=writes("Write Power Query", destructive=True),
        description=(
            "Changes a workbook's Power Query and saves it. action='set' replaces a query's M "
            "formula, creating it if it does not exist; 'rename' renames it and rewrites the "
            "queries that reference it by name; 'remove' deletes it, which has no undo, so ask "
            "the user first. A query already loaded onto a sheet keeps its loaded rows until "
            "Excel refreshes it."
        ),
    )
    def write_query(
        file_path: Annotated[str, Field(description="Absolute path to the Excel workbook.")],
        action: Annotated[str, Field(description="'set', 'rename' or 'remove'.")],
        query_name: Annotated[str, Field(description="The query to change.")],
        formula: Annotated[
            str,
            Field(
                default="",
                description="For set: the whole M expression, such as 'let Source = 1 in Source'.",
            ),
        ] = "",
        new_name: Annotated[
            str, Field(default="", description="For rename: the new name.")
        ] = "",
        description: Annotated[
            str, Field(default="", description="For set: the query's description.")
        ] = "",
        group: Annotated[
            str,
            Field(
                default="",
                description="For set on a new query: the folder in the Queries pane.",
            ),
        ] = "",
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_write_query")
        path = _query_path(file_path, settings)
        wanted = (action or "").strip().lower()
        if wanted not in {"set", "rename", "remove"}:
            raise ToolError("action must be 'set', 'rename' or 'remove'.")

        import pyopenvba

        with _open(path) as book:
            existing = {name.casefold() for name in book.query_names()}
            detail: dict[str, Any] = {}
            try:
                if wanted == "set":
                    if not formula.strip():
                        raise ToolError("formula is required for action='set'.")
                    if query_name.casefold() in existing:
                        query = _find_query(book, query_name)
                        query.formula = formula
                        if description:
                            query.description = description
                        detail = {"query": query.name, "created": False}
                    else:
                        target_group = _resolve_group(book, group)
                        query = book.add_query(query_name, formula, group=target_group)
                        if description:
                            query.description = description
                        detail = {"query": query_name, "created": True}
                elif wanted == "rename":
                    if not new_name.strip():
                        raise ToolError("new_name is required for action='rename'.")
                    book.rename_query(query_name, new_name.strip())
                    detail = {"renamed_from": query_name, "renamed_to": new_name.strip()}
                else:
                    # A query loaded onto a sheet is four things: the definition,
                    # a connection, a query table and the table itself. Removing
                    # only the definition leaves a connection pointing at a query
                    # that no longer exists, which Excel meets on the next
                    # refresh rather than on open, so nothing says so at the time.
                    unloaded = bool(book.unload(query_name))
                    book.remove_query(query_name)
                    detail = {"removed": query_name, "unloaded_from_sheet": unloaded}
                book.save()
            except ToolError:
                raise
            except pyopenvba.PowerQueryError as exc:
                raise ToolError(f"Power Query refused the change: {exc}") from exc
            except PermissionError as exc:
                raise ToolError(
                    f"{path.name} is locked, most likely open in Excel: {exc}. "
                    "Ask the user to close it."
                ) from exc

        return {
            "path": str(path),
            "action": wanted,
            "saved": True,
            **detail,
            "note": (
                "The query definition is stored. Its loaded rows, if it loads to a sheet, "
                "change when Excel next refreshes it."
            ),
        }


def _query_path(raw: str, settings: Settings) -> Path:
    path = resolve_path(raw, settings)
    info = host_info(path)
    if not info.supports_power_query:
        raise ToolError(
            f"Power Query lives in Excel packages. {path.name} is {info.title}"
            f"{' (' + info.extension + ')' if info.extension else ''}. "
            "Supported: .xlsx, .xlsm, .xlsb and .xlam."
        )
    return path


def _open(path: Path) -> Any:
    import pyopenvba

    try:
        return pyopenvba.PowerQueryWorkbook(path)
    except pyopenvba.PyOpenVBAError as exc:
        raise ToolError(f"{path.name}: {exc}") from exc


def _find_query(book: Any, name: str) -> Any:
    wanted = (name or "").strip().casefold()
    for query in book.queries():
        if query.name.casefold() == wanted:
            return query
    listed = ", ".join(book.query_names()) or "(none)"
    raise ToolError(f"No query named {name!r}. Queries in this workbook: {listed}.")


def _steps(query: Any) -> list[str]:
    """A query's applied step names. A formula this cannot parse yields none."""
    try:
        return list(query.steps or [])
    except Exception:
        return []


def _group_name(query: Any) -> str:
    group = getattr(query, "group", None)
    if group is None:
        return ""
    return getattr(group, "name", str(group))


def _resolve_group(book: Any, name: str) -> Any:
    """An existing group of that name, or a new one. Empty means no group."""
    if not name.strip():
        return None
    for group in book.groups():
        if getattr(group, "name", "").casefold() == name.strip().casefold():
            return group
    return book.add_group(name.strip())
