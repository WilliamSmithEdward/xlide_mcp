"""What a project points at, and what an Access database holds besides code.

Two questions the module tools cannot answer.

The first is which type libraries a project references. It decides whether
`CreateObject("Scripting.Dictionary")` is late binding or a missing reference
waiting to fail, what `New Recordset` resolves to, and why code that works on one
machine does not work on another. A reference is a path plus a version, so a
project can be broken by a machine rather than by its own code, and the path is
the only thing that says so.

The second is the rest of an Access database. An .accdb is an application, not a
document: its tables, its saved queries and the relationships between them are
what the VBA is written against, and reading the modules alone shows half of it.

Both read and neither writes. Adding a reference means being right about a GUID
and a version on a machine this server cannot see, and a wrong one produces a
project that fails to compile with a message naming nothing.
"""

from __future__ import annotations

from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import project as project_layer
from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import resolve_path
from ._common import limited, read_only


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_list_references",
        title="List project references",
        annotations=read_only("List project references"),
        description=(
            "Lists the type libraries a VBA project references: the name code uses to qualify "
            "them, the kind, and the registry or file path the project recorded. Call it when "
            "a name will not resolve, when deciding between early and late binding, or when "
            "code works on one machine and not another: a reference names a path and a "
            "version, so a project can be broken by the machine rather than by its own code."
        ),
    )
    def list_references(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            # Access keeps its references in the database and the other three keep
            # theirs in the dir stream, but both record the same libid string, so
            # one reader serves.
            entries = [_reference(ref) for ref in project_layer.references(handle, info)]
            has_project = project_layer.has_project(handle, info)
        result: dict[str, Any] = {
            "path": str(path),
            "host": info.host,
            "count": len(entries),
            "references": entries,
            "note": (
                "A reference is recorded as a path and a version. A file missing from that "
                "path on another machine is the usual cause of code that compiles here and "
                "not there."
            ),
        }
        if not has_project:
            result["note"] = project_layer.no_project_note(path, info)
        return result

    @server.tool(
        name="xlide_access_catalog",
        title="Access tables and queries",
        annotations=read_only("Access tables and queries"),
        description=(
            "Lists what an Access database holds besides its code: tables with their columns, "
            "saved queries with their SQL, and the relationships between tables. An .accdb is "
            "an application rather than a document, and the VBA in it is written against "
            "these, so reading the modules alone shows half of it. Access files only."
        ),
    )
    def access_catalog(
        file_path: Annotated[str, Field(description="Absolute path to the Access database.")],
        include: Annotated[
            str,
            Field(
                default="all",
                description="'tables', 'queries', 'relationships' or 'all'.",
            ),
        ] = "all",
        include_system: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "Include the MSys tables and relationships Access keeps its own "
                    "objects in. Off by default: every database has them, and listing "
                    "them buries the few that are about the user's data."
                ),
            ),
        ] = False,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        if info.host != "access":
            raise ToolError(
                f"{path.name} is a {info.title} file. Tables and saved queries exist in "
                "Access databases only."
            )
        wanted = (include or "all").strip().lower()
        if wanted not in {"tables", "queries", "relationships", "all"}:
            raise ToolError(
                "include must be 'tables', 'queries', 'relationships' or 'all'."
            )

        result: dict[str, Any] = {"path": str(path)}
        with project_layer.open_project(path, info) as handle:
            if wanted in {"tables", "all"}:
                result["tables"] = _tables(handle, include_system)
            if wanted in {"queries", "all"}:
                result["queries"] = _queries(handle)
            if wanted in {"relationships", "all"}:
                result["relationships"] = _relationships(handle, include_system)
        return result


def _reference(reference: Any) -> dict[str, Any]:
    name = getattr(reference, "name_unicode", "") or getattr(reference, "name", "")
    entry: dict[str, Any] = {"name": name}
    kind = getattr(reference, "kind", "")
    if kind:
        entry["kind"] = kind
    libid = getattr(reference, "libid", "")
    if libid:
        entry.update(_describe_libid(libid))
        entry["libid"] = libid
    return entry


def _describe_libid(libid: str) -> dict[str, Any]:
    """Pull the readable parts out of a libid.

    The format is `*\\G{guid}#major.minor#lcid#path#description`. It is split
    rather than parsed strictly: a libid that does not match is left as the one
    string it arrived as, because a half-parsed path is worse than none.
    """
    parts = libid.split("#")
    if len(parts) < 5:
        return {}
    out: dict[str, Any] = {}
    guid = parts[0].removeprefix("*\\G").removeprefix("*\\H")
    if guid.startswith("{"):
        out["guid"] = guid
    if parts[1]:
        out["version"] = parts[1]
    if parts[3]:
        out["path"] = parts[3]
    if parts[4]:
        out["description"] = parts[4]
    return out


def _tables(handle: Any, include_system: bool) -> list[dict[str, Any]]:
    try:
        names = list(handle.table_names(include_system=include_system))
    except Exception as exc:
        raise ToolError(f"The database's tables could not be read: {exc}") from exc

    out: list[dict[str, Any]] = []
    for name in names:
        entry: dict[str, Any] = {"name": name}
        try:
            columns, indexes = handle.table_specs(name)
        except Exception:
            entry["note"] = "The column definitions could not be read."
            out.append(entry)
            continue
        entry["columns"] = [_column(column) for column in columns]
        if indexes:
            entry["indexes"] = [
                {
                    "name": getattr(index, "name", ""),
                    "columns": list(getattr(index, "columns", ()) or ()),
                }
                for index in indexes
            ]
        out.append(entry)
    shown, total = limited(out, 500)
    if total > len(shown):
        shown.append({"note": f"{total - len(shown)} further tables not listed."})
    return shown


def _column(column: Any) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "name": getattr(column, "name", ""),
        "type": str(getattr(column, "type", "")),
    }
    size = getattr(column, "size", None)
    if size:
        entry["size"] = size
    return entry


def _queries(handle: Any) -> list[dict[str, Any]]:
    try:
        saved = list(handle.queries())
    except Exception as exc:
        raise ToolError(f"The database's saved queries could not be read: {exc}") from exc
    out = [
        {
            "name": getattr(query, "name", ""),
            "sql": (getattr(query, "sql", "") or "").strip(),
        }
        for query in saved
    ]
    shown, total = limited(out, 500)
    if total > len(shown):
        shown.append({"note": f"{total - len(shown)} further queries not listed."})
    return shown


def _relationships(handle: Any, include_system: bool) -> list[dict[str, Any]]:
    try:
        related = list(handle.relationships())
    except Exception as exc:
        raise ToolError(f"The database's relationships could not be read: {exc}") from exc
    out: list[dict[str, Any]] = []
    for relationship in related:
        table = getattr(relationship, "table", "")
        referenced = getattr(relationship, "referenced_table", "")
        # Every database carries the navigation-pane relationships Access keeps
        # its own objects in. Listing them by default buries the two or three
        # that are about the user's data.
        if not include_system and (
            table.startswith("MSys") or referenced.startswith("MSys")
        ):
            continue
        out.append(
            {
                "name": getattr(relationship, "name", ""),
                "from_table": table,
                "from_columns": list(getattr(relationship, "columns", ()) or ()),
                "to_table": referenced,
                "to_columns": list(getattr(relationship, "referenced_columns", ()) or ()),
            }
        )
    return out
