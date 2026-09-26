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

References can be written as well as read. Adding one means being right about a
GUID and a version, and a wrong one is a reference the host marks MISSING with a
message naming nothing, so the GUID comes from this machine's registry rather
than from anyone's memory of it: the same list the VBA editor's References dialog
shows. pyOpenVBA writes the record, marks the compiled cache stale so the host
reads the change, and refuses the two edits that break a project: the host's own
library, which is implicit, and Microsoft Forms while a UserForm needs it.
"""

from __future__ import annotations

from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import project as project_layer
from .. import typelibs, xlide_vscode
from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import require_writable, resolve_path
from ..tokens import content_token
from ._common import (
    ALLOW_PROTECTED_DESCRIPTION,
    ALLOW_SIGNATURE_DESCRIPTION,
    page,
    read_only,
    writes,
)

# The four libraries pyOpenVBA can add by name alone, on any platform.
_OFFICE_LIBRARIES = frozenset({"excel", "word", "powerpoint", "access"})


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
                f"not there. {info.title}'s own library and VBA's are implicit and never listed."
            ),
        }
        if not has_project:
            result["note"] = project_layer.no_project_note(path, info)
        return result

    @server.tool(
        name="xlide_manage_reference",
        title="Add or remove a project reference",
        annotations=writes("Add or remove a project reference", destructive=True),
        description=(
            "Adds or removes a type library reference in a VBA project and saves the file. "
            "Use it when xlide_analyze reports missing-library-reference, or when code needs "
            "early binding, as Dim d As Scripting.Dictionary does. Name the library: Excel, "
            "Word, PowerPoint and Access work anywhere, and on Windows any library the VBA "
            "editor's References dialog lists resolves by its description (Microsoft "
            "Scripting Runtime) or by the name code writes (Scripting), taking its GUID, "
            "version and path from this machine's registry. A name that matches several "
            "libraries is refused with the candidates. Off Windows, pass guid and version. "
            "Adding the host's own library, or removing Microsoft Forms while a UserForm "
            "exists, is refused: the first is implicit and the second stops the project "
            "compiling."
        ),
    )
    def manage_reference(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        action: Annotated[str, Field(description="'add' or 'remove'.")],
        library: Annotated[
            str,
            Field(
                description=(
                    "For add: the library's description, the name code uses, or Excel, Word, "
                    "PowerPoint or Access. For remove: the name xlide_list_references reports, "
                    "or the GUID."
                )
            ),
        ],
        guid: Annotated[
            str,
            Field(
                default="",
                description=(
                    "For add: the library's GUID, when it is not registered on this machine. "
                    "library is then the name the reference is recorded under."
                ),
            ),
        ] = "",
        version: Annotated[
            str,
            Field(
                default="",
                description=(
                    "For add with guid: major.minor in hexadecimal, as the registry spells it, "
                    "such as 1.0 or 2.8. Empty takes the newest registered, or 1.0."
                ),
            ),
        ] = "",
        allow_protected: Annotated[
            bool, Field(default=False, description=ALLOW_PROTECTED_DESCRIPTION)
        ] = False,
        allow_invalidate_signature: Annotated[
            bool, Field(default=False, description=ALLOW_SIGNATURE_DESCRIPTION)
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_manage_reference")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        wanted = (action or "").strip().lower()
        if wanted not in {"add", "remove"}:
            raise ToolError("action must be 'add' or 'remove'.")
        if info.host == "vb6":
            raise ToolError(
                "A Visual Basic 6 project's references are Reference= lines in its .vbp. Edit "
                "the manifest, or add the reference in the VB6 IDE."
            )
        if not (library or "").strip() and not (guid or "").strip():
            raise ToolError("Name the library, or give its guid.")

        import pyopenvba

        with project_layer.open_project(path, info) as handle:
            if not project_layer.has_project(handle, info):
                raise ToolError(
                    project_layer.no_project_note(path, info)
                    + " Write the first module, then add the reference."
                )
            before = [_reference(ref) for ref in project_layer.references(handle, info)]
            try:
                if wanted == "add":
                    spec = _library_spec(library, guid, version)
                    handle.add_reference(
                        spec["name"],
                        spec.get("guid"),
                        spec.get("major", 1),
                        spec.get("minor", 0),
                        path=spec.get("path", ""),
                        description=spec.get("description", ""),
                    )
                else:
                    target = (library or guid).strip()
                    if not handle.remove_reference(target):
                        listed = ", ".join(entry["name"] for entry in before) or "(none)"
                        raise ToolError(
                            f"The project has no reference named {target!r}. It references: "
                            f"{listed}. {info.title}'s own library and VBA's are implicit and "
                            "cannot be removed."
                        )
            except pyopenvba.PyOpenVBAError as exc:
                # The base class: Access refuses with its own AccessError.
                raise ToolError(_reference_refusal(str(exc), info)) from exc
            after = [_reference(ref) for ref in project_layer.references(handle, info)]
            if after == before:
                return {
                    "path": str(path),
                    "action": wanted,
                    "changed": False,
                    "saved": False,
                    "references": after,
                    "note": "The project already references that library. Nothing was written.",
                }
            save_warnings = project_layer.save(
                handle,
                info,
                path=path,
                allow_protected=allow_protected,
                allow_invalidate_signature=allow_invalidate_signature,
            )

        known = {entry["name"].casefold() for entry in before}
        result: dict[str, Any] = {
            "path": str(path),
            "action": wanted,
            "changed": True,
            "saved": True,
            "references": after,
        }
        if wanted == "add":
            result["added"] = next(
                (entry for entry in after if entry["name"].casefold() not in known), None
            )
            result["next_step"] = (
                "Run xlide_analyze: names qualified with this library now resolve."
            )
        else:
            remaining = {entry["name"].casefold() for entry in after}
            result["removed"] = [
                entry["name"] for entry in before if entry["name"].casefold() not in remaining
            ]
            result["next_step"] = (
                "Run xlide_analyze: code that named this library no longer compiles."
            )
        if save_warnings:
            result["warnings"] = save_warnings
        notice = xlide_vscode.file_changed(path, "references", tool="xlide_manage_reference")
        if notice:
            result["xlide_vscode"] = notice
        return result

    @server.tool(
        name="xlide_access_catalog",
        title="Access tables and queries",
        annotations=read_only("Access tables and queries"),
        description=(
            "Lists what an Access database holds besides its code: tables with their columns, "
            "saved queries with their SQL, and the relationships between tables. An .accdb is "
            "an application rather than a document, and the VBA in it is written against "
            "these, so reading the modules alone shows half of it. Use include with offset "
            "and next_offsets to page through a large catalog. Long query SQL is previewed; "
            "use xlide_read_access_query for the whole text. If one query's SQL cannot be "
            "read, that entry has sql=null and sql_error rather than hiding the catalog. "
            "Access files only."
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
        offset: Annotated[
            int, Field(default=0, ge=0, description="Items to skip in each list.")
        ] = 0,
        max_results: Annotated[
            int, Field(default=300, ge=1, le=500, description="Most items per list.")
        ] = 300,
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

        result: dict[str, Any] = {
            "path": str(path), "offset": offset, "counts": {}, "next_offsets": {},
        }
        with project_layer.open_project(path, info) as handle:
            if wanted in {"tables", "all"}:
                shown, count, next_offset = _tables(
                    handle, include_system, offset, max_results
                )
                result["tables"] = shown
                result["counts"]["tables"] = count
                result["next_offsets"]["tables"] = next_offset
            if wanted in {"queries", "all"}:
                shown, count, next_offset = _queries(handle, offset, max_results)
                result["queries"] = shown
                result["counts"]["queries"] = count
                result["next_offsets"]["queries"] = next_offset
            if wanted in {"relationships", "all"}:
                found = _relationships(handle, include_system)
                result["relationships"], result["next_offsets"]["relationships"] = page(
                    found, "relationships", offset, max_results
                )
                result["counts"]["relationships"] = len(found)
        return result

    @server.tool(
        name="xlide_read_access_query",
        title="Read Access saved query",
        annotations=read_only("Read Access saved query"),
        description=(
            "Reads the SQL of one saved query in an Access database. Use it after "
            "xlide_access_catalog when a query's SQL was previewed. offset and max_chars "
            "read long SQL in character slices; next_offset points to the next slice. "
            "content_token describes the whole SQL, so compare it across pages if the "
            "database may have changed. Access files only."
        ),
    )
    def read_access_query(
        file_path: Annotated[str, Field(description="Absolute path to the Access database.")],
        query_name: Annotated[str, Field(description="Saved query name, matched without case.")],
        offset: Annotated[int, Field(default=0, ge=0, description="Characters to skip.")] = 0,
        max_chars: Annotated[
            int,
            Field(default=20_000, ge=1, le=40_000, description="Most SQL characters to return."),
        ] = 20_000,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        if info.host != "access":
            raise ToolError(
                f"{path.name} is a {info.title} file. Saved queries require an "
                "Access database."
            )
        with project_layer.open_project(path, info) as handle:
            try:
                saved = list(handle.queries())
            except Exception as exc:
                raise ToolError(f"The database's saved queries could not be read: {exc}") from exc
            requested = query_name.strip().casefold()
            query = next((entry for entry in saved if entry.name.casefold() == requested), None)
            if query is None:
                names = ", ".join(entry.name for entry in saved[:10]) or "(none)"
                more = len(saved) - 10
                suffix = (
                    f" and {more} more; use xlide_access_catalog to list them."
                    if more > 0 else "."
                )
                raise ToolError(
                    f"No saved query named {query_name!r}. This database has: {names}{suffix}"
                )
            try:
                sql = getattr(query, "sql", "") or ""
            except Exception as exc:
                raise ToolError(f"The SQL for {query.name} could not be read: {exc}") from exc
        if offset > len(sql):
            raise ToolError(
                f"{query.name} has {len(sql)} SQL characters; offset {offset} is past the end."
            )
        next_offset = offset + min(max_chars, len(sql) - offset)
        return {
            "path": str(path),
            "query": query.name,
            "sql": sql[offset:next_offset],
            "content_token": content_token(sql),
            "total_chars": len(sql),
            "offset": offset,
            "next_offset": next_offset if next_offset < len(sql) else None,
        }


def _library_spec(library: str, guid: str, version: str) -> dict[str, Any]:
    """What to add: a name, and a GUID, version, path and description where needed.

    Three routes, in order of how much can go wrong. An explicit GUID is taken as
    given, filled in from the registry where it is registered. One of the four
    Office libraries is pyOpenVBA's to spell. Anything else is looked up in this
    machine's registry by description or by the library's own name, because a
    GUID from memory is how a project gets a reference Excel marks MISSING.
    """
    name = (library or "").strip()
    if (guid or "").strip():
        normal = typelibs.normal_guid(guid)
        if normal is None:
            raise ToolError(f"{guid!r} is not a GUID. It looks like {{420B2830-E718-11CF-...}}.")
        spec: dict[str, Any] = {"guid": normal}
        registered = typelibs.search(normal, limit=1)
        if version.strip():
            parsed = typelibs.parse_version(version)
            if parsed is None:
                raise ToolError(
                    f"{version!r} is not a version. Give major.minor in hexadecimal, such as 1.0."
                )
            spec["major"], spec["minor"] = parsed
        elif registered:
            spec["major"], spec["minor"] = registered[0].major, registered[0].minor
        if registered:
            spec["path"] = registered[0].path
            spec["description"] = registered[0].description
            name = name or typelibs.library_name(registered[0])
        if not name:
            raise ToolError(
                "Give library as the name code uses to qualify it, such as Scripting: the "
                "library is not registered here, so its own name cannot be read."
            )
        spec["name"] = name
        return spec

    if name.casefold() in _OFFICE_LIBRARIES:
        return {"name": name}

    if not typelibs.available():
        raise ToolError(
            f"Looking {name!r} up needs Windows, where type libraries are registered. Pass its "
            "guid and version, with library as the name code uses to qualify it."
        )
    found = typelibs.resolve(name)
    if isinstance(found, list):
        if not found:
            raise ToolError(
                f"No type library registered on this machine matches {name!r}. Check the "
                "name in the VBA editor's Tools > References, or pass its guid and version."
            )
        listed = "; ".join(
            f"{lib.description} ({lib.guid} {lib.version})" for lib in found[:10]
        )
        more = f", and {len(found) - 10} more" if len(found) > 10 else ""
        raise ToolError(
            f"{name!r} matches {len(found)} registered libraries: {listed}{more}. Pass the "
            "exact description, or its guid."
        )
    own_name = typelibs.library_name(found)
    if not own_name:
        raise ToolError(
            f"{found.description} is registered as {found.guid} {found.version}, but the name "
            "code uses for it could not be read, which needs pywin32. Call again with that "
            "guid and library set to the name, such as Scripting."
        )
    return {
        "name": own_name,
        "guid": found.guid,
        "major": found.major,
        "minor": found.minor,
        "path": found.path,
        "description": found.description,
    }


def _reference_refusal(text: str, info: Any) -> str:
    """pyOpenVBA's refusal, said in terms of what the agent should do next."""
    lowered = text.casefold()
    if "implicit project library" in lowered:
        return (
            f"{info.title}'s own library and VBA's are part of every {info.title} project "
            "without a reference, which is why the References dialog shows them ticked and "
            f"greyed. Nothing to add. ({text})"
        )
    if "microsoft forms is required" in lowered:
        return (
            f"{text}. Without Microsoft Forms nothing in a project with a UserForm compiles, "
            "so the reference stays while a form does."
        )
    if "different guid" in lowered:
        return f"{text}. Remove the existing reference first, or pick another library."
    return f"The reference was refused: {text}"


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


def _tables(
    handle: Any, include_system: bool, offset: int, max_results: int
) -> tuple[list[dict[str, Any]], int, int | None]:
    try:
        names = list(handle.table_names(include_system=include_system))
    except Exception as exc:
        raise ToolError(f"The database's tables could not be read: {exc}") from exc

    out: list[dict[str, Any]] = []
    for name in names[offset : offset + max_results]:
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
    shown, _ = page(out, "tables", 0, max_results)
    following = offset + len(shown)
    return shown, len(names), following if following < len(names) else None


def _column(column: Any) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "name": getattr(column, "name", ""),
        "type": str(getattr(column, "type", "")),
    }
    size = getattr(column, "size", None)
    if size:
        entry["size"] = size
    return entry


def _queries(
    handle: Any, offset: int, max_results: int
) -> tuple[list[dict[str, Any]], int, int | None]:
    try:
        saved = list(handle.queries())
    except Exception as exc:
        raise ToolError(f"The database's saved queries could not be read: {exc}") from exc
    out: list[dict[str, Any]] = []
    for query in saved[offset : offset + max_results]:
        entry: dict[str, Any] = {"name": getattr(query, "name", "")}
        try:
            sql = getattr(query, "sql", "") or ""
        except Exception as exc:
            entry.update({
                "sql": None,
                "sql_chars": None,
                "sql_truncated": None,
                "sql_error": f"The SQL could not be read: {exc}",
            })
        else:
            entry.update({
                "sql": sql[:8_000],
                "sql_chars": len(sql),
                "sql_truncated": len(sql) > 8_000,
            })
        out.append(entry)
    shown, _ = page(out, "queries", 0, max_results)
    following = offset + len(shown)
    return shown, len(saved), following if following < len(saved) else None


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
