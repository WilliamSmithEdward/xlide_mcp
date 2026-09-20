"""Reading and changing VBA modules: the work almost every task comes down to.

Two things here are guards rather than features, and both exist because an agent's
read and its write are separated by a long gap. The content token makes a write
conditional on the module not having changed since the read. The document-module
rule refuses a rename or delete of ThisWorkbook, Sheet1 and their kind, because the
host recreates those and a project that has lost one no longer matches its file.
"""

from __future__ import annotations

import re
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import project as project_layer
from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import require_writable, resolve_path
from ..tokens import check_content_token, content_token
from ._common import change_summary, limited, read_only, truncate, writes

_VALID_NAME = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,30}$")

# VB reserved words a module cannot be called. Not the full keyword list: these are
# the ones a caller actually reaches for.
_RESERVED_NAMES = frozenset(
    {
        "and", "as", "boolean", "byref", "byte", "byval", "call", "case", "class", "const",
        "currency", "date", "declare", "dim", "do", "double", "each", "else", "elseif", "end",
        "enum", "error", "event", "exit", "false", "for", "function", "get", "global", "gosub",
        "goto", "if", "implements", "in", "integer", "is", "let", "like", "long", "loop", "me",
        "mod", "new", "next", "not", "nothing", "null", "object", "on", "option", "optional",
        "or", "private", "property", "public", "raiseevent", "redim", "rem", "resume", "return",
        "select", "set", "single", "static", "stop", "string", "sub", "then", "to", "true",
        "type", "until", "variant", "wend", "while", "with", "xor",
    }
)

_PROCEDURE_RE = re.compile(
    r"^\s*(?:(?P<scope>Public|Private|Friend|Global)\s+)?"
    r"(?P<static>Static\s+)?"
    r"(?P<kind>Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set)\s+"
    r"(?P<name>[A-Za-z_][A-Za-z0-9_]*)"
    r"\s*(?P<signature>\(.*)?$",
    re.IGNORECASE,
)


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_list_modules",
        title="List modules",
        annotations=read_only("List modules"),
        description=(
            "Lists the VBA modules in an Office file: name, kind (standard, class, document "
            "or userform), line count, and a content_token to pass to a guarded write. "
            "xlide_project_info returns this and more in one call; use this one when the file "
            "is already known and only the module list is wanted."
        ),
    )
    def list_modules(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
        return {
            "path": str(path),
            "host": info.host,
            "count": len(modules),
            "modules": [m.summary() for m in modules],
        }

    @server.tool(
        name="xlide_read_module",
        title="Read module",
        annotations=read_only("Read module"),
        description=(
            "The canonical way to read VBA. Returns a module's source as the VBA editor shows "
            "it, with the attribute header stripped, plus a content_token. Pass that token "
            "back as expected_content_token when you write, and the write is refused if "
            "anything changed the module in between. start_line and end_line read a slice of "
            "a long module; both are 1-based and inclusive."
        ),
    )
    def read_module(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        module_name: Annotated[str, Field(description="Module name, matched without case.")],
        start_line: Annotated[
            int, Field(default=0, ge=0, description="First line to return. 0 means the start.")
        ] = 0,
        end_line: Annotated[
            int, Field(default=0, ge=0, description="Last line to return. 0 means the end.")
        ] = 0,
        include_header: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "Include the Attribute VB_* header. Off by default: the header is managed "
                    "for you on write, and editing it by hand is how a module loses its "
                    "binding to a sheet or a form."
                ),
            ),
        ] = False,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
        module = project_layer.find_module(modules, module_name)

        source = module.full_source if include_header else module.body
        lines = source.splitlines()
        first = max(1, start_line or 1)
        last = min(len(lines), end_line or len(lines))
        if first > len(lines):
            raise ToolError(
                f"{module.name} has {len(lines)} lines; start_line {start_line} is past the end."
            )
        sliced = "\n".join(lines[first - 1 : last])
        text, was_cut = truncate(sliced, hint="Read it in slices with start_line and end_line.")
        result: dict[str, Any] = {
            "path": str(path),
            "module": module.name,
            "kind": module.kind,
            "content_token": module.token,
            "total_lines": len(lines),
            "first_line": first,
            "last_line": last,
            "truncated": was_cut,
            "source": text,
        }
        if first != 1 or last != len(lines):
            result["note"] = (
                "This is a slice. The content_token describes the whole module, so a write "
                "using it must send the whole module."
            )
        return result

    @server.tool(
        name="xlide_list_procedures",
        title="List procedures",
        annotations=read_only("List procedures"),
        description=(
            "Lists the Sub, Function and Property procedures in one module, with each one's "
            "kind, scope, line number and signature. Use it to find where to change something "
            "without reading a long module in full."
        ),
    )
    def list_procedures(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        module_name: Annotated[str, Field(description="Module name, matched without case.")],
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
        module = project_layer.find_module(modules, module_name)
        return {
            "path": str(path),
            "module": module.name,
            "kind": module.kind,
            "procedures": _procedures(module.body),
        }

    @server.tool(
        name="xlide_search_modules",
        title="Search VBA",
        annotations=read_only("Search VBA"),
        description=(
            "Searches every module's source in an Office file and returns each match with its "
            "module, line number and the line itself. Use it to find where a name is declared "
            "or used before changing it. Plain text by default; set is_regex for a Python "
            "regular expression."
        ),
    )
    def search_modules(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        query: Annotated[str, Field(description="Text or pattern to find.")],
        is_regex: Annotated[
            bool, Field(default=False, description="Treat query as a regular expression.")
        ] = False,
        match_case: Annotated[
            bool,
            Field(
                default=False,
                description="Match case. Off by default, because VBA itself ignores case.",
            ),
        ] = False,
        max_results: Annotated[
            int, Field(default=200, ge=1, le=2000, description="Stop after this many matches.")
        ] = 200,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        if not query.strip():
            raise ToolError("No query given.")
        flags = 0 if match_case else re.IGNORECASE
        try:
            pattern = re.compile(query if is_regex else re.escape(query), flags)
        except re.error as exc:
            raise ToolError(f"Not a valid regular expression: {exc}.") from exc

        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)

        matches: list[dict[str, Any]] = []
        truncated = False
        for module in modules:
            for number, line in enumerate(module.body.splitlines(), start=1):
                if not pattern.search(line):
                    continue
                if len(matches) >= max_results:
                    truncated = True
                    break
                matches.append(
                    {"module": module.name, "line": number, "text": line.strip()[:300]}
                )
            if truncated:
                break
        return {
            "path": str(path),
            "query": query,
            "match_count": len(matches),
            "truncated": truncated,
            "matches": matches,
        }

    @server.tool(
        name="xlide_write_module",
        title="Write module",
        annotations=writes("Write module", destructive=True),
        description=(
            "The canonical way to change VBA. Writes a module's source into the Office file "
            "and saves it. It replaces the module's whole source, so send all of it; send the "
            "body only, because the attribute header is managed for you. A module that does "
            "not exist is created: standard by default, or a class with kind='class'. Pass "
            "expected_content_token from your read and the write is refused if the module "
            "changed since. After writing, call xlide_analyze and treat any error as a build "
            "failure. Ask the user first when the project is protected or signed."
        ),
    )
    def write_module(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        module_name: Annotated[str, Field(description="Module to write, or to create.")],
        source: Annotated[str, Field(description="The module's complete VBA source.")],
        expected_content_token: Annotated[
            str,
            Field(
                default="",
                description=(
                    "The content_token from your read. Leave empty only when creating a module."
                ),
            ),
        ] = "",
        kind: Annotated[
            str,
            Field(
                default="standard",
                description="Kind for a module being created: 'standard' or 'class'.",
            ),
        ] = "standard",
        allow_protected: Annotated[
            bool,
            Field(
                default=False,
                description="Write to a password-protected project. Ask the user first.",
            ),
        ] = False,
        allow_invalidate_signature: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "Write to a digitally signed project, dropping the signature. "
                    "Ask the user first."
                ),
            ),
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_write_module")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        wanted_kind = kind.strip().lower()
        if wanted_kind not in {"standard", "class"}:
            raise ToolError("kind must be 'standard' or 'class'.")

        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
            existing = next(
                (m for m in modules if m.name.casefold() == module_name.strip().casefold()), None
            )

            if existing is None:
                _check_new_name(module_name, modules)
                if expected_content_token:
                    raise ToolError(
                        f"No module named {module_name!r} yet, but expected_content_token was "
                        "supplied. Either the name is wrong, or the module was deleted after "
                        "your read. Call xlide_list_modules and decide which."
                    )
                created = True
                before = ""
            else:
                stale = check_content_token(existing.body, expected_content_token, existing.name)
                if stale is not None:
                    raise ToolError(stale.message)
                created = False
                before = existing.body

            if created:
                import pyopenvba

                module_kind = (
                    pyopenvba.VBAModuleKind.standard
                    if wanted_kind == "standard"
                    else pyopenvba.VBAModuleKind.other
                )
                handle.vba_project().add_module(module_name, source, kind=module_kind)
                written_name = module_name
            else:
                handle.set_module(existing.name, source)
                written_name = existing.name

            save_warnings = project_layer.save(
                handle,
                info,
                allow_protected=allow_protected,
                allow_invalidate_signature=allow_invalidate_signature,
            )

        # Read back rather than trust the write. The project re-derives the body
        # from what it stored, and that is the text the next read will return.
        with project_layer.open_project(path, info) as handle:
            after = project_layer.find_module(
                project_layer.read_modules(handle, info), written_name
            )

        result: dict[str, Any] = {
            "path": str(path),
            "module": after.name,
            "kind": after.kind,
            "created": created,
            "content_token": after.token,
            "saved": True,
            **change_summary(before, after.body),
        }
        if save_warnings:
            result["warnings"] = save_warnings
        result["next_step"] = (
            "Call xlide_analyze on this file and fix anything at error severity."
        )
        return result

    @server.tool(
        name="xlide_rename_module",
        title="Rename module",
        annotations=writes("Rename module", destructive=True),
        description=(
            "Renames a VBA module everywhere its name is stored, and saves the file. Calls to "
            "the module's procedures elsewhere in the project are not rewritten; search for "
            "the old name first with xlide_search_modules. Document modules such as "
            "ThisWorkbook and Sheet1 cannot be renamed, because the host owns them."
        ),
    )
    def rename_module(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        module_name: Annotated[str, Field(description="Module to rename.")],
        new_name: Annotated[str, Field(description="New name. A valid VBA identifier.")],
        allow_protected: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
        allow_invalidate_signature: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_rename_module")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
            module = project_layer.find_module(modules, module_name)
            if project_layer.is_document_module(module.kind):
                raise ToolError(
                    f"{module.name} is a document module: {info.title} owns it and recreates "
                    "it by name. Renaming it would leave the project not matching the file. "
                    "Its code can still be written with xlide_write_module."
                )
            _refuse_form_module(module, info, "Renaming")
            orphaned = _shapes_calling(path, info, module.name)
            _check_new_name(new_name, modules)
            # The project's own rename works on every host, Access included:
            # AccessDatabase.rename_module addresses its design modules, and
            # refuses an ordinary one by name.
            handle.vba_project().rename_module(module.name, new_name)
            save_warnings = project_layer.save(
                handle,
                info,
                allow_protected=allow_protected,
                allow_invalidate_signature=allow_invalidate_signature,
            )
        result: dict[str, Any] = {
            "path": str(path),
            "renamed_from": module.name,
            "renamed_to": new_name,
            "saved": True,
            "next_step": (
                f"Calls to {module.name} elsewhere still name it. Search for it with "
                "xlide_search_modules, then run xlide_analyze."
            ),
        }
        if save_warnings:
            result["warnings"] = save_warnings
        if orphaned:
            result["shapes_still_naming_the_old_module"] = orphaned
            result["warning"] = (
                f"{len(orphaned)} shapes name a procedure in {module.name}, and nothing "
                "rewrites an OnAction. Repoint them with xlide_set_shape_macro."
            )
        return result

    @server.tool(
        name="xlide_delete_module",
        title="Delete module",
        annotations=writes("Delete module", destructive=True, idempotent=False),
        description=(
            "Permanently deletes a VBA module from an Office file and saves it. There is no "
            "undo: ask the user first, and read the module before deleting it so its code can "
            "be put back if they change their mind. Document modules cannot be deleted."
        ),
    )
    def delete_module(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        module_name: Annotated[str, Field(description="Module to delete.")],
        expected_content_token: Annotated[
            str,
            Field(
                default="",
                description=(
                    "The content_token from your read, so a module that changed since is not "
                    "deleted on the strength of a stale look at it."
                ),
            ),
        ] = "",
        allow_protected: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
        allow_invalidate_signature: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_delete_module")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
            module = project_layer.find_module(modules, module_name)
            if project_layer.is_document_module(module.kind):
                raise ToolError(
                    f"{module.name} is a document module and cannot be deleted; "
                    f"{info.title} owns it. To empty it, write an empty body with "
                    "xlide_write_module."
                )
            _refuse_form_module(module, info, "Deleting")
            orphaned = _shapes_calling(path, info, module.name)
            stale = check_content_token(module.body, expected_content_token, module.name)
            if stale is not None:
                raise ToolError(stale.message)
            handle.vba_project().delete_module(module.name)
            save_warnings = project_layer.save(
                handle,
                info,
                allow_protected=allow_protected,
                allow_invalidate_signature=allow_invalidate_signature,
            )
        result: dict[str, Any] = {
            "path": str(path),
            "deleted": module.name,
            "lines_deleted": module.line_count,
            "saved": True,
            "next_step": "Run xlide_analyze: anything that called into it no longer resolves.",
        }
        if save_warnings:
            result["warnings"] = save_warnings
        if orphaned:
            result["shapes_now_calling_nothing"] = orphaned
            result["warning"] = (
                f"{len(orphaned)} shapes still name a procedure in {module.name}, and nothing "
                "rewrites an OnAction. Point them elsewhere with xlide_set_shape_macro, or "
                "tell the user which buttons have stopped working."
            )
        return result


def _shapes_calling(path: Any, info: Any, module_name: str) -> list[dict[str, Any]]:
    """Shapes whose OnAction names this module.

    Deleting a module a button calls is a legitimate thing to do, so this warns
    rather than refusing. What it prevents is the silent version: nothing rewrites
    an OnAction, so the button keeps naming a procedure that has gone, and the
    user finds out by clicking it.
    """
    if info.host != "excel" or not info.supports_sheets:
        return []
    try:
        from ..shapes import read_sheet_shapes

        by_sheet = read_sheet_shapes(path)
    except Exception:
        return []
    wanted = module_name.casefold()
    found: list[dict[str, Any]] = []
    for sheet, shapes in by_sheet.items():
        for shape in shapes:
            if not shape.macro:
                continue
            owner = shape.macro.split(".", 1)[0] if "." in shape.macro else ""
            if owner.casefold() == wanted:
                found.append({"sheet": sheet, "shape": shape.name, "macro": shape.macro})
    return found


def _refuse_form_module(
    module: project_layer.ModuleView, info: Any, operation: str
) -> None:
    """A form's code cannot be renamed or deleted on its own.

    A form is two things that have to agree: a designer storage and a module of
    the same name. Renaming only the module leaves a storage with no module,
    which the host does not show, and a module with no storage, which is a class.
    Deleting only the module leaves the storage behind entirely. Both produce a
    file where the form has silently gone, and the module tools have no way to
    move the storage with it.
    """
    if module.kind != "userform":
        return
    if info.host == "access":
        design = module.name.split("_", 1)[-1]
        raise ToolError(
            f"{module.name} is the code behind the Access design {design!r}. "
            f"{operation} it on its own would separate the code from the design. "
            f"Use xlide_manage_form, which moves both together."
        )
    raise ToolError(
        f"{module.name} is a UserForm's code, and its design is stored beside it. "
        f"{operation} the module alone would leave the design orphaned and the form "
        "gone from the editor, and this server has no way to move the design with it. "
        f"Ask the user to do it in the {info.title} editor. Its code can still be "
        "written with xlide_write_module."
    )


def _check_new_name(name: str, existing: list[project_layer.ModuleView]) -> None:
    candidate = (name or "").strip()
    if not _VALID_NAME.match(candidate):
        raise ToolError(
            f"{name!r} is not a valid module name. A name starts with a letter, continues "
            "with letters, digits or underscores, and is at most 31 characters."
        )
    if candidate.casefold() in _RESERVED_NAMES:
        raise ToolError(f"{candidate!r} is a VBA reserved word and cannot name a module.")
    clash = next((m for m in existing if m.name.casefold() == candidate.casefold()), None)
    if clash is not None:
        raise ToolError(
            f"A module named {clash.name!r} already exists. VBA compares names without case, "
            "so the two would collide."
        )


def _procedures(source: str) -> list[dict[str, Any]]:
    """Every procedure declaration in a module body, with its line.

    Continued signatures are joined first: a declaration split over three lines
    with trailing underscores is one declaration, and reporting three is worse
    than reporting none.
    """
    found: list[dict[str, Any]] = []
    lines = source.splitlines()
    index = 0
    while index < len(lines):
        start = index
        statement = lines[index].rstrip()
        while statement.endswith("_") and index + 1 < len(lines):
            index += 1
            statement = statement[:-1].rstrip() + " " + lines[index].strip()
        index += 1
        match = _PROCEDURE_RE.match(statement)
        if not match:
            continue
        kind = re.sub(r"\s+", " ", match.group("kind")).title()
        signature = (match.group("signature") or "").strip()
        found.append(
            {
                "name": match.group("name"),
                "kind": kind,
                "scope": (match.group("scope") or "Public").title(),
                "line": start + 1,
                "signature": f"{kind} {match.group('name')}{signature}".strip(),
            }
        )
    shown, total = limited(found, 1000)
    if total > len(shown):
        shown.append({"note": f"{total - len(shown)} more procedures not listed."})
    return shown


__all__ = ["content_token", "register"]
