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
from pydantic import BaseModel, Field

from .. import project as project_layer
from .. import xlide_vscode
from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import require_writable, resolve_path
from ..tokens import check_content_token, content_token
from ._common import (
    ALLOW_PROTECTED_DESCRIPTION,
    ALLOW_SIGNATURE_DESCRIPTION,
    MAX_RESULT_CHARS,
    change_summary,
    page,
    read_only,
    truncate,
    unified_diff,
    writes,
)


class ModuleEdit(BaseModel):
    start_line: int = Field(ge=1, description="First original line to replace, 1-based.")
    end_line: int = Field(
        ge=0, description="Last original line, inclusive. Use start_line - 1 to insert."
    )
    replacement: str = Field(description="Complete replacement text for these lines.")

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
            "is already known and only the module list is wanted. Use offset and next_offset "
            "to read large projects in pages."
        ),
    )
    def list_modules(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        offset: Annotated[
            int, Field(default=0, ge=0, description="Skip this many modules before a page.")
        ] = 0,
        max_results: Annotated[
            int, Field(default=300, ge=1, le=300, description="Return at most this many modules.")
        ] = 300,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
            has_project = project_layer.has_project(handle, info)
        shown, next_offset = page(
            [m.summary() for m in modules], "modules", offset, max_results
        )
        result: dict[str, Any] = {
            "path": str(path),
            "host": info.host,
            "has_vba_project": has_project,
            "count": len(modules),
            "modules": shown,
            "offset": offset,
            "next_offset": next_offset,
        }
        if next_offset is not None:
            result["note"] = (
                f"{len(modules)} modules in all; call again with offset={next_offset} "
                "for the next page."
            )
        if not has_project:
            result["note"] = project_layer.no_project_note(path, info)
        return result

    @server.tool(
        name="xlide_read_module",
        title="Read module",
        annotations=read_only("Read module"),
        description=(
            "The canonical way to read VBA. Returns a module's source as the VBA editor shows "
            "it, with the attribute header stripped, plus a content_token. Pass that token "
            "back as expected_content_token when you write, and the write is refused if "
            "anything changed the module in between. start_line and end_line read a slice of "
            "a long module; both are 1-based and inclusive. line_ranges reads several "
            "slices in one call, with each pair [first, last] using body line numbers."
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
        line_ranges: Annotated[
            list[list[int]] | None,
            Field(
                default=None,
                description="Optional list of [first, last] inclusive 1-based line ranges.",
            ),
        ] = None,
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
            module = project_layer.find_module_in(handle, info, path, module_name)

        source = module.full_source if include_header else module.body
        lines = source.splitlines()
        if end_line and end_line < (start_line or 1):
            raise ToolError(
                f"end_line {end_line} is before start_line {start_line or 1}. "
                "Use an end line at or after the start line."
            )
        if line_ranges is not None:
            if start_line or end_line:
                raise ToolError("Use line_ranges or start_line/end_line, not both.")
            if include_header:
                raise ToolError("line_ranges uses body line numbers; set include_header=false.")
            if not line_ranges or len(line_ranges) > 20:
                raise ToolError("line_ranges needs 1 to 20 [first, last] pairs.")
            sections: list[dict[str, Any]] = []
            budget = max(1, MAX_RESULT_CHARS // len(line_ranges))
            for pair in line_ranges:
                if len(pair) != 2 or pair[0] < 1 or pair[1] < pair[0] or pair[1] > len(lines):
                    raise ToolError(
                        f"Invalid line range {pair!r}; {module.name} has {len(lines)} lines."
                    )
                part, cut = truncate(
                    "\n".join(lines[pair[0] - 1 : pair[1]]), budget,
                    hint="Read a narrower line range.",
                )
                sections.append({"first_line": pair[0], "last_line": pair[1],
                                 "source": part, "truncated": cut})
            return {
                "path": str(path), "module": module.name, "kind": module.kind,
                "content_token": module.token, "total_lines": len(lines),
                "sections": sections,
                "note": "The token describes the whole module; an edit can name these lines.",
            }
        if not lines:
            if start_line or end_line:
                raise ToolError(f"{module.name} has no body lines to read.")
            return {
                "path": str(path), "module": module.name, "kind": module.kind,
                "content_token": module.token, "total_lines": 0,
                "first_line": 0, "last_line": 0, "truncated": False, "source": "",
            }
        first = max(1, start_line or 1)
        last = min(len(lines), end_line or len(lines))
        if first > len(lines):
            raise ToolError(
                f"{module.name} has {len(lines)} lines; start_line {start_line} is past the end."
            )
        sliced = (
            source.replace("\r\n", "\n")
            if first == 1 and last == len(lines)
            else "\n".join(lines[first - 1 : last])
        )
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
                "This is a slice. The content_token describes the whole module. Use "
                "xlide_edit_module to change these lines, or send the whole module to "
                "xlide_write_module."
            )
        return result

    @server.tool(
        name="xlide_list_procedures",
        title="List procedures",
        annotations=read_only("List procedures"),
        description=(
            "Lists the Sub, Function and Property procedures in one module, with each one's "
            "kind, scope, line number and signature. The module's content_token can guard "
            "an xlide_edit_module call. Use it to find where to change something without "
            "reading a long module in full. Use offset and next_offset to page through a "
            "module with many procedures."
        ),
    )
    def list_procedures(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        module_name: Annotated[str, Field(description="Module name, matched without case.")],
        offset: Annotated[
            int, Field(default=0, ge=0, description="Procedures to skip.")
        ] = 0,
        max_results: Annotated[
            int, Field(default=300, ge=1, le=1000, description="Most procedures to return.")
        ] = 300,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            module = project_layer.find_module_in(handle, info, path, module_name)
        procedures = _procedures(module.body)
        shown, next_offset = page(procedures, "procedures", offset, max_results)
        return {
            "path": str(path),
            "module": module.name,
            "kind": module.kind,
            "content_token": module.token,
            "count": len(procedures),
            "offset": offset,
            "next_offset": next_offset,
            "procedures": shown,
        }

    @server.tool(
        name="xlide_search_modules",
        title="Search VBA",
        annotations=read_only("Search VBA"),
        description=(
            "Searches every module's source in an Office file and returns each match with its "
            "module, line number and the line itself. Use it to find where a name is declared "
            "or used before changing it. content_tokens maps returned module names to tokens "
            "for guarded xlide_edit_module calls without a separate read. Use offset and "
            "next_offset to page through broad searches; total_match_count counts all matches. "
            "Plain text by default; set is_regex for a Python regular expression."
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
            int, Field(default=200, ge=1, le=2000, description="Return at most this many matches.")
        ] = 200,
        offset: Annotated[
            int,
            Field(default=0, ge=0, description="Skip this many matches before a page."),
        ] = 0,
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
            has_project = project_layer.has_project(handle, info)

        matches: list[dict[str, Any]] = []
        content_tokens: dict[str, str] = {}
        total_match_count = 0
        for module in modules:
            for number, line in enumerate(module.body.splitlines(), start=1):
                if not pattern.search(line):
                    continue
                if offset <= total_match_count < offset + max_results:
                    matches.append(
                        {"module": module.name, "line": number, "text": line.strip()[:300]}
                    )
                    content_tokens[module.name] = module.token
                total_match_count += 1
        next_offset = offset + len(matches)
        has_more = next_offset < total_match_count
        result: dict[str, Any] = {
            "path": str(path),
            "query": query,
            "match_count": len(matches),
            "total_match_count": total_match_count,
            "offset": offset,
            "next_offset": next_offset if has_more else None,
            "truncated": has_more,
            "matches": matches,
            "content_tokens": content_tokens,
        }
        if not has_project:
            result["note"] = project_layer.no_project_note(path, info)
        return result

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
            "changed since. The result carries a diff of what the file now holds, read back "
            "after saving, which is what to show the user when they ask what changed. After "
            "writing, call xlide_analyze and treat any error as a build failure. Identical "
            "source is reported without saving. Ask the user first when the project is "
            "protected or signed."
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
        include_diff: Annotated[
            bool,
            Field(
                default=True,
                description=(
                    "Include a unified diff of what changed. Off in a loop that writes many "
                    "modules and reads none of them back."
                ),
            ),
        ] = True,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_write_module")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        wanted_kind = kind.strip().lower()
        if wanted_kind not in {"standard", "class"}:
            raise ToolError("kind must be 'standard' or 'class'.")

        with project_layer.open_project(path, info) as handle:
            # A file saved before its first macro has no project to add a module
            # to. It gets the one its application makes, which in Excel already
            # holds ThisWorkbook and a module per sheet, so the lookup below runs
            # against what the file holds after that, not before.
            project_created = project_layer.ensure_project(handle, info, path)
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
                if content_token(source) == existing.token:
                    unchanged: dict[str, Any] = {
                        "path": str(path), "module": existing.name, "kind": existing.kind,
                        "created": False, "changed": False, "saved": False,
                        "content_token": existing.token,
                        **change_summary(existing.body, existing.body),
                        "note": "The module already has this source. Nothing was written.",
                    }
                    if include_diff:
                        unchanged["diff"] = "(no change)"
                    return unchanged
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
                path=path,
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
            "changed": True,
            "content_token": after.token,
            "saved": True,
            **change_summary(before, after.body),
        }
        if include_diff:
            # Against the read-back, not against what was sent. The project
            # re-derives a body from what it stored, so this is the one diff that
            # shows what the file actually holds rather than what was intended.
            diff, was_cut = unified_diff(
                before,
                after.body,
                label=after.name,
                narrower="Read the module with xlide_read_module.",
            )
            result["diff"] = diff
            if was_cut:
                result["diff_truncated"] = True
        if save_warnings:
            result["warnings"] = save_warnings
        if project_created:
            result["vba_project_created"] = True
            result["note"] = (
                f"{path.name} had no VBA project, so it now has the one {info.title} makes "
                "for a first macro, with this module in it."
            )
        notice = xlide_vscode.module_written(
            path,
            after.name,
            before=before,
            before_existed=not created,
            after=after.body,
            kind=after.kind,
            tool="xlide_write_module",
        )
        if notice:
            result["xlide_vscode"] = notice
        result["next_step"] = (
            "Call xlide_analyze on this file and fix anything at error severity."
        )
        return result

    @server.tool(
        name="xlide_edit_module",
        title="Edit module lines",
        annotations=writes("Edit module lines", destructive=True),
        description=(
            "Replaces, inserts or deletes several line ranges in one existing VBA module and "
            "saves once. Lines are 1-based against the original source returned by "
            "xlide_read_module, without the attribute header. Each edit has start_line, "
            "end_line inclusive, and replacement text. For insertion, set end_line to "
            "start_line - 1; an empty replacement deletes the named lines. Ranges must not "
            "overlap. expected_content_token is required and refuses a stale edit. Set "
            "preview_only to validate the edits and see the proposed diff without saving. "
            "An applied edit includes the read-back diff and new token. Analyze the file afterward."
        ),
    )
    def edit_module(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        module_name: Annotated[str, Field(description="Existing module to edit.")],
        edits: Annotated[
            list[ModuleEdit],
            Field(description="Line edits against the original module body, applied together."),
        ],
        expected_content_token: Annotated[
            str, Field(description="Content token from the last read; required.")
        ],
        preview_only: Annotated[
            bool, Field(default=False, description="Show the proposed diff without saving.")
        ] = False,
        allow_protected: Annotated[
            bool, Field(default=False, description=ALLOW_PROTECTED_DESCRIPTION)
        ] = False,
        allow_invalidate_signature: Annotated[
            bool, Field(default=False, description=ALLOW_SIGNATURE_DESCRIPTION)
        ] = False,
    ) -> dict[str, Any]:
        if not preview_only:
            require_writable(settings, "xlide_edit_module")
        if not expected_content_token:
            raise ToolError("expected_content_token is required. Read the module first.")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            module = project_layer.find_module_in(handle, info, path, module_name)
        stale = check_content_token(module.body, expected_content_token, module.name)
        if stale is not None:
            raise ToolError(stale.message)
        source = _apply_line_edits(module.body, edits)
        if preview_only:
            diff, truncated = unified_diff(
                module.body, source, label=module.name,
                narrower="Read or edit a narrower line range to see the rest.",
            )
            preview: dict[str, Any] = {
                "path": str(path), "module": module.name, "kind": module.kind,
                "applied": False, "saved": False, "edits_applied": len(edits),
                "content_token": module.token,
                **change_summary(module.body, source),
                "diff": diff,
                "next_step": (
                    "To apply these edits, call xlide_edit_module again with "
                    "preview_only=false and the same expected_content_token."
                ),
            }
            if truncated:
                preview["diff_truncated"] = True
            return preview
        result = write_module(
            file_path, module.name, source,
            expected_content_token=expected_content_token,
            allow_protected=allow_protected,
            allow_invalidate_signature=allow_invalidate_signature,
        )
        result["applied"] = result["saved"]
        result["edits_applied"] = len(edits)
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
            bool, Field(default=False, description=ALLOW_PROTECTED_DESCRIPTION)
        ] = False,
        allow_invalidate_signature: Annotated[
            bool, Field(default=False, description=ALLOW_SIGNATURE_DESCRIPTION)
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_rename_module")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
            module = project_layer.find_module_in(handle, info, path, module_name, modules)
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
                path=path,
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
        notice = xlide_vscode.module_renamed(
            path, module.name, new_name, tool="xlide_rename_module"
        )
        if notice:
            result["xlide_vscode"] = notice
        return result

    @server.tool(
        name="xlide_delete_module",
        title="Delete module",
        annotations=writes("Delete module", destructive=True, idempotent=False),
        description=(
            "Permanently deletes a VBA module from an Office file and saves it. There is no "
            "undo: ask the user first, and read the module before deleting it so its code can "
            "be put back if they change their mind. Document modules cannot be deleted. "
            "For Excel, the write is refused if the server cannot check whether a shape "
            "still calls the module."
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
            bool, Field(default=False, description=ALLOW_PROTECTED_DESCRIPTION)
        ] = False,
        allow_invalidate_signature: Annotated[
            bool, Field(default=False, description=ALLOW_SIGNATURE_DESCRIPTION)
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_delete_module")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            module = project_layer.find_module_in(handle, info, path, module_name)
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
                path=path,
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
        # The deleted text goes with the notice, so XLIDE can offer to put it back.
        notice = xlide_vscode.module_written(
            path,
            module.name,
            before=module.body,
            before_existed=True,
            after="",
            after_exists=False,
            kind=module.kind,
            tool="xlide_delete_module",
        )
        if notice:
            result["xlide_vscode"] = notice
        return result


def _apply_line_edits(source: str, edits: list[ModuleEdit]) -> str:
    if not edits or len(edits) > 100:
        raise ToolError("edits needs 1 to 100 line edits.")
    lines = source.splitlines(keepends=True)
    ordered = sorted(edits, key=lambda edit: (edit.start_line, edit.end_line))
    previous: ModuleEdit | None = None
    for edit in ordered:
        if edit.end_line < edit.start_line - 1 or edit.start_line > len(lines) + 1:
            raise ToolError(
                f"Invalid edit {edit.start_line}:{edit.end_line}; the module has "
                f"{len(lines)} lines."
            )
        if edit.end_line > len(lines):
            raise ToolError(
                f"Invalid edit {edit.start_line}:{edit.end_line}; the module has "
                f"{len(lines)} lines."
            )
        if previous is not None and (
            edit.start_line <= previous.end_line or edit.start_line == previous.start_line
        ):
            raise ToolError("Line edits overlap; give each original line one edit.")
        previous = edit
    newline = "\r\n" if "\r\n" in source else "\n"
    for edit in reversed(ordered):
        replacement = edit.replacement
        if replacement and not replacement.endswith(("\r", "\n")) and edit.end_line < len(lines):
            replacement += newline
        lines[edit.start_line - 1 : edit.end_line] = [replacement] if replacement else []
    return "".join(lines)


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
    except Exception as exc:
        raise ToolError(
            f"Could not check which shapes call {module_name!r}: {exc}. Nothing was deleted. "
            "Inspect the drawing layer with xlide_list_shapes or Excel, fix the read error, "
            "then retry the deletion."
        ) from exc
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
    return found


__all__ = ["content_token", "register"]
