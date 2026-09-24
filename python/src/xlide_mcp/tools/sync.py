"""Modules to and from files on disk, for source control and review.

Exported .bas and .cls files are copies. They go stale the moment anything writes
the Office file, and an agent that edits an export has changed nothing until an
import runs. Both tools say so in their results, because "I updated the module"
after writing an exported copy is a report that is simply false, and it is an easy
one to make.

Export previews by default. An export that quietly overwrites a folder of reviewed
files is the kind of thing nobody notices until the diff is already lost.
"""

from __future__ import annotations

import difflib
from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import project as project_layer
from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import require_writable, resolve_directory, resolve_path
from ._common import writes

# A standard module exports as .bas; everything else - classes, document modules
# and the code behind a form - exports as .cls, which is what the VBE writes.
_STANDARD_SUFFIX = ".bas"
_OTHER_SUFFIX = ".cls"


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_export_modules",
        title="Export modules to files",
        annotations=writes("Export modules to files", destructive=True),
        description=(
            "Writes every VBA module in an Office file to a folder as .bas and .cls files, for "
            "source control or review. Previews by default: it reports what it would create, "
            "update and leave alone, and writes nothing until apply=true. Only when the user "
            "asks for it. The exported files are copies: editing one changes nothing inside "
            "the Office file until xlide_import_modules runs."
        ),
    )
    def export_modules(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        export_folder: Annotated[
            str,
            Field(
                default="",
                description=(
                    "Folder to write into. Empty uses a folder named after the file, beside it."
                ),
            ),
        ] = "",
        apply: Annotated[
            bool,
            Field(default=False, description="Write the files. False previews the plan."),
        ] = False,
        delete_stale: Annotated[
            bool,
            Field(
                default=False,
                description=(
                    "Delete .bas and .cls files in the folder that match no module. "
                    "Ask the user first."
                ),
            ),
        ] = False,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        _refuse_vb6(info, "xlide_export_modules")
        folder = (
            resolve_directory(export_folder, settings, must_exist=False)
            if export_folder.strip()
            else path.with_name(path.stem + "_vba")
        )

        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)

        plan: list[dict[str, Any]] = []
        expected: set[str] = set()
        for module in modules:
            suffix = _STANDARD_SUFFIX if module.kind == "standard" else _OTHER_SUFFIX
            target = folder / f"{module.name}{suffix}"
            expected.add(target.name.casefold())
            current = _read_text(target)
            if current is None:
                plan.append({"file": target.name, "action": "create", "lines": module.line_count})
            elif _same_text(current, module.full_source):
                plan.append({"file": target.name, "action": "unchanged"})
            else:
                plan.append(
                    {
                        "file": target.name,
                        "action": "update",
                        "diff_lines": _diff_size(current, module.full_source),
                    }
                )

        stale = [
            item.name
            for item in (folder.iterdir() if folder.is_dir() else [])
            if item.is_file()
            and item.suffix.lower() in {_STANDARD_SUFFIX, _OTHER_SUFFIX}
            and item.name.casefold() not in expected
        ]
        for name in stale:
            plan.append({"file": name, "action": "delete" if delete_stale else "stale"})

        if not apply:
            return {
                "path": str(path),
                "export_folder": str(folder),
                "applied": False,
                "plan": plan,
                "note": (
                    "Nothing was written. Call again with apply=true to write these files. "
                    "Exported files are copies and do not change the Office file."
                ),
            }

        require_writable(settings, "xlide_export_modules")
        folder.mkdir(parents=True, exist_ok=True)
        written = 0
        for module in modules:
            suffix = _STANDARD_SUFFIX if module.kind == "standard" else _OTHER_SUFFIX
            target = folder / f"{module.name}{suffix}"
            target.write_text(module.full_source, encoding="utf-8", newline="")
            written += 1
        deleted = 0
        if delete_stale:
            for name in stale:
                (folder / name).unlink(missing_ok=True)
                deleted += 1
        return {
            "path": str(path),
            "export_folder": str(folder),
            "applied": True,
            "files_written": written,
            "files_deleted": deleted,
            "plan": plan,
            "note": (
                "These are copies. Editing one changes nothing inside the Office file until "
                "xlide_import_modules runs."
            ),
        }

    @server.tool(
        name="xlide_import_modules",
        title="Import modules from files",
        annotations=writes("Import modules from files", destructive=True),
        description=(
            "Reads .bas and .cls files from a folder back into an Office file's VBA project "
            "and saves it. Previews by default: it reports which modules would change and by "
            "how much, and writes nothing until apply=true. A file whose name matches no "
            "module creates one. Document modules such as ThisWorkbook are written but never "
            "created. After applying, call xlide_analyze."
        ),
    )
    def import_modules(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        source_folder: Annotated[str, Field(description="Folder holding the .bas/.cls files.")],
        apply: Annotated[
            bool,
            Field(default=False, description="Write the modules. False previews the plan."),
        ] = False,
        allow_protected: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
        allow_invalidate_signature: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        _refuse_vb6(info, "xlide_import_modules")
        folder = resolve_directory(source_folder, settings)

        incoming: dict[str, tuple[Path, str]] = {}
        for item in sorted(folder.iterdir()):
            if not item.is_file() or item.suffix.lower() not in {
                _STANDARD_SUFFIX,
                _OTHER_SUFFIX,
            }:
                continue
            text = _read_text(item)
            if text is None:
                continue
            incoming[item.stem.casefold()] = (item, text)
        if not incoming:
            raise ToolError(
                f"No .bas or .cls files in {folder}. Export first with xlide_export_modules, "
                "or point source_folder at the folder the files are in."
            )

        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
            by_name = {m.name.casefold(): m for m in modules}

            plan: list[dict[str, Any]] = []
            changing: list[tuple[str, str, Path, str]] = []
            for key, (item, text) in incoming.items():
                module = by_name.get(key)
                if module is None:
                    plan.append({"module": item.stem, "file": item.name, "action": "create"})
                    changing.append(("create", item.stem, item, text))
                elif _same_text(module.full_source, text):
                    plan.append({"module": module.name, "file": item.name, "action": "unchanged"})
                else:
                    plan.append(
                        {
                            "module": module.name,
                            "file": item.name,
                            "action": "update",
                            "diff_lines": _diff_size(module.full_source, text),
                        }
                    )
                    changing.append(("update", module.name, item, text))
            untouched = [m.name for m in modules if m.name.casefold() not in incoming]

            if not apply:
                return {
                    "path": str(path),
                    "source_folder": str(folder),
                    "applied": False,
                    "plan": plan,
                    "modules_without_a_file": untouched,
                    "note": (
                        f"{len(changing)} modules would change. Nothing was written. "
                        "Call again with apply=true."
                    ),
                }

            require_writable(settings, "xlide_import_modules")
            import pyopenvba

            for action, name, item, text in changing:
                if action == "create":
                    kind = (
                        pyopenvba.VBAModuleKind.standard
                        if item.suffix.lower() == _STANDARD_SUFFIX
                        else pyopenvba.VBAModuleKind.other
                    )
                    handle.vba_project().add_module(name, text, kind=kind)
                else:
                    handle.set_module(name, text)

            save_warnings = project_layer.save(
                handle,
                info,
                path=path,
                allow_protected=allow_protected,
                allow_invalidate_signature=allow_invalidate_signature,
            )

        result: dict[str, Any] = {
            "path": str(path),
            "source_folder": str(folder),
            "applied": True,
            "modules_changed": len(changing),
            "plan": plan,
            "saved": True,
            "next_step": "Call xlide_analyze and fix anything at error severity.",
        }
        if save_warnings:
            result["warnings"] = save_warnings
        return result


def _refuse_vb6(info: Any, operation: str) -> None:
    """A VB6 project's modules are already files on disk.

    Exporting them would write a second copy under the wrong extension - a form
    is a .frm, not the .cls an export names a non-standard module - and importing
    would put back what is already there.
    """
    if info.host == "vb6":
        raise ToolError(
            f"{operation} moves modules between a container and files on disk. A Visual "
            "Basic 6 project's modules are already files: edit them where they are, or "
            "read and write them with xlide_read_module and xlide_write_module."
        )


def _read_text(path: Path) -> str | None:
    """A module file's text. VBE exports are ANSI, so UTF-8 is tried first."""
    if not path.is_file():
        return None
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def _same_text(left: str, right: str) -> bool:
    """Equality that ignores line-ending style, which is not somebody's edit."""
    return left.replace("\r\n", "\n").rstrip() == right.replace("\r\n", "\n").rstrip()


def _diff_size(before: str, after: str) -> int:
    before_lines = before.replace("\r\n", "\n").splitlines()
    after_lines = after.replace("\r\n", "\n").splitlines()
    matcher = difflib.SequenceMatcher(None, before_lines, after_lines, autojunk=False)
    changed = 0
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            changed += max(i2 - i1, j2 - j1)
    return changed
