"""UserForms and Access designs: the controls, not just the code behind them.

A form's code is a module like any other and is read and written through the
module tools. Its design - which controls exist, how they nest, what properties
the developer set - lives beside it, and this is the only way to reach it without
opening the application.

One thing here is worth knowing before reading a result. MSForms stores a property
only when it differs from the control's default, so `properties` is what the
developer actually set, which no live host will tell you. A property missing from
the list is at its default, not absent.

Geometry is in points for UserForms, the unit the designer shows, and in twips for
Access, the unit Access keeps. The two are never mixed in one file.
"""

from __future__ import annotations

import contextlib
from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import project as project_layer
from .. import xlide_vscode
from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import require_writable, resolve_path
from ._common import bound, read_only, writes


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_list_forms",
        title="List forms",
        annotations=read_only("List forms"),
        description=(
            "Lists the UserForms in an Office file, or the forms and reports in an Access "
            "database, with each one's control count and, for Access, the sections a control "
            "can go in. A design's code is a module of the same name, read with "
            "xlide_read_module."
        ),
    )
    def list_forms(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            orphans = _orphaned(handle, info)
            listed = []
            for form in _forms(handle, info.title):
                entry: dict[str, Any] = {
                    "name": form.name,
                    "design": _design_kind(form),
                    "controls": len(form.walk()),
                }
                if form.name.casefold() in orphans:
                    entry["orphaned"] = True
                sections = _sections(form)
                if sections:
                    entry["sections"] = sections
                listed.append(entry)
        result: dict[str, Any] = {
            "path": str(path),
            "host": info.host,
            "geometry_unit": "twips" if info.host == "access" else "points",
            "count": len(listed),
            "forms": listed,
        }
        if any(entry.get("orphaned") for entry in listed):
            result["note"] = (
                "A form marked orphaned has a designer storage and no code module, so the "
                f"{info.title} editor does not show it and its code cannot be written. The "
                "bytes are still in the file. It is usually left by a tool that deleted or "
                "renamed the module without the design."
            )
        return result

    @server.tool(
        name="xlide_read_form",
        title="Read form design",
        annotations=read_only("Read form design"),
        description=(
            "Reads one form's design: every control with its name, type, the container it sits "
            "in, and the properties the developer set. Properties left at their default are "
            "not stored and so are not listed. Use it to understand a form's layout, or to see "
            "which control an event procedure belongs to."
        ),
    )
    def read_form(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        form_name: Annotated[str, Field(description="Form name, matched without case.")],
        include_properties: Annotated[
            bool,
            Field(
                default=True,
                description="Include each control's set properties. Off gives just the tree.",
            ),
        ] = True,
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            form = _find_form(handle, form_name, info.title)
            form_display = form.name
            design_kind = _design_kind(form)
            sections = _sections(form)
            every_control = [_control(c, include_properties) for c in form.walk()]
            form_properties = _safe_properties(form) if include_properties else {}
        controls, controls_note = bound(
            every_control,
            "controls",
            "Call again with include_properties false for the tree alone.",
        )
        result: dict[str, Any] = {
            "path": str(path),
            "form": form_display,
            "design": design_kind,
            "geometry_unit": "twips" if info.host == "access" else "points",
            "properties": form_properties,
            "control_count": len(every_control),
            "controls": controls,
        }
        if controls_note:
            result["note"] = controls_note
        if sections:
            result["sections"] = sections
        return result

    @server.tool(
        name="xlide_manage_form",
        title="Create, rename or delete a form",
        annotations=writes("Create, rename or delete a form", destructive=True),
        description=(
            "Creates a UserForm, or an Access form or report, and saves the file. A form is "
            "a designer storage and a code module of the same name, and this writes both, "
            "which is why it exists rather than xlide_write_module. Renaming and deleting "
            "work on Access designs, where both halves move together; for a UserForm they "
            "are refused, because nothing here can move the designer storage and doing half "
            "of it loses the form. Add controls afterwards with xlide_edit_form, and write "
            "its event procedures with xlide_write_module."
        ),
    )
    def manage_form(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        action: Annotated[str, Field(description="'create', 'rename' or 'delete'.")],
        form_name: Annotated[str, Field(description="The form or report to act on.")],
        new_name: Annotated[
            str, Field(default="", description="For rename: the new name.")
        ] = "",
        design: Annotated[
            str,
            Field(
                default="form",
                description="For create in Access: 'form' or 'report'. Elsewhere, a form.",
            ),
        ] = "form",
        caption: Annotated[
            str, Field(default="", description="For create: the caption it opens with.")
        ] = "",
        width: Annotated[
            float,
            Field(default=0.0, description="For create. Points for a UserForm, twips for Access."),
        ] = 0.0,
        height: Annotated[float, Field(default=0.0, description="For create.")] = 0.0,
        allow_protected: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
        allow_invalidate_signature: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_manage_form")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        if not info.supports_forms:
            raise ToolError(
                f"{info.title} projects have no form designer surface here. A Visual Basic 6 "
                "form's design is text in its own .frm file."
            )
        wanted = (action or "").strip().lower()
        if wanted not in {"create", "rename", "delete"}:
            raise ToolError("action must be 'create', 'rename' or 'delete'.")
        wanted_design = (design or "form").strip().lower()
        if wanted_design not in {"form", "report"}:
            raise ToolError("design must be 'form' or 'report'.")
        if wanted_design == "report" and info.host != "access":
            raise ToolError(f"Reports exist in Access databases; {path.name} is {info.title}.")

        with project_layer.open_project(path, info) as handle:
            # A form is a component like a module, so the first one in a file saved
            # before its first macro gives the file its project, as a module does.
            project_created = wanted == "create" and project_layer.ensure_project(
                handle, info, path
            )
            detail = _manage(handle, info, wanted, form_name, new_name, wanted_design,
                             caption, width, height)
            save_warnings = project_layer.save(
                handle,
                info,
                path=path,
                allow_protected=allow_protected,
                allow_invalidate_signature=allow_invalidate_signature,
            )

        result: dict[str, Any] = {
            "path": str(path),
            "action": wanted,
            "saved": True,
            **detail,
        }
        if save_warnings:
            result["warnings"] = save_warnings
        if project_created:
            result["vba_project_created"] = True
        if wanted == "create":
            result["next_step"] = (
                "Add controls with xlide_edit_form, and write the event procedures with "
                f"xlide_write_module on the module named {detail.get('form', form_name)!r}."
            )
        notice = xlide_vscode.file_changed(path, "forms", tool="xlide_manage_form")
        if notice:
            result["xlide_vscode"] = notice
        return result

    @server.tool(
        name="xlide_edit_form",
        title="Edit form design",
        annotations=writes("Edit form design", destructive=True),
        description=(
            "Changes one form's design and saves the file: add a control, remove one, or set a "
            "property on one. Geometry is in points for a UserForm and twips for an Access "
            "design. Setting a property to null clears it, putting the control back at its "
            "default. Removing a container takes its children with it, so ask the user first. "
            "Adding a control does not write its event procedure; do that with "
            "xlide_write_module on the form's code module."
        ),
    )
    def edit_form(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        form_name: Annotated[str, Field(description="Form to change.")],
        action: Annotated[
            str, Field(description="'add_control', 'remove_control' or 'set_property'.")
        ],
        control_name: Annotated[str, Field(description="The control to add, remove or change.")],
        control_type: Annotated[
            str,
            Field(
                default="",
                description=(
                    "For add_control: Label, TextBox, CommandButton, CheckBox, OptionButton, "
                    "ComboBox, ListBox, Frame, MultiPage, Image, SpinButton and the rest."
                ),
            ),
        ] = "",
        container: Annotated[
            str,
            Field(default="", description="For add_control: the Frame or page to put it in."),
        ] = "",
        section: Annotated[
            str,
            Field(
                default="",
                description=(
                    "For add_control on an Access design: the band to put it in, such as "
                    "Detail, PageHeaderSection or PageFooterSection. xlide_list_forms names "
                    "the ones a design has. Ignored for a UserForm, which has no bands."
                ),
            ),
        ] = "",
        caption: Annotated[
            str,
            Field(
                default="",
                description="For add_control on an Access design: the control's caption.",
            ),
        ] = "",
        left: Annotated[float, Field(default=6.0, description="For add_control.")] = 6.0,
        top: Annotated[float, Field(default=6.0, description="For add_control.")] = 6.0,
        width: Annotated[float, Field(default=0.0, description="0 uses the type's default.")] = 0.0,
        height: Annotated[
            float, Field(default=0.0, description="0 uses the type's default.")
        ] = 0.0,
        property_name: Annotated[
            str, Field(default="", description="For set_property: the property to set.")
        ] = "",
        property_value: Annotated[
            Any,
            Field(default=None, description="For set_property. null clears it to the default."),
        ] = None,
        allow_protected: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
        allow_invalidate_signature: Annotated[
            bool, Field(default=False, description="Ask the user first.")
        ] = False,
    ) -> dict[str, Any]:
        require_writable(settings, "xlide_edit_form")
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        wanted = (action or "").strip().lower()
        if wanted not in {"add_control", "remove_control", "set_property"}:
            raise ToolError(
                "action must be 'add_control', 'remove_control' or 'set_property'."
            )
        if not control_name.strip():
            raise ToolError("control_name is required.")

        with project_layer.open_project(path, info) as handle:
            form = _find_form(handle, form_name, info.title)
            form_display = form.name
            detail: dict[str, Any] = {}
            try:
                if wanted == "add_control":
                    if not control_type.strip():
                        raise ToolError("control_type is required for add_control.")
                    # An Access design places a control in a band and takes its
                    # caption at creation; a UserForm has neither, and its
                    # geometry is a float where Access wants twips.
                    if info.host == "access":
                        options: dict[str, Any] = {
                            "left": int(left),
                            "top": int(top),
                            "container": container.strip() or None,
                        }
                        if section.strip():
                            options["section"] = section.strip()
                        if caption:
                            options["caption"] = caption
                        if width:
                            options["width"] = int(width)
                        if height:
                            options["height"] = int(height)
                    else:
                        options = {
                            "container": container.strip() or None,
                            "left": left,
                            "top": top,
                            "width": width or None,
                            "height": height or None,
                        }
                    control = form.add_control(
                        control_type.strip(), control_name.strip(), **options
                    )
                    detail = {"added": control.name, "type": control_type.strip()}
                elif wanted == "remove_control":
                    victim = form.control(control_name.strip())
                    children = [c.name for c in getattr(victim, "children", [])]
                    form.remove_control(control_name.strip())
                    detail = {"removed": control_name.strip(), "children_removed": children}
                else:
                    if not property_name.strip():
                        raise ToolError("property_name is required for set_property.")
                    control = form.control(control_name.strip())
                    control.set_property(property_name.strip(), property_value)
                    detail = {
                        "control": control_name.strip(),
                        "property": property_name.strip(),
                        "value": property_value,
                        "cleared": property_value is None,
                    }
            except ToolError:
                raise
            except Exception as exc:
                raise ToolError(f"The form refused the change: {exc}") from exc

            # The form's designer storage is written by the project's own save;
            # `write_back` is an internal that takes the container, not a step a
            # caller performs.
            save_warnings = project_layer.save(
                handle,
                info,
                path=path,
                allow_protected=allow_protected,
                allow_invalidate_signature=allow_invalidate_signature,
            )

        result: dict[str, Any] = {
            "path": str(path),
            "form": form_display,
            "action": wanted,
            "saved": True,
            **detail,
        }
        if save_warnings:
            result["warnings"] = save_warnings
        notice = xlide_vscode.file_changed(path, "forms", tool="xlide_edit_form")
        if notice:
            result["xlide_vscode"] = notice
        return result


def _manage(
    handle: Any,
    info: Any,
    action: str,
    form_name: str,
    new_name: str,
    design: str,
    caption: str,
    width: float,
    height: float,
) -> dict[str, Any]:
    """Create, rename or delete a design, or refuse with the reason it cannot."""
    name = (form_name or "").strip()
    if not name:
        raise ToolError("form_name is required.")

    if action == "create":
        if any(f.name.casefold() == name.casefold() for f in _forms(handle, info.title)):
            raise ToolError(f"A form or report named {name!r} already exists in this file.")
        options: dict[str, Any] = {}
        if caption:
            options["caption"] = caption
        # A UserForm is measured in points and an Access design in twips, so the
        # number is passed through in whichever unit the host counts in.
        if width:
            options["width"] = int(width) if info.host == "access" else width
        if height:
            options["height"] = int(height) if info.host == "access" else height
        maker = handle.add_report if design == "report" else handle.add_form
        try:
            created = maker(name, **options)
        except Exception as exc:
            raise ToolError(f"The {design} could not be created: {exc}") from exc
        return {"form": created.name, "design": design, "created": True}

    # Renaming and deleting have to move the designer storage and the code
    # module together. Access does; nothing here can for a UserForm.
    if info.host != "access":
        raise ToolError(
            f"A UserForm's design is stored beside its code, and this server cannot move "
            f"the design, so {action} would lose the form. Ask the user to do it in the "
            f"{info.title} editor. Its code can still be written with xlide_write_module."
        )

    found = next(
        (f for f in _forms(handle, info.title) if f.name.casefold() == name.casefold()), None
    )
    if found is None:
        listed = ", ".join(f.name for f in _forms(handle, info.title)) or "(none)"
        raise ToolError(f"No form or report named {name!r}. In this file: {listed}.")
    kind = _design_kind(found)

    if action == "rename":
        if not new_name.strip():
            raise ToolError("new_name is required for action='rename'.")
        rename = handle.rename_report if kind == "report" else handle.rename_form
        try:
            rename(found.name, new_name.strip())
        except Exception as exc:
            raise ToolError(f"The {kind} could not be renamed: {exc}") from exc
        return {
            "renamed_from": found.name,
            "renamed_to": new_name.strip(),
            "design": kind,
            "form": new_name.strip(),
        }

    remove = handle.delete_report if kind == "report" else handle.delete_form
    try:
        remove(found.name)
    except Exception as exc:
        raise ToolError(f"The {kind} could not be deleted: {exc}") from exc
    return {"deleted": found.name, "design": kind}


def _forms(handle: Any, host_title: str) -> list[Any]:
    """Every design in a file: UserForms, or an Access database's forms AND reports.

    Access keeps reports in a collection of their own, so a reader that calls
    forms() alone reports a database's reports as not existing. They are the same
    kind of object, they are edited by the same calls, and `kind` tells them apart.
    """
    try:
        designs = list(handle.forms())
    except AttributeError as exc:
        raise ToolError(f"{host_title} files do not expose form designs here.") from exc
    except Exception as exc:
        raise ToolError(f"The form designs could not be read: {exc}") from exc

    reports = getattr(handle, "reports", None)
    if callable(reports):
        # A database with no reports storage at all: the forms still list.
        with contextlib.suppress(Exception):
            designs.extend(reports())
    return designs


def _find_form(handle: Any, name: str, host_title: str) -> Any:
    wanted = (name or "").strip().casefold()
    forms = _forms(handle, host_title)
    for form in forms:
        if form.name.casefold() == wanted:
            return form
    listed = ", ".join(f.name for f in forms) or "(none)"
    raise ToolError(f"No form or report named {name!r}. In this file: {listed}.")


def _orphaned(handle: Any, info: Any) -> set[str]:
    """Designs with a storage and no module, folded for comparison.

    pyOpenVBA puts it plainly: a storage without a module is not a component the
    host will show. Counting one as an ordinary form reports a form the user
    cannot see and cannot open, and saying nothing about it hides a file that
    needs repairing. Access names a design's module `Form_X`, so the pairing is
    only checkable on the hosts that name both the same.
    """
    if info.host == "access":
        return set()
    try:
        modules = {m.name.casefold() for m in handle.vba_project().modules}
        return {f.name.casefold() for f in handle.forms()} - modules
    except Exception:
        return set()


def _design_kind(design: Any) -> str:
    """'form' or 'report'. A UserForm has no kind and is always a form."""
    return str(getattr(design, "kind", "") or "form")


def _sections(design: Any) -> list[str]:
    """An Access design's bands. A UserForm has none."""
    try:
        return [section.name for section in getattr(design, "sections", []) or []]
    except Exception:
        return []


def _control(control: Any, include_properties: bool) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "name": control.name,
        "type": control.kind,
        "container": bool(getattr(control, "is_container", False)),
    }
    if include_properties:
        entry["properties"] = _safe_properties(control)
    return entry


def _safe_properties(owner: Any) -> dict[str, Any]:
    """Properties as plain JSON, with the ones nothing can name left out.

    An Access design stores property ids the library has no name for, and it
    stores a lot of them: 22 of a bare form's 29. Returning `Unidentified314: 4`
    beside `Caption: "Totals"` buries what a reader came for, and the number is
    not something an agent can act on. The count is reported instead, so the fact
    that something is there is not hidden.
    """
    try:
        raw = owner.properties()
    except Exception:
        return {}
    out: dict[str, Any] = {}
    unnamed = 0
    for key, value in raw.items():
        if key.startswith("Unidentified"):
            unnamed += 1
            continue
        if isinstance(value, (str, int, float, bool)) or value is None:
            out[key] = value
        elif isinstance(value, bytes):
            out[key] = f"<{len(value)} bytes>"
        else:
            out[key] = str(value)
    if unnamed:
        out["_unnamed_property_count"] = unnamed
    return out
