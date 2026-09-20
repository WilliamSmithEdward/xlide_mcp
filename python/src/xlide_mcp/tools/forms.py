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

from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import project as project_layer
from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import require_writable, resolve_path
from ._common import read_only, writes


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_list_forms",
        title="List forms",
        annotations=read_only("List forms"),
        description=(
            "Lists the UserForms in an Office file, or the forms and reports in an Access "
            "database, with each one's control count. A form's code is a module of the same "
            "name, read with xlide_read_module."
        ),
    )
    def list_forms(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        with project_layer.open_project(path, info) as handle:
            forms = _forms(handle, info.title)
            listed = [
                {"name": form.name, "controls": len(form.walk())} for form in forms
            ]
        return {
            "path": str(path),
            "host": info.host,
            "geometry_unit": "twips" if info.host == "access" else "points",
            "count": len(listed),
            "forms": listed,
        }

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
            controls = [_control(c, include_properties) for c in form.walk()]
            form_properties = _safe_properties(form) if include_properties else {}
        return {
            "path": str(path),
            "form": form_display,
            "geometry_unit": "twips" if info.host == "access" else "points",
            "properties": form_properties,
            "control_count": len(controls),
            "controls": controls,
        }

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
                    control = form.add_control(
                        control_type.strip(),
                        control_name.strip(),
                        container=container.strip() or None,
                        left=left,
                        top=top,
                        width=width or None,
                        height=height or None,
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
        return result


def _forms(handle: Any, host_title: str) -> list[Any]:
    try:
        return list(handle.forms())
    except AttributeError as exc:
        raise ToolError(f"{host_title} files do not expose form designs here.") from exc
    except Exception as exc:
        raise ToolError(f"The form designs could not be read: {exc}") from exc


def _find_form(handle: Any, name: str, host_title: str) -> Any:
    wanted = (name or "").strip().casefold()
    forms = _forms(handle, host_title)
    for form in forms:
        if form.name.casefold() == wanted:
            return form
    listed = ", ".join(f.name for f in forms) or "(none)"
    raise ToolError(f"No form named {name!r}. Forms in this file: {listed}.")


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
    """Properties as plain JSON. A value the transport cannot carry becomes its text."""
    try:
        raw = owner.properties()
    except Exception:
        return {}
    out: dict[str, Any] = {}
    for key, value in raw.items():
        if isinstance(value, (str, int, float, bool)) or value is None:
            out[key] = value
        elif isinstance(value, bytes):
            out[key] = f"<{len(value)} bytes>"
        else:
            out[key] = str(value)
    return out
