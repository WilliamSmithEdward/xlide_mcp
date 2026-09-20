"""Visual Basic 6 projects, presented the way an Office file's VBA project is.

A .vbp is a text manifest and its modules are ordinary files on disk, so an agent
with file tools can already edit them. What it cannot do on its own is read the
project as a project: which files are modules rather than resources, what each one
is called inside VB (a form's name is an attribute, not its file name), which
libraries the project references, and above all analyze the modules together so a
call from one into another resolves.

So this presents a .vbp through the same surface as everything else. The same
read, the same write, the same analysis, and one workflow for the agent whatever
it opened.

Three things are different from an Office project and each shows in the results.
Writes land on disk as they are made rather than at a save. Nothing is ever
password-protected or digitally signed, because a .vbp has nowhere to put either.
And a module carries more header than a VBA one: a form or a control keeps a
designer block ahead of its attributes, which is preserved untouched, because it
is what VB uses to draw the thing and it is not code.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .errors import ToolError

# Manifest lines that name a module, and what kind each one is. A form and a
# control give only a file name; the others give `Name; File`.
_MODULE_LINES: dict[str, str] = {
    "Module": "standard",
    "Class": "class",
    "Form": "form",
    "UserControl": "usercontrol",
    "PropertyPage": "propertypage",
    "MDIForm": "form",
    "Designer": "designer",
}

# Extensions VB6 writes a module to, for the reverse lookup when a manifest line
# carries only a file name.
MODULE_EXTENSIONS = frozenset({".bas", ".cls", ".frm", ".ctl", ".pag", ".dsr"})

_ATTRIBUTE_NAME = re.compile(r'^Attribute\s+VB_Name\s*=\s*"([^"]*)"', re.MULTILINE)

# A .vbp is written in the machine's ANSI code page, and so are the modules.
_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")


@dataclass
class Vb6Component:
    """One module of a VB6 project, split the way a write has to put it back."""

    name: str
    xlide_kind: str
    path: Path
    designer: str = ""
    """A form's or control's `VERSION ... Begin ... End` block. Not code."""
    header: str = ""
    """The contiguous `Attribute VB_*` lines."""
    body: str = ""

    @property
    def kind(self) -> str:
        """The coarse kind, spelled the way the Access container spells it."""
        return "module" if self.xlide_kind == "standard" else "class"

    @property
    def source(self) -> str:
        """The whole file, as `read_modules` expects a component's source."""
        return self.designer + self.header + self.body


@dataclass
class Vb6Reference:
    name: str
    libid: str


@dataclass
class _Manifest:
    name: str = ""
    kind: str = ""
    components: list[Vb6Component] = field(default_factory=list)
    references: list[Vb6Reference] = field(default_factory=list)
    objects: list[str] = field(default_factory=list)
    missing: list[str] = field(default_factory=list)
    """Files the manifest names and the folder does not hold."""


class Vb6Project:
    """A .vbp, with the surface the rest of this server reads a project through.

    It is deliberately shaped like a pyOpenVBA container rather than given a
    surface of its own: every module tool, the analyzer and the search then work
    on a VB6 project without knowing there is one.
    """

    def __init__(self, path: Path) -> None:
        self.path = path
        self._manifest = _parse(path)

    # -- the container surface --------------------------------------------

    def __enter__(self) -> Vb6Project:
        return self

    def __exit__(self, *exc_info: object) -> None:
        return None

    def vba_project(self) -> Vb6Project:
        """A .vbp is its own project; there is no inner object to hand back."""
        return self

    @property
    def name(self) -> str:
        return self._manifest.name or self.path.stem

    @property
    def modules(self) -> list[Vb6Component]:
        return self._manifest.components

    @property
    def protection(self) -> None:
        """A .vbp has nowhere to put a password, so there is never one to report."""
        return None

    @property
    def references(self) -> list[Vb6Reference]:
        return self._manifest.references

    def module_names(self) -> list[str]:
        return [component.name for component in self._manifest.components]

    def vba_project_bytes(self) -> bytes:
        """No vbaProject.bin exists, so signature detection has nothing to find."""
        return b""

    def forms(self) -> list[Any]:
        """A VB6 form's design is text in its own .frm, not a separate storage.

        Returning none is the honest answer for the designer tools: what they
        edit is an MSForms control tree, and this is not one.
        """
        return []

    def validate(self) -> list[str]:
        """Structural problems: a manifest naming a file the folder does not hold."""
        return [
            f"The project names {name}, and the folder does not hold it."
            for name in self._manifest.missing
        ]

    # -- reading and writing ----------------------------------------------

    def get_module(self, name: str) -> str:
        return self._find(name).source

    def set_module(self, name: str, source: str) -> None:
        """Replace a module's code. The designer block and attributes are kept.

        A caller that sends a whole file, designer block and all, is taken at its
        word; a caller that sends a body gets the module's own header back in
        front of it, the same bargain the VBA hosts make.
        """
        component = self._find(name)
        designer, header, body = _split(source, component.xlide_kind)
        component.designer = designer or component.designer
        component.header = header or component.header
        component.body = body
        self._write(component)

    def add_module(self, name: str, source: str, *, kind: Any = None, **_: Any) -> Vb6Component:
        """Create a module, its file, and the manifest line that names it."""
        if self._has(name):
            raise ToolError(f"A module named {name!r} is already in this project.")
        wanted = _kind_for(kind)
        suffix = {"standard": ".bas", "class": ".cls"}[wanted]
        target = self.path.with_name(f"{name}{suffix}")
        if target.exists():
            raise ToolError(
                f"{target.name} already exists beside the project but is not one of its "
                "modules. Rename one of them, or add the existing file to the project in VB6."
            )
        designer, header, body = _split(source, wanted)
        component = Vb6Component(
            name=name,
            xlide_kind=wanted,
            path=target,
            designer=designer or ("VERSION 1.0 CLASS\r\n" if wanted == "class" else ""),
            header=header or f'Attribute VB_Name = "{name}"\r\n',
            body=body,
        )
        self._manifest.components.append(component)
        self._write(component)
        self._write_manifest()
        return component

    def rename_module(self, old_name: str, new_name: str) -> Vb6Component:
        """Rename the module, its file, and the manifest line, together.

        VB6 keeps the name in three places. Changing one of them produces a
        project that will not load, so they move as one or not at all.
        """
        component = self._find(old_name)
        target = component.path.with_name(f"{new_name}{component.path.suffix}")
        if target.exists() and target != component.path:
            raise ToolError(f"{target.name} already exists beside the project.")

        component.header = _renamed_header(component.header, new_name)
        previous = component.path
        component.name = new_name
        component.path = target
        self._write(component)
        if previous != target:
            previous.unlink(missing_ok=True)
        self._write_manifest()
        return component

    def delete_module(self, name: str) -> None:
        """Remove the module from the manifest and delete its file."""
        component = self._find(name)
        self._manifest.components.remove(component)
        component.path.unlink(missing_ok=True)
        # A form keeps its binary resources in a sidecar of the same stem.
        for sidecar in (".frx", ".ctx", ".pgx", ".dsx"):
            component.path.with_suffix(sidecar).unlink(missing_ok=True)
        self._write_manifest()

    def save(self, *_: Any, **__: Any) -> None:
        """Nothing to do: a VB6 write lands on disk as it is made."""
        return None

    # -- internals ---------------------------------------------------------

    def _find(self, name: str) -> Vb6Component:
        wanted = (name or "").strip().casefold()
        for component in self._manifest.components:
            if component.name.casefold() == wanted:
                return component
        listed = ", ".join(self.module_names()) or "(none)"
        raise ToolError(f"No module named {name!r}. Modules in this project: {listed}.")

    def _has(self, name: str) -> bool:
        wanted = (name or "").strip().casefold()
        return any(c.name.casefold() == wanted for c in self._manifest.components)

    def _write(self, component: Vb6Component) -> None:
        try:
            component.path.write_text(component.source, encoding="cp1252", newline="")
        except UnicodeEncodeError:
            # A module holding characters the ANSI page cannot carry. UTF-8 keeps
            # the text; VB6 will read it as mojibake, which is visible, where
            # dropping the characters would not be.
            component.path.write_text(component.source, encoding="utf-8", newline="")
        except OSError as exc:
            raise ToolError(f"{component.path.name} could not be written: {exc}.") from exc

    def _write_manifest(self) -> None:
        """Rewrite only the module lines, leaving every other setting alone.

        A .vbp carries compiler switches, version numbers and paths this server
        has no business reforming, so the file is edited rather than regenerated.
        """
        original = _read_text(self.path)
        lines = original.splitlines()
        kept = [
            line
            for line in lines
            if line.split("=", 1)[0].strip() not in _MODULE_LINES
        ]
        module_lines = [_manifest_line(c) for c in self._manifest.components]
        # The module lines go where the first one was, so a hand-ordered file
        # keeps its shape.
        insert_at = next(
            (
                index
                for index, line in enumerate(lines)
                if line.split("=", 1)[0].strip() in _MODULE_LINES
            ),
            0,
        )
        insert_at = min(insert_at, len(kept))
        rebuilt = kept[:insert_at] + module_lines + kept[insert_at:]
        try:
            self.path.write_text("\r\n".join(rebuilt) + "\r\n", encoding="cp1252", newline="")
        except OSError as exc:
            raise ToolError(f"{self.path.name} could not be written: {exc}.") from exc


# ------------------------------------------------------------------ parsing


def _parse(path: Path) -> _Manifest:
    text = _read_text(path)
    manifest = _Manifest()
    for line in text.splitlines():
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip()
        if key == "Name":
            manifest.name = value.strip('"')
        elif key == "Type":
            manifest.kind = value
        elif key == "Reference":
            manifest.references.append(_reference(value))
        elif key == "Object":
            manifest.objects.append(value)
        elif key in _MODULE_LINES:
            component = _component(path, key, value, manifest)
            if component is not None:
                manifest.components.append(component)
    return manifest


def _component(
    project: Path, key: str, value: str, manifest: _Manifest
) -> Vb6Component | None:
    """One manifest module line, read from its file.

    `Module=Name; File.bas` names the module; `Form=File.frm` does not, and the
    form's real name is an attribute inside it.
    """
    declared, _, file_part = value.partition(";")
    file_name = (file_part or declared).strip()
    if not file_name:
        return None
    file_path = (project.parent / file_name.replace("\\", "/")).resolve()
    if not file_path.is_file():
        manifest.missing.append(file_name)
        return None

    text = _read_text(file_path)
    kind = _MODULE_LINES[key]
    designer, header, body = _split(text, kind)
    attribute = _ATTRIBUTE_NAME.search(header)
    name = (
        attribute.group(1)
        if attribute
        else (declared.strip() if file_part else file_path.stem)
    )
    return Vb6Component(
        name=name or file_path.stem,
        xlide_kind=kind,
        path=file_path,
        designer=designer,
        header=header,
        body=body,
    )


def _split(text: str, kind: str) -> tuple[str, str, str]:
    """A module file as (designer block, attribute header, code body).

    The designer block is a form's or a control's `VERSION ... Begin ... End`.
    It is not code, the analyzer must not see it, and a write must put it back
    exactly, so it is kept whole and separate rather than parsed.
    """
    from pyopenvba.vba import split_attribute_header
    from pyvbaanalysis.reader import strip_export_header

    if not text:
        return "", "", ""
    remainder = strip_export_header(text)
    designer = text[: len(text) - len(remainder)] if remainder != text else ""
    if kind == "class" and not designer and remainder.startswith("VERSION"):
        # A .cls keeps `VERSION 1.0 CLASS` with no Begin block behind it.
        head, _, rest = remainder.partition("\n")
        designer, remainder = head + "\n", rest
    header, body = split_attribute_header(remainder)
    return designer, header, body


def _renamed_header(header: str, new_name: str) -> str:
    if _ATTRIBUTE_NAME.search(header):
        return _ATTRIBUTE_NAME.sub(f'Attribute VB_Name = "{new_name}"', header, count=1)
    return f'Attribute VB_Name = "{new_name}"\r\n' + header


def _manifest_line(component: Vb6Component) -> str:
    key = next(k for k, v in _MODULE_LINES.items() if v == component.xlide_kind)
    file_name = component.path.name
    if component.xlide_kind in {"form", "usercontrol", "propertypage", "designer"}:
        return f"{key}={file_name}"
    return f"{key}={component.name}; {file_name}"


def _reference(value: str) -> Vb6Reference:
    """`*\\G{guid}#ver#lcid#path#description`. The description names it."""
    parts = value.split("#")
    name = parts[4].strip() if len(parts) >= 5 else value
    return Vb6Reference(name=name, libid=value)


def _read_text(path: Path) -> str:
    """VB6 writes ANSI. UTF-8 is tried first for a file some other tool touched."""
    try:
        data = path.read_bytes()
    except OSError as exc:
        raise ToolError(f"{path.name} could not be read: {exc}.") from exc
    for encoding in _ENCODINGS:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def _kind_for(kind: Any) -> str:
    """Map whatever the module tools passed onto a VB6 module kind."""
    text = str(getattr(kind, "name", kind) or "standard").lower()
    return "class" if "other" in text or "class" in text else "standard"


def is_vb6(path: Path) -> bool:
    return path.suffix.lower() == ".vbp"
