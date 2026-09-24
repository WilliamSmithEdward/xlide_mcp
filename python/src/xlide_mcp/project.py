"""The VBA project as this server presents it: modules, kinds, and guarded saves.

pyOpenVBA gives each component's full stream text and a coarse standard/other
kind. Two refinements happen here and nowhere else, so every tool sees the same
module:

* The kind is refined to standard / class / document / userform, using the same
  classifier pyvbaanalysis uses, so a module's kind in a listing and its kind in a
  diagnostic are the one fact.
* The `Attribute VB_*` header is stripped from what a read returns. That is the
  surface the VBE shows and the text an agent should be editing; pyOpenVBA
  re-prepends the module's own header on write, so a body written back keeps the
  attributes the host needs.

Saves carry two guards that both mean "stop and ask the user", never "try harder":
a password-protected project, and a digital signature the save would invalidate.

A file can also hold no project at all. Excel, Word and PowerPoint write no
vbaProject.bin into a macro-enabled file until its first macro exists, so a
.xlsm nobody has written code in yet is an ordinary file, not a damaged one. The
listings answer it as empty, a read of something named says the file has no
project, and the first module written into a .xlsm, .docm or .pptm gives it the
project its application would have made.
"""

from __future__ import annotations

import contextlib
import warnings
from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from . import locks
from .errors import ToolError
from .hosts import HostInfo, container, require_readable
from .tokens import content_token

ModuleKind = Literal["standard", "class", "document", "userform"]

# A VB6 manifest names more kinds than VBA has. They map onto the four this
# server reports, so one vocabulary reaches the caller whatever was opened.
_VB6_KINDS = {
    "standard": "standard",
    "class": "class",
    "form": "userform",
    "usercontrol": "userform",
    "propertypage": "userform",
    "designer": "class",
}

# Modules the host owns. They can be written, and cannot be renamed or deleted:
# the host recreates them and the project stops matching the document.
_DOCUMENT_KIND: ModuleKind = "document"


@dataclass(frozen=True)
class ModuleView:
    """One module, as every tool in this server sees it."""

    name: str
    kind: ModuleKind
    body: str
    """Source with the attribute header stripped: what the VBE shows."""
    full_source: str
    """Source as the project stores it, header included."""

    header_lines: int = 0
    """Lines of attribute header. What analysis positions have to be shifted by."""

    @property
    def line_count(self) -> int:
        return len(self.body.splitlines())

    @property
    def token(self) -> str:
        return content_token(self.body)

    def summary(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "kind": self.kind,
            "lines": self.line_count,
            "content_token": self.token,
        }


@dataclass(frozen=True)
class ProjectStatus:
    """What a caller has to know before writing to a project."""

    project_name: str
    module_count: int
    password_protected: bool
    digitally_signed: bool | None
    """None where this server does not detect signatures for the host (Access)."""

    has_project: bool = True
    """False for a macro-enabled file saved before its first macro existed."""

    def warnings(self) -> list[str]:
        out: list[str] = []
        if self.password_protected:
            out.append(
                "The VBA project is password-protected. A write needs allow_protected=true, "
                "and the password itself is neither read nor changed. Ask the user first."
            )
        if self.digitally_signed:
            out.append(
                "The VBA project is digitally signed. Any change to the macros invalidates "
                "the signature, and the save drops it, so a write needs "
                "allow_invalidate_signature=true. Ask the user first."
            )
        return out


# The formats a project can be added to. The applications were measured making
# one for a first macro in these three; a binary .xls, .doc or .ppt keeps its
# project in storages this server does not create, and a .xlsb its workbook
# properties in binary records.
_PROJECT_CAN_BE_ADDED = frozenset({".xlsm", ".docm", ".pptm"})


def can_add_project(info: HostInfo) -> bool:
    """Whether a file of this format can be given a VBA project it does not have."""
    return info.extension in _PROJECT_CAN_BE_ADDED


def has_project(handle: Any, info: HostInfo) -> bool:
    """Whether the file holds a VBA project, empty or not.

    Access always does: a blank database Access writes carries an empty project,
    so a missing one there is damage rather than a normal state, and it stays an
    error. A VB6 project is its manifest.
    """
    check = getattr(handle, "has_vba_project", None)
    if info.host in {"access", "vb6"} or not callable(check):
        return True
    return bool(check())


def no_project_note(path: Path, info: HostInfo) -> str:
    """What to tell an agent about a file that has no VBA project yet."""
    if info.extension in _PROJECT_CAN_BE_ADDED:
        return (
            f"{path.name} has no VBA project yet. {info.title} writes none until the first "
            "macro exists, so this is a normal file, not a damaged one. xlide_write_module "
            "creates the project along with the first module."
        )
    return (
        f"{path.name} has no VBA project. A {info.extension} file keeps one in storages this "
        f"server does not create, so the first macro has to be written in {info.title}, or "
        f"the file saved as a macro-enabled OOXML file first."
    )


def ensure_project(handle: Any, info: HostInfo, path: Path) -> bool:
    """Give a file with no VBA project the one its application makes for a first macro.

    Returns whether a project was created. pyOpenVBA builds it as the application
    does: in Excel a document module for the workbook and one per sheet, in Word
    ThisDocument, in PowerPoint nothing until a module is added.
    """
    if has_project(handle, info):
        return False
    if info.extension not in _PROJECT_CAN_BE_ADDED:
        raise ToolError(no_project_note(path, info))
    handle.add_vba_project()
    return True


def form_module_names(handle: Any, info: HostInfo) -> set[str]:
    """The modules that are the code behind a design, folded for comparison.

    A UserForm's module is a class as far as its text is concerned: what makes it
    a form is the designer storage beside it, which the text cannot see. Reading
    the kind from the text alone calls every form a class, and a class is a thing
    this server will happily rename - which silently separates the form from its
    code.

    Access needs no such lookup: it names a design's module `Form_X` or
    `Report_X`, and reading its designs to answer a module listing would be work
    for nothing.
    """
    if info.host == "access":
        return set()
    try:
        return {form.name.casefold() for form in handle.forms()}
    except Exception:
        return set()


def _access_design_module(name: str) -> bool:
    return name.casefold().startswith(("form_", "report_"))


def read_modules(handle: Any, info: HostInfo) -> list[ModuleView]:
    """Every module in an open project, in the project's own order."""
    import pyopenvba
    from pyopenvba.vba import split_attribute_header
    from pyvbaanalysis.reader import classify_module_kind

    form_names = form_module_names(handle, info)
    views: list[ModuleView] = []
    for component in _components(handle):
        try:
            source = component.source
        except pyopenvba.PyOpenVBAError as exc:
            raise ToolError(
                f"Module {component.name!r} could not be decompressed: {exc}. "
                "The VBA project may be damaged; xlide_validate_project reports what it can see."
            ) from exc
        # A VB6 component knows exactly what it is, because the project manifest
        # said so. Classifying it from its text would turn a form into a class.
        declared = getattr(component, "xlide_kind", "")
        if declared:
            kind_value = _VB6_KINDS.get(declared, declared)
        elif component.name.casefold() in form_names or (
            info.host == "access" and _access_design_module(component.name)
        ):
            kind_value = "userform"
        else:
            kind_value = classify_module_kind(
                source,
                extension=info.extension,
                pyopenvba_standard=is_standard_component(component),
            ).value
        # The attribute header is what the VBE hides, so the body is what an agent
        # should be reading and editing. pyvbaanalysis keeps attributes instead,
        # because to its parser they are ordinary module statements; that is the
        # analysis surface, not this one, and the two are deliberately different.
        if declared:
            # A VB6 component has already split itself, and it has to: a form
            # keeps a `VERSION ... Begin ... End` designer block that
            # split_attribute_header does not recognize, so letting it try hands
            # the caller VB's form markup as though it were code to edit.
            body = component.body
            header = source[: len(source) - len(body)]
        else:
            header, body = split_attribute_header(source)
        views.append(
            ModuleView(
                name=component.name,
                kind=kind_value,  # type: ignore[arg-type]
                body=body,
                full_source=source,
                header_lines=len(header.splitlines()),
            )
        )
    return views


def is_standard_component(component: Any) -> bool:
    """Whether a component is a standard module, whichever host it came from.

    The two hosts spell it differently, and comparing against only one of them is
    how every Access standard module came back classified as a class: a package
    host gives a `VBAModuleKind` enum, and Access gives the string `"module"`.
    An Access module carries no designer header, so the text cannot settle it and
    this is the only thing that can.
    """
    import pyopenvba

    kind = component.kind
    if isinstance(kind, str):
        return kind.casefold() == "module"
    return bool(kind == pyopenvba.VBAModuleKind.standard)


def find_module(
    modules: list[ModuleView], name: str, *, no_project: str = ""
) -> ModuleView:
    """Look a module up the way VBA compares names: without regard to case.

    `no_project` is the note for a file with no VBA project, which is the answer
    to give instead of an empty list of the modules it does not have.
    """
    wanted = name.strip().casefold()
    for module in modules:
        if module.name.casefold() == wanted:
            return module
    if no_project and not modules:
        raise ToolError(f"No module named {name!r}: {no_project}")
    known = ", ".join(m.name for m in modules) or "(none)"
    raise ToolError(f"No module named {name!r}. Modules in this project: {known}.")


def find_module_in(
    handle: Any,
    info: HostInfo,
    path: Path,
    name: str,
    modules: list[ModuleView] | None = None,
) -> ModuleView:
    """`find_module` over an open file, saying so when the file has no project.

    Pass `modules` when the caller has already read them, so they are not read
    twice.
    """
    if modules is None:
        modules = read_modules(handle, info)
    note = "" if modules or has_project(handle, info) else no_project_note(path, info)
    return find_module(modules, name, no_project=note)


def project_status(handle: Any, info: HostInfo) -> ProjectStatus:
    """Name, module count, and the two states that gate a write."""
    if not has_project(handle, info):
        return ProjectStatus(
            project_name="",
            module_count=0,
            password_protected=False,
            digitally_signed=False,
            has_project=False,
        )
    project = handle.vba_project()
    protection = getattr(project, "protection", None)
    protected = bool(getattr(protection, "has_password", False))
    if info.host == "access":
        # A probe: an Access build that cannot answer leaves the PROJECT-stream
        # reading in place rather than failing the whole summary.
        with contextlib.suppress(Exception):
            protected = bool(handle.vba_is_protected())
    return ProjectStatus(
        project_name=getattr(project, "name", "") or Path(getattr(handle, "path", "")).stem,
        module_count=len(list(_components(handle))),
        password_protected=protected,
        digitally_signed=_signature_present(handle, info),
    )


def save(
    handle: Any,
    info: HostInfo,
    *,
    path: Path | None = None,
    allow_protected: bool = False,
    allow_invalidate_signature: bool = False,
) -> list[str]:
    """Write the project back. Returns whatever the save warned about.

    A refused save raises with the flag that would allow it named, because the
    decision is the user's and the agent needs to be able to describe it to them.
    `path` names the file in a lock refusal: pyOpenVBA's handles do not carry it.
    """
    import pyopenvba

    # A signature lives beside a vbaProject.bin, which an Access database does not
    # have, so its save takes no such flag and passing one is a TypeError rather
    # than a refusal.
    options: dict[str, Any] = {"allow_protected": allow_protected}
    if info.host != "access":
        options["allow_invalidate_signature"] = allow_invalidate_signature
        # pyOpenVBA drops a stale signature and warns, which is how Office saves
        # changed code too. Warning after the fact is not asking first, so the
        # refusal is made here, before anything is written.
        if not allow_invalidate_signature and _signature_present(handle, info):
            raise ToolError(
                f"The VBA project is digitally signed, and saving changed code drops the "
                f"signature: {info.title} signs the code, and this is no longer the code it "
                "signed. Nothing was written. Ask the user, then call again with "
                "allow_invalidate_signature=true. They can sign it again in the VBA editor."
            )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        try:
            handle.save(**options)
        except pyopenvba.PyOpenVBAError as exc:
            # The base class: Access refuses a protected project with its own
            # AccessError, which is not a VBAProjectError.
            raise _save_refusal(exc, info) from exc
        except PermissionError as exc:
            if path is None:
                raise ToolError(
                    f"The file is locked, most likely open in {info.title}. Nothing was "
                    f"written. Ask the user to close it in {info.title} and try again."
                ) from exc
            raise ToolError(locks.lock_message(path, info.title)) from exc
        except OSError as exc:
            raise ToolError(f"The file could not be written: {exc}.") from exc
    return [str(w.message) for w in caught]


def _save_refusal(exc: Exception, info: HostInfo) -> ToolError:
    text = str(exc)
    if "protect" in text.lower():
        return ToolError(
            f"{info.title} refused the write: the VBA project is password-protected. "
            "Ask the user whether to proceed, then call again with allow_protected=true. "
            f"{text}"
        )
    return ToolError(f"The VBA project refused the write: {text}")


@dataclass(frozen=True)
class AnalysisInput:
    """One module as the analyzer sees it, and how to map its positions back.

    The analyzer is given the attribute header, because to its parser those are
    ordinary module statements and some rules read them. A reader is given the
    body, because that is what the VBA editor shows. The two texts differ by a
    known number of leading lines, and `line_offset` is that number: without it,
    every diagnostic in a module with a header points an agent at the wrong line.
    """

    name: str
    module_input: Any
    analyzed_source: str
    line_offset: int


def analysis_inputs(modules: list[ModuleView]) -> list[AnalysisInput]:
    """The project as pyvbaanalysis wants it, with the position mapping alongside."""
    from pyvbaanalysis import ModuleInput, ModuleSymbolKind
    from pyvbaanalysis.reader import strip_export_header

    prepared: list[AnalysisInput] = []
    for module in modules:
        analyzed = strip_export_header(module.full_source)
        offset = len(analyzed.splitlines()) - len(module.body.splitlines())
        prepared.append(
            AnalysisInput(
                name=module.name,
                module_input=ModuleInput(
                    module_name=module.name,
                    module_kind=ModuleSymbolKind(module.kind),
                    source=analyzed,
                ),
                analyzed_source=analyzed,
                line_offset=max(0, offset),
            )
        )
    return prepared


def references(handle: Any, info: HostInfo) -> list[Any]:
    """The libraries a project declares, in priority order, or none for no project.

    One call for every host since pyOpenVBA 6.1: Access keeps its list in the
    database and the others in the dir stream, and both answer `references()`.
    """
    import pyopenvba

    if info.host == "vb6":
        # A .vbp's manifest lists them; it is a property there, not a call.
        return list(handle.references)
    try:
        return list(handle.references())
    except pyopenvba.NoVBAProjectError:
        return []


def referenced_hosts(handle: Any, info: HostInfo) -> list[str] | None:
    """The other Office libraries a project references, as the analyzer names them.

    What lets it report `Dim doc As Word.Document` in a workbook that does not
    reference Word, and check code that does against Word's model as well as
    Excel's. None means the list is unknown, which the analyzer treats as nothing
    proven missing; an empty list means it is known to name nothing else.
    """
    if info.host == "vb6":
        return None
    from pyvbaanalysis.host import referenced_host_tokens

    try:
        declared = references(handle, info)
    except Exception:
        return None
    libids = [ref.libid for ref in declared if isinstance(getattr(ref, "libid", None), str)]
    return referenced_host_tokens(info.host, libids)


def open_project(path: Path, info: HostInfo | None = None) -> container:
    """`with open_project(path) as book:`"""
    return container(path, info or require_readable(path))


def _components(handle: Any) -> Iterator[Any]:
    """Every module component of an open project, whatever the host.

    A file with no project has none, which is an answer rather than a failure.
    """
    import pyopenvba

    try:
        return iter(handle.vba_project().modules)
    except pyopenvba.NoVBAProjectError:
        return iter(())


def _signature_present(handle: Any, info: HostInfo) -> bool | None:
    """Whether the project carries a digital signature, or None if not detectable.

    Office keeps the signature of a .xlsm, .docm or .pptm in parts beside
    vbaProject.bin rather than in its streams, and only `vba_signature` looks
    there; reading the streams alone called every such file unsigned.

    Access keeps its VBA in the database rather than a vbaProject.bin, so neither
    place is where its signature would live. None is the honest answer there, not
    False.
    """
    if info.host == "access":
        return None
    import pyopenvba

    try:
        signature = getattr(handle, "vba_signature", None)
        if callable(signature):
            return bool(signature().present)
        from pyopenvba.cfb import CFB
        from pyopenvba.vba import detect_signature

        data = handle.vba_project_bytes()
        if not data:
            return False
        return bool(detect_signature(CFB(data)).present)
    except pyopenvba.NoVBAProjectError:
        return False
    except Exception:
        return None


def is_document_module(kind: str) -> bool:
    return kind == _DOCUMENT_KIND
