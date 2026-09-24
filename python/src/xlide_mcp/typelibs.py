"""The type libraries registered on this machine: what a VBA reference points at.

A reference is a GUID and a version, plus a path and a description that only
help the host find the file. An agent asked to add "the Scripting Runtime" would
otherwise have to recall its GUID, and a recalled GUID that is one digit out is a
reference Excel marks MISSING with a message naming nothing. The registry holds
the real ones: every library the VBA editor's References dialog lists is under
HKEY_CLASSES_ROOT\\TypeLib, as {guid}\\{major.minor} with the description as the
key's value and the file under 0\\win64 or 0\\win32.

The name a reference is recorded under is the library's own, the qualifier VBA
code writes (`Scripting`, not "Microsoft Scripting Runtime"). Only the type
library itself carries it, so it is read with pythoncom when that is installed,
and left for pyOpenVBA to default otherwise.

Windows only. Anywhere else a caller passes the GUID and version itself.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from typing import Any

_GUID = re.compile(r"^\{?[0-9A-Fa-f]{8}-(?:[0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}\}?$")
_VERSION = re.compile(r"^([0-9A-Fa-f]{1,4})\.([0-9A-Fa-f]{1,4})$")


@dataclass(frozen=True)
class RegisteredLibrary:
    guid: str
    """Upper case, with braces, as a libid spells it."""
    major: int
    minor: int
    description: str
    path: str

    @property
    def version(self) -> str:
        """As the registry and a libid spell it: hexadecimal, major.minor."""
        return f"{self.major:x}.{self.minor:x}"

    def summary(self) -> dict[str, Any]:
        return {
            "description": self.description,
            "guid": self.guid,
            "version": self.version,
            "path": self.path,
        }


def available() -> bool:
    return sys.platform == "win32"


def normal_guid(text: str) -> str | None:
    """A GUID with braces and in upper case, or None when it is not one."""
    candidate = (text or "").strip()
    if not _GUID.match(candidate):
        return None
    return "{" + candidate.strip("{}").upper() + "}"


def parse_version(text: str) -> tuple[int, int] | None:
    """`major.minor` in hexadecimal, as the registry and a libid spell it."""
    match = _VERSION.match((text or "").strip())
    if match is None:
        return None
    return int(match.group(1), 16), int(match.group(2), 16)


def registered() -> list[RegisteredLibrary]:
    """Every registered type library version, newest first within a GUID."""
    if not available():
        return []
    import winreg

    found: list[RegisteredLibrary] = []
    try:
        root = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "TypeLib")
    except OSError:
        return []
    with root:
        for guid_key in _subkeys(root):
            guid = normal_guid(guid_key)
            if guid is None:
                continue
            try:
                library = winreg.OpenKey(root, guid_key)
            except OSError:
                continue
            with library:
                for version_key in _subkeys(library):
                    version = parse_version(version_key)
                    if version is None:
                        continue
                    try:
                        with winreg.OpenKey(library, version_key) as entry:
                            description = _default_value(entry)
                            path = _library_file(entry)
                    except OSError:
                        continue
                    if not description:
                        continue
                    found.append(
                        RegisteredLibrary(
                            guid=guid,
                            major=version[0],
                            minor=version[1],
                            description=description,
                            path=path,
                        )
                    )
    found.sort(key=lambda lib: (lib.description.casefold(), lib.guid, -lib.major, -lib.minor))
    return found


def search(text: str, limit: int = 40) -> list[RegisteredLibrary]:
    """Libraries whose description contains the text, or whose GUID it is.

    One entry per GUID, its newest version: that is the one a new reference
    should point at, and listing every version buries the answer.
    """
    wanted = (text or "").strip().casefold()
    guid = normal_guid(text)
    if guid is not None:
        matching = [library for library in registered() if library.guid == guid]
    else:
        matching = [lib for lib in registered() if wanted in lib.description.casefold()]
    return sorted(_newest(matching), key=lambda lib: lib.description.casefold())[:limit]


def named(text: str) -> list[RegisteredLibrary]:
    """Every library whose own name is the text, newest version first.

    The name is the qualifier VBA code writes, ADODB or VBScript_RegExp_55, and
    only the library itself carries it, so each one is read: about a third of a
    second for the few hundred a machine registers.
    """
    wanted = (text or "").strip().casefold()
    if not wanted:
        return []
    found = [lib for lib in _newest(registered()) if library_name(lib).casefold() == wanted]
    return sorted(found, key=_version, reverse=True)


def resolve(text: str) -> RegisteredLibrary | list[RegisteredLibrary]:
    """The one library a description, a name or a GUID picks out, or the candidates.

    An exact description wins over a partial one, so "Microsoft Scripting Runtime"
    is not ambiguous merely because another description contains it. Otherwise the
    name code writes settles it: "Scripting" is inside "Microsoft WMI Scripting V1.2
    Library" too, and ADODB is in no description at all. Several libraries share a
    name where each version has its own GUID, as ADO's do, and then the newest is
    taken when one is newer than the rest.
    """
    wanted = text.strip().casefold()
    candidates = search(text, limit=1000)
    exact = [lib for lib in candidates if lib.description.casefold() == wanted]
    if len(exact) == 1:
        return exact[0]
    if len(candidates) == 1:
        return candidates[0]
    by_name = named(text)
    if by_name and (len(by_name) == 1 or _version(by_name[0]) > _version(by_name[1])):
        return by_name[0]
    return exact or candidates or by_name


def _version(library: RegisteredLibrary) -> tuple[int, int]:
    return library.major, library.minor


def _newest(libraries: list[RegisteredLibrary]) -> list[RegisteredLibrary]:
    """One entry per GUID, its newest version: the one a new reference should point at."""
    newest: dict[str, RegisteredLibrary] = {}
    for library in libraries:
        current = newest.get(library.guid)
        if current is None or _version(library) > _version(current):
            newest[library.guid] = library
    return list(newest.values())


def library_name(library: RegisteredLibrary) -> str:
    """The name the library gives itself, which VBA writes as the qualifier, or ''."""
    if not available():
        return ""
    try:
        import pythoncom
    except ImportError:
        return ""
    try:
        loaded = pythoncom.LoadRegTypeLib(library.guid, library.major, library.minor, 0)
        return str(loaded.GetDocumentation(-1)[0] or "")
    except Exception:
        return ""


def _subkeys(key: Any) -> list[str]:
    import winreg

    names: list[str] = []
    index = 0
    while True:
        try:
            names.append(winreg.EnumKey(key, index))
        except OSError:
            return names
        index += 1


def _default_value(key: Any) -> str:
    import winreg

    try:
        value, _kind = winreg.QueryValueEx(key, None)
    except OSError:
        return ""
    return str(value or "").strip()


def _library_file(version_key: Any) -> str:
    """The library's file, preferring the bitness Office on this machine most likely has."""
    import winreg

    for lcid in _subkeys(version_key):
        if not lcid.isdigit():
            continue
        for platform in ("win64", "win32"):
            try:
                with winreg.OpenKey(version_key, rf"{lcid}\{platform}") as entry:
                    value = _default_value(entry)
            except OSError:
                continue
            if value:
                return value
    return ""
