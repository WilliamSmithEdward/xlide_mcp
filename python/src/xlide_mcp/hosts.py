"""Which Office application a file belongs to, and what can be done with it.

The extension decides two separate things and they are not the same question:
which container reader opens the file, and which object model its VBA is measured
against. Word code judged by Excel's surface produces confident nonsense, so the
host travels with the file everywhere in this server.

The readable sets below are exactly what pyOpenVBA opens, not every extension the
applications use. A .xltm is an Excel template with a VBA project inside it, and
this server still refuses it, because pyOpenVBA would fall through to its legacy
CFB path and fail with a parse error that names nothing. A clear refusal that says
"save it as .xlsm" is worth more than a confusing one.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from .errors import ToolError

HostName = Literal["excel", "word", "powerpoint", "access"]

# Extensions this server opens, by host.
EXCEL_READABLE = frozenset({".xlsm", ".xlsb", ".xlam", ".xls"})
WORD_READABLE = frozenset({".docm", ".dotm", ".doc"})
POWERPOINT_READABLE = frozenset({".pptm", ".potm"})
ACCESS_READABLE = frozenset({".accdb", ".mdb"})
READABLE = EXCEL_READABLE | WORD_READABLE | POWERPOINT_READABLE | ACCESS_READABLE

# Power Query lives outside the VBA project, in a custom XML part, so a plain
# .xlsx carries queries and no macros at all.
POWER_QUERY_READABLE = frozenset({".xlsx", ".xlsm", ".xlsb", ".xlam"})

# Worksheet cells are read straight out of the OOXML package, which the binary
# .xlsb and the legacy .xls do not use.
SHEETS_READABLE = frozenset({".xlsx", ".xlsm", ".xlam"})

# Recognized as Office, deliberately not opened. Each one gets a reason, because
# "unsupported" on its own sends the caller looking for a bug that is not there.
NOT_READABLE: dict[str, str] = {
    ".xltm": "Excel template. Save a copy as .xlsm to work on its VBA.",
    ".xlt": "Legacy Excel template. Save a copy as .xlsm to work on its VBA.",
    ".xla": "Legacy Excel add-in. Save a copy as .xlam to work on its VBA.",
    ".dot": "Legacy Word template. Save a copy as .docm to work on its VBA.",
    ".ppsm": "PowerPoint macro-enabled show. Save a copy as .pptm to work on its VBA.",
    ".ppam": "PowerPoint add-in. Save a copy as .pptm to work on its VBA.",
    ".ppa": "Legacy PowerPoint add-in. Save a copy as .pptm to work on its VBA.",
    ".ppt": (
        "Legacy PowerPoint. Its VBA project sits in a compressed record this server "
        "does not open. Save a copy as .pptm."
    ),
    ".accda": "Access add-in. Save a copy as .accdb to work on its VBA.",
    ".mda": "Legacy Access add-in. Save a copy as .accdb to work on its VBA.",
}

# Extensions create_project can write, and the pyOpenVBA class that writes each.
CREATABLE = frozenset({".xlsm", ".xlsb", ".xlam", ".docm", ".pptm", ".accdb", ".xlsx"})

_HOST_BY_EXTENSION: dict[str, HostName] = {
    **dict.fromkeys(EXCEL_READABLE | {".xlsx", ".xltm", ".xlt", ".xla"}, "excel"),
    **dict.fromkeys(WORD_READABLE | {".dot"}, "word"),
    **dict.fromkeys(POWERPOINT_READABLE | {".ppt", ".ppsm", ".ppam", ".ppa"}, "powerpoint"),
    **dict.fromkeys(ACCESS_READABLE | {".accda", ".mda"}, "access"),
}

HOST_TITLE: dict[HostName, str] = {
    "excel": "Excel",
    "word": "Word",
    "powerpoint": "PowerPoint",
    "access": "Access",
}


@dataclass(frozen=True)
class HostInfo:
    """What a file's extension says about it, before the file is opened."""

    host: HostName
    extension: str
    readable: bool
    reason: str = ""

    @property
    def title(self) -> str:
        return HOST_TITLE[self.host]

    @property
    def supports_power_query(self) -> bool:
        return self.extension in POWER_QUERY_READABLE

    @property
    def supports_sheets(self) -> bool:
        return self.extension in SHEETS_READABLE

    @property
    def supports_forms(self) -> bool:
        # Access keeps forms and reports in the database; the other three keep
        # UserForms in the VBA project. Both read through forms().
        return self.readable

    @property
    def writes_on_edit(self) -> bool:
        """Access writes into the database as the edit is made, not at save."""
        return self.host == "access"


def host_info(path: str | Path) -> HostInfo:
    """Classify a path by extension. Never touches the file."""
    extension = Path(path).suffix.lower()
    host = _HOST_BY_EXTENSION.get(extension)
    if host is None:
        raise ToolError(
            f"{extension or Path(path).name!r} is not an Office file this server handles. "
            f"Readable: {', '.join(sorted(READABLE))}, plus .xlsx for Power Query and sheets."
        )
    if extension in READABLE:
        return HostInfo(host=host, extension=extension, readable=True)
    if extension == ".xlsx":
        # Not a defect and not a refusal to explain away: a .xlsx has no VBA
        # project by design, and its Power Query and cells are fully reachable.
        return HostInfo(
            host=host,
            extension=extension,
            readable=False,
            reason=(
                "A .xlsx has no VBA project by design. Its worksheet cells and Power Query "
                "are readable; save it as .xlsm in Excel to give it macros."
            ),
        )
    return HostInfo(
        host=host,
        extension=extension,
        readable=False,
        reason=NOT_READABLE.get(extension, "This server does not open this extension."),
    )


def require_readable(path: str | Path) -> HostInfo:
    """Classify, and refuse anything whose VBA project cannot be opened."""
    info = host_info(path)
    if not info.readable:
        reason = info.reason or "This server does not open this extension."
        raise ToolError(f"{Path(path).name}: {reason}")
    return info


def require_openable(path: str | Path) -> HostInfo:
    """Classify for a tool that opens the document in its application.

    Wider than `require_readable` on purpose. Excel opens a .xlsx perfectly well
    and will run VBA injected into the session; what a .xlsx cannot do is hold a
    macro of its own. A gate that conflated the two refused a run that works.
    """
    info = host_info(path)
    if info.readable or info.extension == ".xlsx":
        return info
    raise ToolError(
        f"{Path(path).name}: {info.reason or 'This server does not open this extension.'}"
    )


def container_class(info: HostInfo) -> Any:
    """The pyOpenVBA class that opens this host's files."""
    import pyopenvba

    return {
        "excel": pyopenvba.ExcelFile,
        "word": pyopenvba.WordFile,
        "powerpoint": pyopenvba.PowerPointFile,
        "access": pyopenvba.AccessDatabase,
    }[info.host]


def open_container(path: Path, info: HostInfo | None = None) -> Any:
    """Open a file's VBA project. The caller closes it, or uses `container`."""
    resolved = info or require_readable(path)
    import pyopenvba

    try:
        return container_class(resolved)(path)
    except pyopenvba.PyOpenVBAError as exc:
        raise ToolError(f"{path.name}: {exc}") from exc


class container:
    """`with container(path) as book:` - opens, and always closes.

    pyOpenVBA's own classes are context managers, so this adds one thing: the
    host classification travels with the handle, because almost every caller
    needs both and looking it up twice is how the two drift apart.
    """

    def __init__(self, path: Path, info: HostInfo | None = None) -> None:
        self.path = path
        self.info = info or require_readable(path)
        self._handle: Any = None

    def __enter__(self) -> Any:
        self._handle = open_container(self.path, self.info)
        return self._handle

    def __exit__(self, *exc_info: object) -> None:
        handle, self._handle = self._handle, None
        if handle is not None:
            handle.close()


def iter_office_files(root: Path, *, include_unreadable: bool = False) -> Iterator[Path]:
    """Every Office file under a root, skipping the directories nobody means.

    Sorted, so two calls on an unchanged tree answer identically: an agent that
    picks "the first workbook" should get the same one twice.
    """
    skip = {
        ".git", ".hg", ".svn", "node_modules", "__pycache__", ".venv", "venv",
        ".tox", ".mypy_cache", ".pytest_cache", ".ruff_cache", "dist", "build",
        ".vscode-test", ".idea",
    }
    wanted = READABLE | {".xlsx"}
    if include_unreadable:
        wanted = wanted | set(NOT_READABLE)
    for dirpath, dirnames, filenames in _walk_sorted(root):
        dirnames[:] = sorted(d for d in dirnames if d not in skip and not d.startswith("~$"))
        for name in sorted(filenames):
            # Office writes a lock file beside an open document; it is not a file
            # anyone wants listed, and opening one fails.
            if name.startswith("~$") or name.startswith("."):
                continue
            if Path(name).suffix.lower() in wanted:
                yield Path(dirpath) / name


def _walk_sorted(root: Path) -> Iterator[tuple[str, list[str], list[str]]]:
    import os

    yield from os.walk(root, topdown=True, onerror=None, followlinks=False)
