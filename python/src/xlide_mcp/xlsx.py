"""Worksheet cells, straight out of the OOXML package.

pyOpenVBA covers the VBA project, the forms and the queries; the grid is the one
part of a workbook it does not reach, so it is read and written here. The approach
is XLIDE's, and the reason is the same: a write splices into the original sheet
XML and only the touched rows are re-serialized, so styles, conditional formatting,
charts, pivot caches, the VBA project and every other part of the package survive
byte for byte. Rewriting the workbook through a general-purpose library would
rebuild parts it does not model, and a user's workbook is the wrong place to find
out which ones those are.

The XML is scanned rather than parsed into a tree. That buys three things: the
byte offsets a splice needs, no entity expansion on a file that came from outside,
and no dependency.

Two limits are deliberate. `.xlsb` and `.xls` are not OOXML and are not read here;
a caller gets a refusal naming .xlsx and .xlsm rather than a parse error. And a
written value is a value: the result Excel shows for a formula that depends on it
does not change until Excel next opens the workbook and recalculates.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from .errors import ToolError

CellValue = str | float | bool | None

MAX_ROW = 1_048_576
MAX_COLUMN = 16_384

# A read wider than this is not a read an agent can use; it is a way to spend the
# whole context on one call. The cap is reported, with the range that would fit.
MAX_CELLS_PER_READ = 20_000

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_TYPE_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
_REL_TYPE_CALCCHAIN = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"


class XlsxError(ToolError):
    """A workbook's grid could not be read or written."""


# --------------------------------------------------------------- XML scanning


@dataclass(frozen=True)
class Tag:
    name: str
    attrs: dict[str, str]
    self_closing: bool
    start: int
    end: int

    @property
    def closing(self) -> bool:
        return self.name.startswith("/")


_ATTR_RE = re.compile(r"""([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')""")
_NAME_RE = re.compile(r"^/?([\w:.-]+)")


def next_tag(xml: str, start: int) -> Tag | None:
    """The next element tag at or after `start`, honouring quoted attributes.

    Declarations and comments are skipped, so a `>` inside a comment does not end
    a tag that has not started.
    """
    i = xml.find("<", start)
    while i >= 0 and i + 1 < len(xml) and xml[i + 1] in "?!":
        if xml.startswith("<!--", i):
            close = xml.find("-->", i + 4)
            i = xml.find("<", i + 4 if close < 0 else close + 3)
        else:
            close = xml.find(">", i)
            i = xml.find("<", i + 1 if close < 0 else close + 1)
    if i < 0:
        return None
    j = i + 1
    quote = ""
    while j < len(xml):
        ch = xml[j]
        if quote:
            if ch == quote:
                quote = ""
        elif ch in "\"'":
            quote = ch
        elif ch == ">":
            break
        j += 1
    inner = xml[i + 1 : j]
    self_closing = inner.endswith("/")
    body = inner[:-1] if self_closing else inner
    match = _NAME_RE.match(body)
    name = match.group(1) if match else ""
    attrs = {
        m.group(1): decode_xml(m.group(3) if m.group(3) is not None else m.group(4) or "")
        for m in _ATTR_RE.finditer(body)
    }
    full_name = f"/{name}" if body.startswith("/") else name
    return Tag(name=full_name, attrs=attrs, self_closing=self_closing, start=i, end=j + 1)


_ENTITY_RE = re.compile(r"&#x([0-9a-fA-F]+);|&#(\d+);")


def decode_xml(text: str) -> str:
    # [XML] 2.11 normalizes line endings before entities expand, so a literal CRLF
    # becomes LF while an explicit &#13; survives as CR.
    normalized = text.replace("\r\n", "\n").replace("\r", "\n") if "\r" in text else text
    if "&" not in normalized:
        return normalized
    normalized = _ENTITY_RE.sub(
        lambda m: chr(int(m.group(1), 16)) if m.group(1) else chr(int(m.group(2))), normalized
    )
    return (
        normalized.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def encode_xml(text: str) -> str:
    """Escape for element text. Control characters are illegal in XML 1.0."""
    escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return _ILLEGAL_XML.sub("", escaped)


def encode_attr(text: str) -> str:
    return encode_xml(text).replace('"', "&quot;")


# ------------------------------------------------------------- A1 conversions


def column_to_index(letters: str) -> int:
    """A -> 1. Zero for anything that is not a column."""
    index = 0
    for ch in letters.upper():
        if not ("A" <= ch <= "Z"):
            return 0
        index = index * 26 + (ord(ch) - 64)
        if index > MAX_COLUMN:
            return 0
    return index


def index_to_column(index: int) -> str:
    """1 -> A."""
    if index < 1:
        raise XlsxError(f"Column index out of range: {index}.")
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


_CELL_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")


def parse_cell_ref(ref: str) -> tuple[int, int]:
    """'B3' -> (row 3, column 2)."""
    match = _CELL_REF_RE.match(ref.strip())
    if not match:
        raise XlsxError(f"Not a cell reference: {ref!r}. Use A1 style, such as B3.")
    row = int(match.group(2))
    column = column_to_index(match.group(1))
    if not column or not 1 <= row <= MAX_ROW:
        raise XlsxError(f"Cell reference out of range: {ref!r}.")
    return row, column


@dataclass(frozen=True)
class CellRange:
    first_row: int
    first_column: int
    last_row: int
    last_column: int

    @property
    def cell_count(self) -> int:
        return (self.last_row - self.first_row + 1) * (self.last_column - self.first_column + 1)

    def __str__(self) -> str:
        start = f"{index_to_column(self.first_column)}{self.first_row}"
        end = f"{index_to_column(self.last_column)}{self.last_row}"
        return start if start == end else f"{start}:{end}"


def parse_range(ref: str) -> CellRange:
    """'A1:D50', or a single cell. Reversed corners are normalized."""
    text = (ref or "").strip().replace("$", "")
    if "!" in text:
        text = text.rsplit("!", 1)[1]
    if not text:
        raise XlsxError("No range given. Use A1 style, such as A1:D50.")
    if ":" not in text:
        row, column = parse_cell_ref(text)
        return CellRange(row, column, row, column)
    left, right = text.split(":", 1)
    first_row, first_column = parse_cell_ref(left)
    last_row, last_column = parse_cell_ref(right)
    return CellRange(
        min(first_row, last_row),
        min(first_column, last_column),
        max(first_row, last_row),
        max(first_column, last_column),
    )


# ------------------------------------------------------------ future functions

# Functions newer than the ECMA-376 predefined set are stored with a prefix the
# formula bar hides. Ported from XLIDE's xlsxFunctionNames.ts, which checked each
# name against what Excel 16 actually stores.
_WORKSHEET_ONLY = frozenset({"FILTER", "PY", "SORT"})
_FUTURE_NAMES = """
ACOT ACOTH AGGREGATE ANCHORARRAY ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST
BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH
CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS
COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COPILOT COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS
DECIMAL DETECTLANGUAGE DROP ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT
F.INV F.INV.RT F.TEST FIELDVALUE FILTER FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS
FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA
GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS GROUPBY HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT
IMCSC IMCSCH IMPORTCSV IMPORTTEXT IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOMITTED ISOWEEKNUM LAMBDA
LET LOGNORM.DIST LOGNORM.INV LONGTEXT MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT
NEGBINOM.DIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC
PERCENTILE.INC PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI PIVOTBY POISSON.DIST
PQSOURCE PY PYTHON_STR PYTHON_TYPE PYTHON_TYPENAME QUARTILE.EXC QUARTILE.INC QUERYSTRING RANDARRAY
RANK.AVG RANK.EQ REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE SHEET SHEETS
SINGLE SKEW.P SORT SORTBY STDEV.P STDEV.S STOCKHISTORY SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV
T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW TRANSLATE TRIMRANGE UNICHAR
UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST WRAPCOLS WRAPROWS XLOOKUP
XMATCH XOR Z.TEST
"""
FUTURE_FUNCTIONS: dict[str, str] = {
    name: ("_xlfn._xlws." if name in _WORKSHEET_ONLY else "_xlfn.") + name
    for name in _FUTURE_NAMES.split()
}
_STORED_TO_TYPED = {stored: typed for typed, stored in FUTURE_FUNCTIONS.items()}

_FUNCTION_CALL_RE = re.compile(r"(_xlfn\._xlws\.|_xlfn\.)?([A-Za-z][A-Za-z0-9_.]*)\s*\(")


def formula_for_file(formula: str) -> str:
    """A formula as typed, in the form the file stores: future functions prefixed."""
    def replace(match: re.Match[str]) -> str:
        name = match.group(2)
        stored = FUTURE_FUNCTIONS.get(name.upper())
        # An unrecognized name keeps whatever prefix it arrived with rather than
        # losing it: the table is what Excel 16 was seen storing, not everything.
        prefix = stored[: -len(name)] if stored else (match.group(1) or "")
        return f"{prefix}{name}{match.group(0)[match.end(2) - match.start(0) :]}"

    return _outside_strings(formula, replace)


def formula_for_display(formula: str) -> str:
    """A stored formula as the formula bar shows it: prefixes removed."""
    def replace(match: re.Match[str]) -> str:
        if not match.group(1):
            return match.group(0)
        typed = _STORED_TO_TYPED.get(match.group(1) + match.group(2).upper())
        if typed is None:
            return match.group(0)
        return match.group(2) + match.group(0)[match.end(2) - match.start(0) :]

    return _outside_strings(formula, replace)


def _outside_strings(formula: str, replace: Any) -> str:
    """Apply a function-name substitution everywhere except inside "..." literals."""
    out: list[str] = []
    index = 0
    length = len(formula)
    while index < length:
        ch = formula[index]
        if ch == '"':
            end = index + 1
            while end < length:
                if formula[end] == '"':
                    if end + 1 < length and formula[end + 1] == '"':
                        end += 2
                        continue
                    end += 1
                    break
                end += 1
            out.append(formula[index:end])
            index = end
            continue
        if ch == "'":
            end = formula.find("'", index + 1)
            end = length if end < 0 else end + 1
            out.append(formula[index:end])
            index = end
            continue
        match = _FUNCTION_CALL_RE.match(formula, index)
        if match:
            out.append(replace(match))
            index = match.end()
            continue
        out.append(ch)
        index += 1
    return "".join(out)


# A reference, not preceded by a name character and not followed by one or by the
# `(` that would make it a function call.
_REFERENCE_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_.(])")


def shift_formula(formula: str, row_delta: int, column_delta: int) -> str:
    """Move a formula's relative references, the way filling a shared formula does."""
    if not row_delta and not column_delta:
        return formula

    def move(match: re.Match[str]) -> str:
        col_anchor, letters, row_anchor, digits = match.groups()
        column = column_to_index(letters)
        row = int(digits)
        if not column or row < 1:
            return match.group(0)
        if not col_anchor:
            column += column_delta
        if not row_anchor:
            row += row_delta
        if not 1 <= column <= MAX_COLUMN or not 1 <= row <= MAX_ROW:
            return "#REF!"
        return f"{col_anchor}{index_to_column(column)}{row_anchor}{row}"

    out: list[str] = []
    index = 0
    length = len(formula)
    while index < length:
        ch = formula[index]
        if ch in "\"'":
            quote = ch
            end = index + 1
            while end < length and formula[end] != quote:
                end += 1
            end = min(end + 1, length)
            out.append(formula[index:end])
            index = end
            continue
        match = _REFERENCE_RE.match(formula, index)
        # A reference cannot start immediately after a name character: the `LOG10`
        # in LOG10(A1) parses as column LOG row 10 without this guard.
        if match and (index == 0 or formula[index - 1] not in "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                      "abcdefghijklmnopqrstuvwxyz0123456789_."):
            out.append(move(match))
            index = match.end()
            continue
        out.append(ch)
        index += 1
    return "".join(out)


# ------------------------------------------------------------------- the book


@dataclass(frozen=True)
class SheetSummary:
    name: str
    used_range: str
    hidden: bool
    part: str


@dataclass(frozen=True)
class NamedRange:
    name: str
    refers_to: str


@dataclass
class Cell:
    ref: str
    row: int
    column: int
    value: CellValue = None
    formula: str | None = None
    kind: Literal["number", "text", "boolean", "error", "date", "empty"] = "empty"


@dataclass
class _SharedFormula:
    text: str
    row: int
    column: int


class Workbook:
    """An OOXML workbook's grid. Reads on demand; writes splice and save in place."""

    def __init__(self, path: Path) -> None:
        self.path = path
        suffix = path.suffix.lower()
        if suffix not in {".xlsx", ".xlsm", ".xlam"}:
            raise XlsxError(
                f"Worksheet cells are read from the OOXML package, which {suffix} is not. "
                "Supported for sheets: .xlsx, .xlsm and .xlam. A .xlsb or .xls holds its grid "
                "in a binary format this server does not read; its VBA is still fully readable."
            )
        self._parts: dict[str, bytes] = {}
        self._order: list[str] = []
        try:
            with zipfile.ZipFile(path) as archive:
                for info in archive.infolist():
                    if info.is_dir():
                        continue
                    self._order.append(info.filename)
                    self._parts[info.filename] = archive.read(info)
        except zipfile.BadZipFile as exc:
            raise XlsxError(f"{path.name} is not a readable Office package: {exc}.") from exc
        except OSError as exc:
            raise XlsxError(f"{path.name} could not be opened: {exc}.") from exc
        self._shared_strings: list[str] | None = None
        self._sheets: list[SheetSummary] | None = None

    # -- parts ------------------------------------------------------------

    def _text(self, part: str) -> str:
        data = self._parts.get(part)
        if data is None:
            raise XlsxError(f"{self.path.name} has no part {part}.")
        return data.decode("utf-8", errors="replace")

    def _optional_text(self, part: str) -> str | None:
        data = self._parts.get(part)
        return None if data is None else data.decode("utf-8", errors="replace")

    # -- workbook ---------------------------------------------------------

    def sheets(self) -> list[SheetSummary]:
        if self._sheets is None:
            self._sheets = self._read_sheets()
        return self._sheets

    def sheet_names(self) -> list[str]:
        return [s.name for s in self.sheets()]

    def _read_sheets(self) -> list[SheetSummary]:
        rels = self._relationships("xl/_rels/workbook.xml.rels")
        xml = self._text("xl/workbook.xml")
        summaries: list[SheetSummary] = []
        position = 0
        while (tag := next_tag(xml, position)) is not None:
            position = tag.end
            if tag.name != "sheet":
                continue
            name = tag.attrs.get("name", "")
            rid = tag.attrs.get("r:id") or tag.attrs.get("id") or ""
            target = rels.get(rid)
            if not name or not target:
                continue
            part = _resolve_part("xl/workbook.xml", target)
            if part not in self._parts:
                continue
            summaries.append(
                SheetSummary(
                    name=name,
                    used_range=self._dimension(part),
                    hidden=tag.attrs.get("state", "visible") != "visible",
                    part=part,
                )
            )
        if not summaries:
            raise XlsxError(f"{self.path.name} lists no worksheets.")
        return summaries

    def named_ranges(self) -> list[NamedRange]:
        xml = self._optional_text("xl/workbook.xml") or ""
        names: list[NamedRange] = []
        position = 0
        while (tag := next_tag(xml, position)) is not None:
            position = tag.end
            if tag.name != "definedName" or tag.self_closing:
                continue
            close = xml.find("</definedName>", tag.end)
            if close < 0:
                continue
            names.append(
                NamedRange(
                    name=tag.attrs.get("name", ""),
                    refers_to=decode_xml(xml[tag.end : close]).strip(),
                )
            )
            position = close
        return [n for n in names if n.name and not n.name.startswith("_xlnm.")]

    def _relationships(self, part: str) -> dict[str, str]:
        xml = self._optional_text(part)
        if xml is None:
            return {}
        out: dict[str, str] = {}
        position = 0
        while (tag := next_tag(xml, position)) is not None:
            position = tag.end
            if tag.name == "Relationship":
                out[tag.attrs.get("Id", "")] = tag.attrs.get("Target", "")
        return out

    def canonical_sheet_name(self, name: str) -> str:
        """The sheet's own spelling of a name the caller matched without case."""
        return self._sheet(name).name

    # Parts of the package other than the grid: the drawing layer reads through
    # these rather than opening the archive a second time.

    def part_text(self, part: str) -> str:
        """One part's text. Raises if the package does not hold it."""
        return self._text(part)

    def optional_part_text(self, part: str) -> str | None:
        """One part's text, or None where the package does not hold it."""
        return self._optional_text(part)

    def part_relationships(self, part: str) -> dict[str, dict[str, str]]:
        """A part's relationships by id, each resolved to a package part and type."""
        rels_part = _relationships_part(part)
        xml = self._optional_text(rels_part)
        if xml is None:
            return {}
        out: dict[str, dict[str, str]] = {}
        position = 0
        while (tag := next_tag(xml, position)) is not None:
            position = tag.end
            if tag.name != "Relationship":
                continue
            identifier = tag.attrs.get("Id", "")
            target = tag.attrs.get("Target", "")
            if not identifier or not target:
                continue
            out[identifier] = {
                "part": _resolve_part(part, target),
                "type": tag.attrs.get("Type", ""),
            }
        return out

    def _sheet(self, name: str) -> SheetSummary:
        wanted = (name or "").strip().casefold()
        for sheet in self.sheets():
            if sheet.name.casefold() == wanted:
                return sheet
        listed = ", ".join(s.name for s in self.sheets())
        raise XlsxError(f"No sheet named {name!r}. Sheets in this workbook: {listed}.")

    def _dimension(self, part: str) -> str:
        xml = self._optional_text(part)
        if xml is None:
            return ""
        position = 0
        while (tag := next_tag(xml, position)) is not None:
            position = tag.end
            if tag.name == "dimension":
                return tag.attrs.get("ref", "")
            if tag.name in {"sheetData", "/worksheet"}:
                break
        return ""

    # -- shared strings ---------------------------------------------------

    def _string(self, index: int) -> str:
        if self._shared_strings is None:
            self._shared_strings = self._read_shared_strings()
        if 0 <= index < len(self._shared_strings):
            return self._shared_strings[index]
        return ""

    def _read_shared_strings(self) -> list[str]:
        xml = self._optional_text("xl/sharedStrings.xml")
        if xml is None:
            return []
        strings: list[str] = []
        position = 0
        depth_start = -1
        while (tag := next_tag(xml, position)) is not None:
            position = tag.end
            if tag.name == "si" and not tag.self_closing:
                depth_start = tag.end
            elif tag.name == "/si" and depth_start >= 0:
                strings.append(_concat_text_runs(xml[depth_start : tag.start]))
                depth_start = -1
            elif tag.name == "si" and tag.self_closing:
                strings.append("")
        return strings

    # -- reading ----------------------------------------------------------

    def read(self, sheet_name: str, ref: str) -> tuple[CellRange, list[list[Cell]]]:
        """Every cell in a range, as a dense grid. Missing cells come back empty."""
        sheet = self._sheet(sheet_name)
        wanted = parse_range(ref)
        if wanted.cell_count > MAX_CELLS_PER_READ:
            raise XlsxError(
                f"{wanted} is {wanted.cell_count:,} cells, over the {MAX_CELLS_PER_READ:,} "
                "a single read returns. Read it in blocks."
            )
        found = self._scan_cells(sheet, wanted)
        grid: list[list[Cell]] = []
        for row in range(wanted.first_row, wanted.last_row + 1):
            line: list[Cell] = []
            for column in range(wanted.first_column, wanted.last_column + 1):
                ref_text = f"{index_to_column(column)}{row}"
                line.append(found.get((row, column)) or Cell(ref_text, row, column))
            grid.append(line)
        return wanted, grid

    def _scan_cells(self, sheet: SheetSummary, wanted: CellRange) -> dict[tuple[int, int], Cell]:
        xml = self._text(sheet.part)
        out: dict[tuple[int, int], Cell] = {}
        shared: dict[str, _SharedFormula] = {}
        position = 0
        row_number = 0
        while (tag := next_tag(xml, position)) is not None:
            position = tag.end
            if tag.name == "row":
                row_number = int(tag.attrs.get("r", "0") or 0)
                # Rows are written in order, and a shared formula's master always
                # precedes its slaves, so past the range there is nothing left.
                if row_number > wanted.last_row:
                    break
                continue
            if tag.name != "c" or tag.closing:
                continue
            ref = tag.attrs.get("r", "")
            if ref:
                row, column = parse_cell_ref(ref)
            else:
                row, column = row_number, 0
            body = ""
            if not tag.self_closing:
                close = xml.find("</c>", tag.end)
                if close < 0:
                    break
                body = xml[tag.end : close]
                position = close + 4
            formula = self._cell_formula(body, row, column, shared)
            inside = (
                wanted.first_row <= row <= wanted.last_row
                and wanted.first_column <= column <= wanted.last_column
            )
            if not inside:
                continue
            value, kind = self._cell_value(tag.attrs.get("t", "n"), body)
            out[(row, column)] = Cell(
                ref=ref or f"{index_to_column(column)}{row}",
                row=row,
                column=column,
                value=value,
                formula=formula,
                kind=kind,
            )
        return out

    def _cell_formula(
        self, body: str, row: int, column: int, shared: dict[str, _SharedFormula]
    ) -> str | None:
        tag = next_tag(body, 0)
        while tag is not None and tag.name != "f":
            tag = next_tag(body, tag.end)
        if tag is None:
            return None
        text = ""
        if not tag.self_closing:
            close = body.find("</f>", tag.end)
            if close >= 0:
                text = decode_xml(body[tag.end : close])
        index = tag.attrs.get("si")
        if tag.attrs.get("t") == "shared" and index is not None:
            if text:
                shared[index] = _SharedFormula(text=text, row=row, column=column)
            else:
                master = shared.get(index)
                if master is None:
                    return None
                text = shift_formula(master.text, row - master.row, column - master.column)
        if not text:
            return None
        return "=" + formula_for_display(text)

    def _cell_value(self, cell_type: str, body: str) -> tuple[CellValue, str]:
        if cell_type == "inlineStr":
            return _concat_text_runs(body), "text"
        raw = _first_element_text(body, "v")
        if raw is None:
            return None, "empty"
        if cell_type == "s":
            try:
                return self._string(int(raw)), "text"
            except ValueError:
                return "", "text"
        if cell_type == "str":
            return raw, "text"
        if cell_type == "b":
            return raw.strip() not in {"0", "", "false", "FALSE"}, "boolean"
        if cell_type == "e":
            return raw, "error"
        if cell_type == "d":
            return raw, "date"
        try:
            number = float(raw)
        except ValueError:
            return raw, "text"
        return (int(number) if number.is_integer() and abs(number) < 2**53 else number), "number"

    # -- writing ----------------------------------------------------------

    def write(
        self,
        sheet_name: str,
        start_cell: str,
        data: list[list[Any]],
    ) -> tuple[CellRange, int]:
        """Put values and formulas into a block, splicing the sheet's own XML.

        A value written over a formula removes the formula, which is what typing
        into the cell does. Returns the range written and the cell count.
        """
        if not data or not any(isinstance(row, list) for row in data):
            raise XlsxError("data must be a list of rows, each row a list of cell values.")
        sheet = self._sheet(sheet_name)
        first_row, first_column = parse_cell_ref(start_cell)
        width = max(len(row) for row in data)
        last_row = first_row + len(data) - 1
        last_column = first_column + width - 1
        if last_row > MAX_ROW or last_column > MAX_COLUMN:
            raise XlsxError(
                f"The block starting at {start_cell} runs past the end of the sheet."
            )
        written = CellRange(first_row, first_column, last_row, last_column)

        updates: dict[tuple[int, int], str] = {}
        touched_formula = False
        for row_offset, row_values in enumerate(data):
            for column_offset, raw in enumerate(row_values):
                row = first_row + row_offset
                column = first_column + column_offset
                xml, is_formula = _cell_xml(row, column, raw)
                updates[(row, column)] = xml
                touched_formula = touched_formula or is_formula

        original = self._text(sheet.part)
        spliced, replaced_formula = _splice_cells(original, updates)
        self._parts[sheet.part] = spliced.encode("utf-8")
        self._parts["xl/workbook.xml"] = _set_full_calc_on_load(
            self._text("xl/workbook.xml")
        ).encode("utf-8")
        if touched_formula or replaced_formula:
            # The calculation chain names the cells holding formulas and their
            # order. Leaving a stale one behind is a repair prompt on the next
            # open; Excel rebuilds it from the formulas themselves.
            self._drop_calc_chain()
        return written, len(updates)

    def _drop_calc_chain(self) -> None:
        part = "xl/calcChain.xml"
        if part not in self._parts:
            return
        del self._parts[part]
        self._order = [name for name in self._order if name != part]
        content_types = self._optional_text("[Content_Types].xml")
        if content_types is not None:
            self._parts["[Content_Types].xml"] = _drop_override(
                content_types, "/xl/calcChain.xml"
            ).encode("utf-8")
        rels_part = "xl/_rels/workbook.xml.rels"
        rels = self._optional_text(rels_part)
        if rels is not None:
            self._parts[rels_part] = _drop_relationship(rels, _REL_TYPE_CALCCHAIN).encode("utf-8")

    def save(self, destination: Path | None = None) -> Path:
        """Rewrite the package. Untouched parts go back with their original bytes.

        The write lands on a temporary file beside the target and is moved into
        place, so an interrupted save leaves the original workbook intact rather
        than a truncated one.
        """
        target = destination or self.path
        temporary = target.with_name(target.name + ".xlide-tmp")
        try:
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as archive:
                for name in self._order:
                    data = self._parts.get(name)
                    if data is not None:
                        archive.writestr(name, data)
            shutil.move(str(temporary), str(target))
        except OSError as exc:
            temporary.unlink(missing_ok=True)
            if isinstance(exc, PermissionError):
                raise XlsxError(
                    f"{target.name} could not be written: it is locked, most likely open in "
                    "Excel. Ask the user to close it."
                ) from exc
            raise XlsxError(f"{target.name} could not be written: {exc}.") from exc
        return target


# --------------------------------------------------------------- XML surgery


def _cell_xml(row: int, column: int, raw: Any) -> tuple[str, bool]:
    """One `<c>` element, and whether it holds a formula.

    Text is written inline rather than added to the shared string table. Excel
    reads both, and an inline string touches one part instead of two, so a write
    cannot leave the table and the sheet disagreeing about an index.
    """
    ref = f"{index_to_column(column)}{row}"
    if raw is None or raw == "":
        return f'<c r="{ref}"/>', False
    if isinstance(raw, bool):
        return f'<c r="{ref}" t="b"><v>{1 if raw else 0}</v></c>', False
    if isinstance(raw, (int, float)):
        return f'<c r="{ref}"><v>{_number_text(raw)}</v></c>', False
    text = str(raw)
    if text.startswith("="):
        stored = encode_xml(formula_for_file(text[1:]))
        return f'<c r="{ref}"><f>{stored}</f></c>', True
    inline = f'<is><t xml:space="preserve">{encode_xml(text)}</t></is>'
    return f'<c r="{ref}" t="inlineStr">{inline}</c>', False


def _number_text(value: float | int) -> str:
    if isinstance(value, int):
        return str(value)
    if value.is_integer() and abs(value) < 2**53:
        return str(int(value))
    return repr(value)


def _splice_cells(xml: str, updates: dict[tuple[int, int], str]) -> tuple[str, bool]:
    """Replace or insert each cell, rebuilding only the rows that change.

    Returns the new XML and whether any replaced cell held a formula, which is
    what decides the calculation chain has to go.
    """
    body_start, body_end, self_closing = _sheet_data_span(xml)
    rows_by_number = _parse_rows(xml, body_start, body_end)
    replaced_formula = False

    by_row: dict[int, dict[int, str]] = {}
    for (row, column), cell_xml in updates.items():
        by_row.setdefault(row, {})[column] = cell_xml

    for row_number, cells in by_row.items():
        existing = rows_by_number.get(row_number)
        merged: dict[int, str] = {}
        attrs = f' r="{row_number}"'
        if existing is not None:
            attrs = existing.attributes
            for column, cell_xml in existing.cells.items():
                merged[column] = cell_xml
            for column in cells:
                previous = existing.cells.get(column)
                if previous is not None and "<f" in previous:
                    replaced_formula = True
        merged.update(cells)
        ordered = "".join(merged[column] for column in sorted(merged))
        # The spans attribute claims which columns the row holds. Excel repairs a
        # wrong one; dropping it lets Excel derive it, which is always right.
        attrs = re.sub(r'\s+spans="[^"]*"', "", attrs)
        rows_by_number[row_number] = _Row(
            number=row_number,
            attributes=attrs,
            cells=merged,
            xml=f"<row{attrs}>{ordered}</row>",
        )

    rebuilt = "".join(rows_by_number[number].xml for number in sorted(rows_by_number))
    if self_closing:
        return xml[:body_start] + f"<sheetData>{rebuilt}</sheetData>" + xml[body_end:], False
    return xml[:body_start] + rebuilt + xml[body_end:], replaced_formula


@dataclass
class _Row:
    number: int
    attributes: str
    cells: dict[int, str]
    xml: str = ""


def _sheet_data_span(xml: str) -> tuple[int, int, bool]:
    """Offsets of the sheetData content, and whether the element was empty."""
    position = 0
    while (tag := next_tag(xml, position)) is not None:
        position = tag.end
        if tag.name != "sheetData":
            continue
        if tag.self_closing:
            return tag.start, tag.end, True
        close = xml.find("</sheetData>", tag.end)
        if close < 0:
            break
        return tag.end, close, False
    raise XlsxError("The worksheet part has no sheetData element.")


def _parse_rows(xml: str, start: int, end: int) -> dict[int, _Row]:
    rows: dict[int, _Row] = {}
    position = start
    while position < end and (tag := next_tag(xml, position)) is not None:
        if tag.start >= end:
            break
        position = tag.end
        if tag.name != "row" or tag.closing:
            continue
        number = int(tag.attrs.get("r", "0") or 0)
        attributes = xml[tag.start + len("<row") : tag.end - (2 if tag.self_closing else 1)]
        if tag.self_closing:
            rows[number] = _Row(number, attributes, {}, xml[tag.start : tag.end])
            continue
        close = xml.find("</row>", tag.end)
        if close < 0 or close > end:
            break
        rows[number] = _Row(
            number=number,
            attributes=attributes,
            cells=_parse_row_cells(xml, tag.end, close, number),
            xml=xml[tag.start : close + len("</row>")],
        )
        position = close + len("</row>")
    return rows


def _parse_row_cells(xml: str, start: int, end: int, row_number: int) -> dict[int, str]:
    cells: dict[int, str] = {}
    position = start
    while position < end and (tag := next_tag(xml, position)) is not None:
        if tag.start >= end:
            break
        position = tag.end
        if tag.name != "c" or tag.closing:
            continue
        ref = tag.attrs.get("r", "")
        try:
            _, column = parse_cell_ref(ref) if ref else (row_number, 0)
        except XlsxError:
            continue
        if not column:
            continue
        if tag.self_closing:
            cells[column] = xml[tag.start : tag.end]
            continue
        close = xml.find("</c>", tag.end)
        if close < 0 or close > end:
            break
        cells[column] = xml[tag.start : close + len("</c>")]
        position = close + len("</c>")
    return cells


def _set_full_calc_on_load(xml: str) -> str:
    """Ask Excel to recalculate on the next open.

    A value written here changes nothing about the cached results of the formulas
    that depend on it. Without this the workbook opens showing stale numbers that
    look current, which is the failure worth preventing.
    """
    position = 0
    while (tag := next_tag(xml, position)) is not None:
        position = tag.end
        if tag.name != "calcPr":
            continue
        body = xml[tag.start : tag.end]
        if "fullCalcOnLoad" in body:
            return re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', xml, count=1)
        suffix = "/>" if tag.self_closing else ">"
        return xml[: tag.end - len(suffix)] + ' fullCalcOnLoad="1"' + suffix + xml[tag.end :]
    close = xml.rfind("</workbook>")
    if close < 0:
        return xml
    return xml[:close] + '<calcPr calcId="0" fullCalcOnLoad="1"/>' + xml[close:]


def _drop_override(content_types: str, part_name: str) -> str:
    position = 0
    while (tag := next_tag(content_types, position)) is not None:
        position = tag.end
        if tag.name == "Override" and tag.attrs.get("PartName") == part_name:
            return content_types[: tag.start] + content_types[tag.end :]
    return content_types


def _drop_relationship(rels: str, relationship_type: str) -> str:
    position = 0
    while (tag := next_tag(rels, position)) is not None:
        position = tag.end
        if tag.name == "Relationship" and tag.attrs.get("Type") == relationship_type:
            return rels[: tag.start] + rels[tag.end :]
    return rels


def _relationships_part(part: str) -> str:
    """Where a part's relationships live: _rels/<name>.rels beside it."""
    if "/" in part:
        directory, name = part.rsplit("/", 1)
        return f"{directory}/_rels/{name}.rels"
    return f"_rels/{part}.rels"


def _resolve_part(source_part: str, target: str) -> str:
    """A relationship target, as a package part name."""
    if target.startswith("/"):
        return target.lstrip("/")
    base = source_part.rsplit("/", 1)[0] if "/" in source_part else ""
    segments: list[str] = base.split("/") if base else []
    for segment in target.split("/"):
        if segment in {"", "."}:
            continue
        if segment == "..":
            if segments:
                segments.pop()
            continue
        segments.append(segment)
    return "/".join(segments)


def _concat_text_runs(fragment: str) -> str:
    """Every <t> in a fragment, joined. Rich text is many runs of one string."""
    out: list[str] = []
    position = 0
    in_phonetic = False
    while (tag := next_tag(fragment, position)) is not None:
        position = tag.end
        if tag.name in {"rPh", "phoneticPr"}:
            in_phonetic = not tag.self_closing
            continue
        if tag.name == "/rPh":
            in_phonetic = False
            continue
        if tag.name != "t" or tag.self_closing or in_phonetic:
            continue
        close = fragment.find("</t>", tag.end)
        if close < 0:
            break
        out.append(decode_xml(fragment[tag.end : close]))
        position = close + 4
    return "".join(out)


def _first_element_text(fragment: str, name: str) -> str | None:
    position = 0
    while (tag := next_tag(fragment, position)) is not None:
        position = tag.end
        if tag.name != name:
            continue
        if tag.self_closing:
            return ""
        close = fragment.find(f"</{name}>", tag.end)
        if close < 0:
            return ""
        return decode_xml(fragment[tag.end : close])
    return None


def sheets_summary(path: Path) -> list[SheetSummary]:
    """Sheet names and used ranges, for a one-shot project summary."""
    return Workbook(path).sheets()
