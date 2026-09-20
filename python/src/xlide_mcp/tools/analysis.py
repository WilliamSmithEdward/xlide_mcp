"""Static analysis: the build gate for VBA that has no build.

VBA compiles when the user presses F5, in an application that may not be on this
machine, so nothing otherwise stands between a written module and a run-time
failure in front of the user. pyvbaanalysis is the analyzer XLIDE validated against
the real compiler, ported; it reports a problem only when it can prove one, which
is what makes "treat any error as a build failure" a rule an agent can follow
without drowning in false alarms.

Diagnostics come back with line and column rather than the character offsets the
library returns, because a character offset into a module an agent is holding as
text is a number it has to convert before it can act on it.
"""

from __future__ import annotations

from typing import Annotated, Any

from mcp.server.mcpserver import MCPServer
from pydantic import Field

from .. import project as project_layer
from ..config import Settings
from ..errors import ToolError
from ..hosts import require_readable
from ..paths import resolve_path
from ._common import limited, read_only

# A file with thousands of findings is a file nobody is going to fix from one
# tool result. Report the count honestly, return the ones worth acting on.
MAX_DIAGNOSTICS = 300


def register(server: MCPServer, settings: Settings) -> None:
    @server.tool(
        name="xlide_analyze",
        title="Analyze VBA",
        annotations=read_only("Analyze VBA"),
        description=(
            "Runs static analysis over every VBA module in an Office file and returns the "
            "problems with module, line, column, code and message. Needs no Office "
            "installation and runs nothing. Each file is measured against its own host's "
            "object model, so Word code is never judged by Excel's surface. Call this after "
            "every VBA change and treat any problem at error severity as a build failure: fix "
            "it and analyze again until it is clean. Warnings are worth reading; some are "
            "style, some are the bug."
        ),
    )
    def analyze(
        file_path: Annotated[str, Field(description="Absolute path to the Office file.")],
        module_name: Annotated[
            str,
            Field(
                default="",
                description=(
                    "Report only this module's problems. The whole project is still analyzed, "
                    "so cross-module references still resolve."
                ),
            ),
        ] = "",
        min_severity: Annotated[
            str,
            Field(
                default="information",
                description="Lowest severity to report: 'error', 'warning' or 'information'.",
            ),
        ] = "information",
    ) -> dict[str, Any]:
        path = resolve_path(file_path, settings)
        info = require_readable(path)
        floor = _severity_floor(min_severity)

        from pyvbaanalysis import analyze_project

        with project_layer.open_project(path, info) as handle:
            modules = project_layer.read_modules(handle, info)
        prepared = project_layer.analysis_inputs(modules)
        by_name = {item.name.casefold(): item for item in prepared}
        bodies = {m.name.casefold(): m.body for m in modules}

        try:
            by_module = analyze_project(
                [item.module_input for item in prepared], host=info.host
            )
        except Exception as exc:
            raise ToolError(
                f"Analysis of {path.name} failed: {exc}. "
                "xlide_validate_project reports whether the container itself is damaged."
            ) from exc

        wanted = module_name.strip().casefold()
        findings: list[dict[str, Any]] = []
        counts = {"error": 0, "warning": 0, "information": 0}
        for name, diagnostics in by_module.items():
            prepared_module = by_name.get(name.casefold())
            for diagnostic in diagnostics:
                severity = diagnostic.severity.value
                counts[severity] = counts.get(severity, 0) + 1
                if _RANK[severity] < floor:
                    continue
                if wanted and name.casefold() != wanted:
                    continue
                findings.append(
                    _finding(
                        name,
                        diagnostic,
                        analyzed=(
                            prepared_module.analyzed_source if prepared_module else ""
                        ),
                        body=bodies.get(name.casefold(), ""),
                        line_offset=prepared_module.line_offset if prepared_module else 0,
                    )
                )

        findings.sort(key=lambda f: (-_RANK[f["severity"]], f["module"], f["line"]))
        shown, total = limited(findings, MAX_DIAGNOSTICS)
        result: dict[str, Any] = {
            "path": str(path),
            "host": info.host,
            "modules_analyzed": len(by_module),
            "counts": counts,
            "reported": len(shown),
            "problems": shown,
            "verdict": "clean" if counts["error"] == 0 else "errors found",
        }
        if total > len(shown):
            result["note"] = (
                f"{total - len(shown)} further problems not listed. Raise min_severity, or "
                "narrow to one module, to see the rest."
            )
        if counts["error"]:
            result["next_step"] = (
                "Fix every problem at error severity and analyze again. An error here is what "
                "the VBA compiler would reject or what would fail at run time."
            )
        return result

    @server.tool(
        name="xlide_analyze_source",
        title="Analyze VBA source",
        annotations=read_only("Analyze VBA source"),
        description=(
            "Runs static analysis over VBA source you are holding, before it is written to a "
            "file. Use it to check code you have just generated: it costs nothing, needs no "
            "file, and catches the compile errors that would otherwise surface in front of the "
            "user. Pass host so the code is measured against the right object model, and "
            "file_path instead if the code is destined for a file that already exists, which "
            "resolves calls into the rest of that project."
        ),
    )
    def analyze_source(
        source: Annotated[str, Field(description="The VBA source to check.")],
        module_name: Annotated[
            str, Field(default="Module1", description="Name to report problems against.")
        ] = "Module1",
        kind: Annotated[
            str,
            Field(
                default="standard",
                description="'standard', 'class', 'document' or 'userform'.",
            ),
        ] = "standard",
        host: Annotated[
            str,
            Field(
                default="",
                description=(
                    "'excel', 'word', 'powerpoint' or 'access'. Empty checks the language "
                    "alone, with no host object model."
                ),
            ),
        ] = "",
        file_path: Annotated[
            str,
            Field(
                default="",
                description=(
                    "An existing Office file this module belongs to. Its other modules are "
                    "analyzed alongside, so calls into them resolve. Overrides host."
                ),
            ),
        ] = "",
    ) -> dict[str, Any]:
        from pyvbaanalysis import ModuleInput, analyze_project

        module_kind = _module_kind(kind)
        inputs = [
            ModuleInput(module_name=module_name, module_kind=module_kind, source=source)
        ]
        resolved_host = host.strip().lower() or None

        if file_path.strip():
            path = resolve_path(file_path, settings)
            info = require_readable(path)
            resolved_host = info.host
            with project_layer.open_project(path, info) as handle:
                modules = project_layer.read_modules(handle, info)
            inputs.extend(
                item.module_input
                for item in project_layer.analysis_inputs(modules)
                if item.name.casefold() != module_name.casefold()
            )

        if resolved_host and resolved_host not in {"excel", "word", "powerpoint", "access"}:
            raise ToolError(
                f"{host!r} is not a host. Use 'excel', 'word', 'powerpoint' or 'access'."
            )

        by_module = analyze_project(inputs, host=resolved_host)
        diagnostics = by_module.get(module_name, [])
        # The caller's own source is analyzed exactly as given, so there is no
        # header to shift by here.
        findings = [
            _finding(module_name, d, analyzed=source, body=source) for d in diagnostics
        ]
        findings.sort(key=lambda f: (-_RANK[f["severity"]], f["line"]))
        errors = sum(1 for f in findings if f["severity"] == "error")
        return {
            "module": module_name,
            "kind": module_kind.value,
            "host": resolved_host or "(language only)",
            "counts": {
                "error": errors,
                "warning": sum(1 for f in findings if f["severity"] == "warning"),
                "information": sum(1 for f in findings if f["severity"] == "information"),
            },
            "problems": findings,
            "verdict": "clean" if errors == 0 else "errors found",
        }

    @server.tool(
        name="xlide_rules",
        title="Analyzer rules",
        annotations=read_only("Analyzer rules"),
        description=(
            "The analyzer's rule catalogue: every diagnostic code with its title, default "
            "severity, category, whether it mirrors a VBA compile failure, and the MS-VBAL "
            "section it enforces. Use it to explain a code to the user, or to decide whether "
            "a finding is a compile error or a judgement call."
        ),
    )
    def rules(
        code: Annotated[
            str, Field(default="", description="One rule code. Empty lists them all.")
        ] = "",
        search: Annotated[
            str, Field(default="", description="Only rules whose code or title contains this.")
        ] = "",
    ) -> dict[str, Any]:
        from pyvbaanalysis import rule_metadata_by_code

        catalogue = rule_metadata_by_code()
        if code.strip():
            meta = catalogue.get(code.strip())
            if meta is None:
                raise ToolError(
                    f"No rule with code {code!r}. Call this tool with no arguments to list them."
                )
            return {"rule": _rule(meta)}
        needle = search.strip().casefold()
        entries = [
            _rule(meta)
            for key, meta in sorted(catalogue.items())
            if not needle or needle in key.casefold() or needle in (meta.title or "").casefold()
        ]
        shown, total = limited(entries, 400)
        result: dict[str, Any] = {"count": total, "rules": shown}
        if total > len(shown):
            result["note"] = f"{total - len(shown)} more; narrow with search."
        return result


_RANK = {"error": 3, "warning": 2, "information": 1}


def _severity_floor(name: str) -> int:
    rank = _RANK.get((name or "").strip().lower())
    if rank is None:
        raise ToolError("min_severity must be 'error', 'warning' or 'information'.")
    return rank


def _module_kind(name: str) -> Any:
    from pyvbaanalysis import ModuleSymbolKind

    try:
        return ModuleSymbolKind((name or "standard").strip().lower())
    except ValueError as exc:
        raise ToolError(
            f"{name!r} is not a module kind. Use 'standard', 'class', 'document' or 'userform'."
        ) from exc


def _finding(
    module: str,
    diagnostic: Any,
    *,
    analyzed: str,
    body: str,
    line_offset: int = 0,
) -> dict[str, Any]:
    """One diagnostic, with its position moved into the coordinates a reader gets.

    The analyzer measures against the source it was given, header included.
    `xlide_read_module` hands back the body. Reporting the analyzer's line without
    subtracting the header would send an agent to edit a line that is not the one
    the problem is on.
    """
    from pyvbaanalysis import line_col

    line, column = (1, 1)
    in_header = False
    span = getattr(diagnostic, "span", None)
    if span is not None and analyzed:
        try:
            raw_line, column = line_col(analyzed, span.start)
        except Exception:
            raw_line, column = (1, 1)
        line = raw_line - line_offset
        if line < 1:
            in_header = True
            line = 1
    entry: dict[str, Any] = {
        "module": module,
        "line": line,
        "column": column,
        "severity": diagnostic.severity.value,
        "code": diagnostic.code,
        "message": diagnostic.message,
    }
    if in_header:
        entry["in_module_header"] = True
    text_source = body or analyzed
    if text_source and not in_header:
        lines = text_source.splitlines()
        if 1 <= line <= len(lines):
            entry["text"] = lines[line - 1].strip()[:200]
    if getattr(diagnostic, "spec_reference", None):
        entry["spec_reference"] = diagnostic.spec_reference
    return entry


def _rule(meta: Any) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "code": meta.code,
        "title": meta.title,
        "default_severity": meta.default_severity.value,
        "category": meta.category.value,
        "compile_error_equivalent": bool(meta.vbe_compile_equivalent),
    }
    if getattr(meta, "spec_reference", None):
        entry["spec_reference"] = meta.spec_reference
    if getattr(meta, "confidence", None):
        entry["confidence"] = meta.confidence
    return entry
