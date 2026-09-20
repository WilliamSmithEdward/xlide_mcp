"""Write the repo's conformance corpus: the behaviour every implementation owes.

The tool contract says what the arguments are. This says what the answers mean,
which is the part a port gets wrong. Each case is a short script of tool calls and
a set of assertions over the last one's result, in a vocabulary small enough that
a runner in any language is an afternoon's work.

The cases are defined here in Python rather than typed as JSON so the VBA source
in them keeps its CRLF line endings through one escaping layer instead of two.
The artifact is the JSON; the Python suite runs the JSON, not this file, so a case
that is wrong fails in Python before any port ever sees it.

    python tools/export_conformance.py           # write ../contract/conformance.json
    python tools/export_conformance.py --check   # fail if it is out of date

Placeholders inside arguments:
    ${fixture}          the path of the case's fixture file
    ${folder:name}      a scratch folder, created on first use
    ${step[N].path}     a value from an earlier step's result, 0-based

Assertions, all optional, over the last step's result:
    equals, not_equals, contains, not_contains, starts_with,
    at_least (a count, or a collection's length),
    type ("string" | "number" | "boolean" | "array" | "object" | "null")
A step may instead declare `error_contains`, which requires it to fail with that
text in the message.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[2]
CONFORMANCE_PATH = REPO_ROOT / "contract" / "conformance.json"

CRLF = "\r\n"

HELPERS = CRLF.join(
    [
        "Option Explicit",
        "",
        "Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long",
        "    AddNums = a + b",
        "End Function",
        "",
        "Public Sub Greet()",
        "    Dim who As String",
        '    who = "world"',
        '    Debug.Print "hello " & who',
        "End Sub",
        "",
    ]
)

ACCESS_MODULE = CRLF.join(
    [
        "Option Compare Database",
        "Option Explicit",
        "",
        "Public Function Twice(ByVal n As Long) As Long",
        "    Twice = n * 2",
        "End Function",
        "",
    ]
)

VB6_HELPERS = CRLF.join(
    [
        'Attribute VB_Name = "Helpers"',
        "Option Explicit",
        "",
        "Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long",
        "    AddNums = a + b",
        "End Function",
        "",
    ]
)

# A VB6 form keeps a designer block ahead of its attributes. It is what VB draws
# the form from, it is not code, and handing it to an agent as editable source
# invites an edit that stops the project loading.
VB6_FORM = CRLF.join(
    [
        "VERSION 5.00",
        "Begin VB.Form Form1 ",
        '   Caption         =   "Demo"',
        "End",
        'Attribute VB_Name = "Form1"',
        "Option Explicit",
        "",
        "Private Sub Form_Load()",
        "    Debug.Print AddNums(1, 2)",
        "End Sub",
        "",
    ]
)

BROKEN = CRLF.join(
    [
        "Option Explicit",
        "",
        "Public Sub Broken()",
        "    Dim n As Long",
        '    n = "not a number"',
        "End Sub",
        "",
    ]
)

FIXTURES: dict[str, Any] = {
    "workbook": {
        "kind": "excel-macro-workbook",
        "file_name": "Budget.xlsm",
        "why": (
            "A macro-enabled workbook created from the application's own template, with one "
            "standard module added. The module analyzes clean on purpose: a fixture that "
            "analyzes dirty makes every analysis case argue with the fixture."
        ),
        "modules": [{"name": "Helpers", "kind": "standard", "source": HELPERS}],
    },
    "word_document": {
        "kind": "word-document",
        "file_name": "Report.docm",
        "why": (
            "A host that is not Excel, so the refusals that turn on which application a "
            "file belongs to have a real subject rather than a hypothetical one."
        ),
        "modules": [],
    },
    "committed_workbook": {
        "kind": "excel-macro-workbook-in-a-git-repository",
        "file_name": "Budget.xlsm",
        "why": (
            "The same workbook, committed to a git repository at the workspace root with "
            "the file marked binary. A normalizing filter would corrupt the container, so "
            "a runner that skips that step tests a broken fixture rather than the tool."
        ),
        "modules": [{"name": "Helpers", "kind": "standard", "source": HELPERS}],
        "git": {"commit_message": "the workbook as it was", "attributes": "*.xlsm binary"},
    },
    "access_database": {
        "kind": "access-database",
        "file_name": "App.accdb",
        "why": (
            "The host that is different. Access keeps its VBA in the database rather than a "
            "vbaProject.bin, its container class releases differently, its save takes no "
            "signature flag, and it spells a module's kind as a string where the package "
            "hosts give an enum. Every one of those differences was a real defect first."
        ),
        "modules": [{"name": "Helpers", "kind": "standard", "source": ACCESS_MODULE}],
        "tables": [
            {"name": "Orders", "columns": [{"name": "Id", "type": "long"}]},
        ],
    },
    "vb6_project": {
        "kind": "visual-basic-6-project",
        "file_name": "Demo.vbp",
        "why": (
            "A .vbp is a text manifest whose modules are files on disk. It is in the surface "
            "because reading it AS a project - which files are modules, what each is called "
            "inside VB, and analyzing them together - is what a file tool cannot do alone."
        ),
        "files": {
            "Demo.vbp": CRLF.join(
                [
                    "Type=Exe",
                    "Form=Form1.frm",
                    "Module=Helpers; Helpers.bas",
                    'Startup="Form1"',
                    'Name="DemoProject"',
                    "MajorVer=1",
                    "",
                ]
            ),
            "Helpers.bas": VB6_HELPERS,
            "Form1.frm": VB6_FORM,
        },
        "encoding": "The machine's ANSI code page, which is what VB6 writes.",
    },
    "access_designs": {
        "kind": "access-database",
        "file_name": "Designs.accdb",
        "why": (
            "A database holding both kinds of design, because they live in separate "
            "collections and a reader that knows about only one reports the other as absent."
        ),
        "forms": [
            {
                "name": "Summary",
                "caption": "Totals",
                "width": 8000,
                "height": 3000,
                "controls": [
                    {
                        "type": "Label",
                        "name": "Title",
                        "left": 240,
                        "top": 240,
                        "width": 2000,
                        "height": 300,
                        "caption": "Hello",
                    }
                ],
            }
        ],
        "reports": [
            {
                "name": "Monthly",
                "controls": [
                    {
                        "type": "Label",
                        "name": "Banner",
                        "section": "PageHeaderSection",
                        "caption": "Header",
                    }
                ],
            }
        ],
    },
    "crowded_workbook": {
        "kind": "excel-macro-workbook",
        "file_name": "Crowded.xlsm",
        "why": (
            "More modules than any listing returns, so the bound and the note that explains "
            "it have something to act on. A legacy project of this size is ordinary."
        ),
        "modules": [
            {
                "name": "Mod000 .. Mod319",
                "kind": "standard",
                "source": "Any short standard module; 320 of them, numbered to three digits.",
            }
        ],
    },
    "loaded_query": {
        "kind": "excel-workbook",
        "file_name": "Loaded.xlsx",
        "why": (
            "A workbook whose one query is loaded onto a sheet, so the case about removing "
            "every part of it has all the parts to remove."
        ),
        "power_query": [
            {
                "name": "Numbers",
                "formula": "let Source = {1..10} in Source",
                "load_to_sheet": {"columns": ["Value"], "cell": "A1"},
            }
        ],
    },
    "workbook_with_a_form": {
        "kind": "excel-macro-workbook",
        "file_name": "Forms.xlsm",
        "why": (
            "A workbook holding a UserForm, which is two things that have to agree: a "
            "designer storage and a code module of the same name. Every case about keeping "
            "them together needs one."
        ),
        "forms": [
            {
                "name": "Wizard",
                "caption": "Setup",
                "width": 300,
                "height": 200,
                "controls": [
                    {"type": "CommandButton", "name": "Ok", "left": 12, "top": 12}
                ],
            }
        ],
    },
    "shapes_workbook": {
        "kind": "shipped-binary",
        "file_name": "Shapes.xlsm",
        "repository_path": "python/tests/fixtures/shapes.xlsm",
        "why": (
            "The one fixture shipped as a binary rather than described. A Forms-toolbar "
            "button lives in a VML shape, a ctrlProp part, a <controls> entry and a hidden "
            "DrawingML twin, and only Excel writes all four in agreement, so no library can "
            "build this. Copy it from the repository. Excel 16 authored it, and "
            "python/tests/test_shapes_live.py rebuilds it and checks it still reads the same."
        ),
        "contents": (
            "One sheet, Controls: RunButton, a Forms button at B2:C3 running "
            "Module1.DoTheThing with the caption 'Run it'; GoShape, a rectangle at E2:G4 "
            "running the same procedure with the text 'Go'; Ready, a check box at B6:D7 "
            "captioned 'Ready?' and linked to $D$6."
        ),
    },
    "plain_workbook": {
        "kind": "excel-workbook",
        "file_name": "Data.xlsx",
        "why": (
            "A workbook with no VBA project at all, carrying one Power Query. It pins the "
            "rule that Power Query and worksheet cells are reachable where macros are not."
        ),
        "power_query": [{"name": "Numbers", "formula": "let Source = {1..10} in Source"}],
    },
}


def case(
    case_id: str,
    why: str,
    fixture: str | None,
    steps: list[dict[str, Any]],
    expect: list[dict[str, Any]] | None = None,
    requires: str = "files",
) -> dict[str, Any]:
    entry: dict[str, Any] = {
        "id": case_id,
        "why": why,
        "requires": requires,
        "steps": steps,
    }
    if fixture:
        entry["fixture"] = fixture
    if expect:
        entry["expect"] = expect
    return entry


def step(tool: str, arguments: dict[str, Any], error_contains: str = "") -> dict[str, Any]:
    entry: dict[str, Any] = {"tool": tool, "arguments": arguments}
    if error_contains:
        entry["error_contains"] = error_contains
    return entry


def cases() -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []

    # ---------------------------------------------------------------- reading
    out.append(
        case(
            "project-info.reports-the-write-guards",
            "Before writing, an agent has to know whether the project is protected or signed. "
            "Both facts are present and false on an ordinary file, never merely absent.",
            "workbook",
            [step("xlide_project_info", {"file_path": "${fixture}"})],
            [
                {"path": "host", "equals": "excel"},
                {"path": "vba_readable", "equals": True},
                {"path": "password_protected", "equals": False},
                {"path": "digitally_signed", "equals": False},
                {"path": "modules", "at_least": 1},
            ],
        )
    )
    out.append(
        case(
            "read-module.strips-the-attribute-header",
            "A read returns what the VBA editor shows. The Attribute VB_* header is managed "
            "for the caller, and an agent that edits it by hand breaks a module's binding.",
            "workbook",
            [
                step(
                    "xlide_read_module",
                    {"file_path": "${fixture}", "module_name": "Helpers"},
                )
            ],
            [
                {"path": "module", "equals": "Helpers"},
                {"path": "kind", "equals": "standard"},
                {"path": "source", "not_contains": "Attribute VB_Name"},
                {"path": "source", "contains": "Public Function AddNums"},
                {"path": "content_token", "starts_with": "xlide1:"},
            ],
        )
    )
    out.append(
        case(
            "read-module.include-header-shows-it",
            "The header is reachable when it is genuinely wanted, so the default hiding it "
            "costs nothing.",
            "workbook",
            [
                step(
                    "xlide_read_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Helpers",
                        "include_header": True,
                    },
                )
            ],
            [{"path": "source", "contains": "Attribute VB_Name"}],
        )
    )
    out.append(
        case(
            "module-names.match-without-case",
            "VBA compares names without case, so every tool that takes a module name does too.",
            "workbook",
            [
                step(
                    "xlide_read_module",
                    {"file_path": "${fixture}", "module_name": "hELPers"},
                )
            ],
            [{"path": "module", "equals": "Helpers"}],
        )
    )
    out.append(
        case(
            "read-module.missing-names-what-exists",
            "A refusal that lists the modules that do exist saves the round trip an agent "
            "would otherwise spend discovering them.",
            "workbook",
            [
                step(
                    "xlide_read_module",
                    {"file_path": "${fixture}", "module_name": "Nope"},
                    error_contains="Helpers",
                )
            ],
        )
    )
    out.append(
        case(
            "list-procedures.joins-a-continued-signature",
            "A declaration split over lines with trailing underscores is one procedure. "
            "Reporting it as several is worse than reporting none.",
            "workbook",
            [
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Wrapped",
                        "source": CRLF.join(
                            [
                                "Option Explicit",
                                "",
                                "Public Function Pair(ByVal a As Long, _",
                                "                     ByVal b As Long) As Long",
                                "    Pair = a + b",
                                "End Function",
                                "",
                            ]
                        ),
                    },
                ),
                step(
                    "xlide_list_procedures",
                    {"file_path": "${fixture}", "module_name": "Wrapped"},
                ),
            ],
            [
                {"path": "procedures", "at_least": 1},
                {"path": "procedures[0].name", "equals": "Pair"},
                {"path": "procedures[0].kind", "equals": "Function"},
            ],
        )
    )

    # ---------------------------------------------------------------- writing
    out.append(
        case(
            "write-module.round-trips-and-reissues-the-token",
            "A write answers with the token of what it wrote, so the next write can be "
            "guarded without a second read.",
            "workbook",
            [
                step(
                    "xlide_read_module",
                    {"file_path": "${fixture}", "module_name": "Helpers"},
                ),
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Helpers",
                        "source": HELPERS.replace("a + b", "a + b + 1"),
                        "expected_content_token": "${step[0].content_token}",
                    },
                ),
                step(
                    "xlide_read_module",
                    {"file_path": "${fixture}", "module_name": "Helpers"},
                ),
            ],
            [
                {"path": "source", "contains": "a + b + 1"},
                {"path": "content_token", "equals": "${step[1].content_token}"},
            ],
        )
    )
    out.append(
        case(
            "write-module.reports-the-diff-of-what-landed",
            "The diff is taken against the module read back after saving, not against the "
            "source that was sent. That is what makes it evidence rather than an echo: a "
            "container that stored something other than what it was given shows here.",
            "workbook",
            [
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Helpers",
                        "source": HELPERS.replace("a + b", "a + b + 1"),
                    },
                )
            ],
            [
                {"path": "diff", "contains": "-    AddNums = a + b"},
                {"path": "diff", "contains": "+    AddNums = a + b + 1"},
                {"path": "lines_added", "equals": 1},
                {"path": "lines_removed", "equals": 1},
            ],
        )
    )
    out.append(
        case(
            "write-module.refuses-a-stale-token",
            "The guard that stops one writer discarding another's change. A token from a "
            "read that something else has since overtaken must refuse, and the refusal has "
            "to hand back the current token so the caller can recover in one step.",
            "workbook",
            [
                step(
                    "xlide_read_module",
                    {"file_path": "${fixture}", "module_name": "Helpers"},
                ),
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Helpers",
                        "source": HELPERS + "' someone else got here first" + CRLF,
                        "expected_content_token": "${step[0].content_token}",
                    },
                ),
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Helpers",
                        "source": HELPERS + "' my edit" + CRLF,
                        "expected_content_token": "${step[0].content_token}",
                    },
                    error_contains="xlide1:",
                ),
            ],
        )
    )
    out.append(
        case(
            "write-module.creates-what-is-missing",
            "Creating a module is a write with no token, not a separate tool.",
            "workbook",
            [
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Fresh",
                        "source": "Option Explicit" + CRLF,
                    },
                )
            ],
            [
                {"path": "created", "equals": True},
                {"path": "kind", "equals": "standard"},
                {"path": "saved", "equals": True},
            ],
        )
    )
    out.append(
        case(
            "write-module.creates-a-class-on-request",
            "Class modules are created through the same tool, by kind.",
            "workbook",
            [
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Widget",
                        "source": "Option Explicit" + CRLF,
                        "kind": "class",
                    },
                )
            ],
            [{"path": "created", "equals": True}, {"path": "kind", "equals": "class"}],
        )
    )
    out.append(
        case(
            "write-module.refuses-an-invalid-name",
            "A name VBA cannot carry is refused before the file is touched.",
            "workbook",
            [
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "1Bad",
                        "source": "Option Explicit" + CRLF,
                    },
                    error_contains="not a valid module name",
                )
            ],
        )
    )
    out.append(
        case(
            "write-module.refuses-a-reserved-word",
            "A module cannot be called Sub, whatever the caller intended.",
            "workbook",
            [
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Sub",
                        "source": "Option Explicit" + CRLF,
                    },
                    error_contains="reserved word",
                )
            ],
        )
    )
    out.append(
        case(
            "rename-module.refuses-a-document-module",
            "The host owns ThisWorkbook and recreates it by name. A renamed one leaves the "
            "project no longer matching its file.",
            "workbook",
            [
                step(
                    "xlide_rename_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "ThisWorkbook",
                        "new_name": "Renamed",
                    },
                    error_contains="document module",
                )
            ],
        )
    )
    out.append(
        case(
            "delete-module.refuses-a-document-module",
            "Deleting one is the same category error as renaming it.",
            "workbook",
            [
                step(
                    "xlide_delete_module",
                    {"file_path": "${fixture}", "module_name": "ThisWorkbook"},
                    error_contains="document module",
                )
            ],
        )
    )
    out.append(
        case(
            "rename-module.refuses-a-name-that-collides-without-case",
            "VBA compares without case, so Helpers and HELPERS are one name.",
            "workbook",
            [
                step(
                    "xlide_rename_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Helpers",
                        "new_name": "HELPERS",
                    },
                    error_contains="already exists",
                )
            ],
        )
    )
    out.append(
        case(
            "delete-module.removes-it",
            "A delete is reported by name, with what it cost.",
            "workbook",
            [
                step(
                    "xlide_delete_module",
                    {"file_path": "${fixture}", "module_name": "Helpers"},
                ),
                step("xlide_list_modules", {"file_path": "${fixture}"}),
            ],
            [{"path": "modules", "not_contains": "Helpers"}],
        )
    )

    # --------------------------------------------------------------- analysis
    out.append(
        case(
            "analyze.reports-a-clean-file-as-clean",
            "The build gate is only usable if a clean file comes back clean.",
            "workbook",
            [step("xlide_analyze", {"file_path": "${fixture}"})],
            [
                {"path": "counts.error", "equals": 0},
                {"path": "verdict", "equals": "clean"},
            ],
        )
    )
    out.append(
        case(
            "analyze.positions-match-what-a-read-returns",
            "The most consequential rule in the whole surface. Analysis measures a source "
            "that carries the attribute header; a read strips it. An implementation that "
            "does not shift the position reports every diagnostic one line late, and an "
            "agent acting on it edits the wrong line.",
            "workbook",
            [
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Broken",
                        "source": BROKEN,
                    },
                ),
                step(
                    "xlide_analyze",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Broken",
                        "min_severity": "error",
                    },
                ),
            ],
            [
                {"path": "problems", "at_least": 1},
                {"path": "problems[0].severity", "equals": "error"},
                {"path": "problems[0].line", "equals": 5},
                {"path": "problems[0].text", "contains": "not a number"},
                {"path": "verdict", "equals": "errors found"},
            ],
        )
    )
    out.append(
        case(
            "analyze.counts-stay-project-wide-when-a-module-is-named",
            "Narrowing the report must not hide from the agent that the file is not clean.",
            "workbook",
            [
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Broken",
                        "source": BROKEN,
                    },
                ),
                step(
                    "xlide_analyze",
                    {"file_path": "${fixture}", "module_name": "Helpers"},
                ),
            ],
            [
                {"path": "counts.error", "at_least": 1},
                {"path": "verdict", "equals": "errors found"},
            ],
        )
    )
    out.append(
        case(
            "analyze-source.needs-no-file",
            "Checking generated code before writing it costs nothing and catches the compile "
            "errors that would otherwise surface in front of the user.",
            None,
            [step("xlide_analyze_source", {"source": BROKEN, "host": "excel"})],
            [
                {"path": "verdict", "equals": "errors found"},
                {"path": "counts.error", "at_least": 1},
            ],
        )
    )
    out.append(
        case(
            "analyze-source.against-a-project-resolves-its-calls",
            "A draft module that calls into the project must resolve against it, or every "
            "check of new code drowns in names it cannot see.",
            "workbook",
            [
                step(
                    "xlide_analyze_source",
                    {
                        "source": CRLF.join(
                            [
                                "Option Explicit",
                                "",
                                "Public Sub Use()",
                                "    Debug.Print AddNums(1, 2)",
                                "End Sub",
                                "",
                            ]
                        ),
                        "module_name": "Draft",
                        "file_path": "${fixture}",
                    },
                )
            ],
            [
                {"path": "counts.error", "equals": 0},
                {"path": "host", "equals": "excel"},
            ],
        )
    )
    out.append(
        case(
            "analyze.refuses-an-unknown-host",
            "Only the four hosts have object models here, and a typo must not silently fall "
            "back to analyzing against nothing.",
            None,
            [
                step(
                    "xlide_analyze_source",
                    {"source": "Sub A()" + CRLF + "End Sub" + CRLF, "host": "outlook"},
                    error_contains="not a host",
                )
            ],
        )
    )

    # ------------------------------------------------------------- discovery
    out.append(
        case(
            "list-projects.explains-what-it-cannot-open",
            "A file left silently out of a listing sends the user looking for it. A .xlsx "
            "is listed, marked unreadable, and told why.",
            "plain_workbook",
            [step("xlide_list_projects", {})],
            [
                {"path": "files", "at_least": 1},
                {"path": "files", "contains": "no VBA project by design"},
                {"path": "files", "contains": "\"readable\": false"},
            ],
        )
    )
    out.append(
        case(
            "project-info.reaches-power-query-in-a-file-with-no-macros",
            "A workbook with no VBA can still be doing all of its work in Power Query.",
            "plain_workbook",
            [step("xlide_project_info", {"file_path": "${fixture}"})],
            [
                {"path": "vba_readable", "equals": False},
                {"path": "power_query.count", "equals": 1},
                {"path": "sheets", "at_least": 1},
            ],
        )
    )
    out.append(
        case(
            "create-project.never-overwrites",
            "Creating over an existing file would destroy work with no undo.",
            "workbook",
            [
                step(
                    "xlide_create_project",
                    {"file_path": "${fixture}"},
                    error_contains="already exists",
                )
            ],
        )
    )

    # ------------------------------------------------------------------ cells
    out.append(
        case(
            "cells.round-trip-values-and-formulas",
            "A formula is stored as typed and read back as typed, whatever prefix the file "
            "format requires in between.",
            "plain_workbook",
            [
                step(
                    "xlide_write_cells",
                    {
                        "file_path": "${fixture}",
                        "sheet": "Sheet1",
                        "start_cell": "A1",
                        "data": [[1], [2], ["=SUM(A1:A2)"], ["=XLOOKUP(1,A1:A2,A1:A2)"]],
                    },
                ),
                step(
                    "xlide_read_cells",
                    {
                        "file_path": "${fixture}",
                        "sheet": "Sheet1",
                        "cell_range": "A1:A4",
                        "include": "both",
                    },
                ),
            ],
            [
                {"path": "values[0][0]", "equals": 1},
                {"path": "formulas[2][0]", "equals": "=SUM(A1:A2)"},
                {"path": "formulas[3][0]", "equals": "=XLOOKUP(1,A1:A2,A1:A2)"},
                {"path": "values[2][0]", "type": "null"},
            ],
        )
    )
    out.append(
        case(
            "cells.a-write-never-claims-a-result",
            "Nothing here calculates. A result Excel has not produced must not be reported "
            "as though it had.",
            "plain_workbook",
            [
                step(
                    "xlide_write_cells",
                    {
                        "file_path": "${fixture}",
                        "sheet": "Sheet1",
                        "start_cell": "A1",
                        "data": [["=1+1"]],
                    },
                )
            ],
            [
                {"path": "recalculated", "equals": False},
                {"path": "saved", "equals": True},
            ],
        )
    )
    out.append(
        case(
            "cells.report-what-was-overwritten",
            "Overwriting data is the user's decision, and they can only make it if the "
            "agent is told what it displaced.",
            "plain_workbook",
            [
                step(
                    "xlide_write_cells",
                    {
                        "file_path": "${fixture}",
                        "sheet": "Sheet1",
                        "start_cell": "A1",
                        "data": [["first"], ["=1+1"]],
                    },
                ),
                step(
                    "xlide_write_cells",
                    {
                        "file_path": "${fixture}",
                        "sheet": "Sheet1",
                        "start_cell": "A1",
                        "data": [["second"], ["third"]],
                    },
                ),
            ],
            [
                {"path": "cells_overwritten", "equals": 2},
                {"path": "formulas_replaced", "equals": 1},
            ],
        )
    )
    out.append(
        case(
            "cells.unknown-sheet-names-the-ones-that-exist",
            "The same courtesy a missing module gets.",
            "plain_workbook",
            [
                step(
                    "xlide_read_cells",
                    {"file_path": "${fixture}", "sheet": "Nope", "cell_range": "A1"},
                    error_contains="Sheet1",
                )
            ],
        )
    )

    # ------------------------------------------------------------ power query
    out.append(
        case(
            "power-query.read-and-write",
            "M is code, edited the way VBA is.",
            "plain_workbook",
            [
                step(
                    "xlide_write_query",
                    {
                        "file_path": "${fixture}",
                        "action": "set",
                        "query_name": "Numbers",
                        "formula": "let Source = {1..20} in Source",
                    },
                ),
                step(
                    "xlide_read_query",
                    {"file_path": "${fixture}", "query_name": "numbers"},
                ),
            ],
            [
                {"path": "query", "equals": "Numbers"},
                {"path": "formula", "contains": "{1..20}"},
            ],
        )
    )
    out.append(
        case(
            "power-query.a-query-loading-nowhere-has-no-refresh-settings",
            "Refresh settings belong to a connection, and a query that loads nowhere has "
            "none. An implementation that treats this as an error fails on the common case.",
            "plain_workbook",
            [
                step(
                    "xlide_read_query",
                    {"file_path": "${fixture}", "query_name": "Numbers"},
                )
            ],
            [{"path": "refresh_note", "contains": "loads nowhere"}],
        )
    )
    out.append(
        case(
            "power-query.refused-on-another-host",
            "Power Query is an Excel feature. A Word file asking for it gets a refusal that "
            "names the host, not an empty list that reads as 'this document has no queries'.",
            "word_document",
            [
                step(
                    "xlide_list_queries",
                    {"file_path": "${fixture}"},
                    error_contains="Excel packages",
                )
            ],
        )
    )

    # -------------------------------------------------------------- disk sync
    out.append(
        case(
            "export.previews-before-it-writes",
            "An export that silently overwrote a folder of reviewed files would lose the "
            "diff before anyone noticed.",
            "workbook",
            [
                step(
                    "xlide_export_modules",
                    {"file_path": "${fixture}", "export_folder": "${folder:vba}"},
                )
            ],
            [{"path": "applied", "equals": False}, {"path": "plan", "at_least": 1}],
        )
    )
    out.append(
        case(
            "import.previews-before-it-writes",
            "The rule agents get wrong most often stated as a guard: an import reports what "
            "it would change and writes nothing, so an edited export is still only a copy "
            "until someone applies it deliberately.",
            "workbook",
            [
                step(
                    "xlide_export_modules",
                    {
                        "file_path": "${fixture}",
                        "export_folder": "${folder:vba}",
                        "apply": True,
                    },
                ),
                step(
                    "xlide_import_modules",
                    {"file_path": "${fixture}", "source_folder": "${folder:vba}"},
                ),
            ],
            [
                {"path": "applied", "equals": False},
                {"path": "plan", "at_least": 1},
            ],
        )
    )

    # ----------------------------------------------------------------- access
    out.append(
        case(
            "access.a-standard-module-is-classified-standard",
            "An Access module carries no designer header, so its text cannot say whether it "
            "is a class, and Access spells the kind as a string where the package hosts give "
            "an enum. Comparing against only the enum reports every standard module in every "
            "database as a class, which then exports it as .cls and analyzes it as one.",
            "access_database",
            [step("xlide_list_modules", {"file_path": "${fixture}"})],
            [
                {"path": "host", "equals": "access"},
                {"path": "modules", "contains": '"kind": "standard"'},
            ],
        )
    )
    out.append(
        case(
            "access.a-module-round-trips",
            "Access writes into the database as the edit is made and its save takes no "
            "signature flag, so the write path that works for a package host is not the one "
            "that works here.",
            "access_database",
            [
                step(
                    "xlide_read_module",
                    {"file_path": "${fixture}", "module_name": "Helpers"},
                ),
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Helpers",
                        "source": ACCESS_MODULE.replace("n * 2", "n * 3"),
                        "expected_content_token": "${step[0].content_token}",
                    },
                ),
                step(
                    "xlide_read_module",
                    {"file_path": "${fixture}", "module_name": "Helpers"},
                ),
            ],
            [{"path": "source", "contains": "n * 3"}],
        )
    )
    out.append(
        case(
            "access.signature-state-is-unknown-rather-than-false",
            "Access keeps its VBA in the database, not in the streams beside a "
            "vbaProject.bin where this server looks for a signature. Answering false would "
            "be a claim nothing checked.",
            "access_database",
            [step("xlide_project_info", {"file_path": "${fixture}"})],
            [{"path": "digitally_signed", "type": "null"}],
        )
    )
    out.append(
        case(
            "access.the-catalog-reads-tables-and-queries",
            "An .accdb is an application rather than a document, and its VBA is written "
            "against tables that reading the modules alone never shows.",
            "access_database",
            [step("xlide_access_catalog", {"file_path": "${fixture}", "include": "tables"})],
            [
                {"path": "tables", "contains": '"name": "Orders"'},
                {"path": "tables", "not_contains": "MSys"},
            ],
        )
    )

    out.append(
        case(
            "access.reports-are-listed-beside-forms",
            "Access keeps reports in a collection of their own. A reader that calls forms() "
            "alone reports every report in every database as not existing, and an agent then "
            "tells the user their database has none.",
            "access_designs",
            [step("xlide_list_forms", {"file_path": "${fixture}"})],
            [
                {"path": "count", "equals": 2},
                {"path": "forms", "contains": '"design": "report"'},
                {"path": "forms", "contains": '"design": "form"'},
                {"path": "geometry_unit", "equals": "twips"},
            ],
        )
    )
    out.append(
        case(
            "access.a-design-names-the-sections-a-control-can-go-in",
            "A control on an Access design lives in a band. Without the band names an agent "
            "has to guess one, and a wrong guess is only refused after the call.",
            "access_designs",
            [step("xlide_read_form", {"file_path": "${fixture}", "form_name": "Monthly"})],
            [
                {"path": "design", "equals": "report"},
                {"path": "sections", "contains": "PageHeaderSection"},
            ],
        )
    )
    out.append(
        case(
            "access.unnamed-properties-are-counted-not-listed",
            "An Access design stores more property ids than the library can name: 22 of a "
            "bare form's 29. Listing `Unidentified314: 4` beside `Caption` buries what a "
            "reader came for, and hiding them without saying so would be a different lie.",
            "access_designs",
            [step("xlide_read_form", {"file_path": "${fixture}", "form_name": "Summary"})],
            [
                {"path": "properties", "not_contains": "Unidentified"},
                {"path": "properties._unnamed_property_count", "at_least": 1},
            ],
        )
    )

    # ------------------------------------------------------------------- vb6
    out.append(
        case(
            "vb6.a-form-reads-as-code-not-as-markup",
            "A VB6 form keeps a designer block ahead of its attributes. It is what VB draws "
            "the form from rather than something to edit, and a reader that hands it over as "
            "source invites an edit that stops the project loading.",
            "vb6_project",
            [step("xlide_read_module", {"file_path": "${fixture}", "module_name": "Form1"})],
            [
                {"path": "kind", "equals": "userform"},
                {"path": "source", "starts_with": "Option Explicit"},
                {"path": "source", "not_contains": "Begin VB.Form"},
                {"path": "source", "contains": "Private Sub Form_Load"},
            ],
        )
    )
    out.append(
        case(
            "vb6.analysis-resolves-across-the-project",
            "The thing a file tool cannot do. Form_Load calls AddNums, which lives in another "
            "file; analyzed one file at a time it is an undefined name.",
            "vb6_project",
            [step("xlide_analyze", {"file_path": "${fixture}"})],
            [
                {"path": "host", "equals": "vb6"},
                {"path": "counts.error", "equals": 0},
            ],
        )
    )
    out.append(
        case(
            "vb6.a-rename-moves-the-name-the-file-and-the-manifest",
            "VB6 keeps a module's name in three places: the manifest line, the file name and "
            "the VB_Name attribute. Moving one without the others leaves a project that will "
            "not load.",
            "vb6_project",
            [
                step(
                    "xlide_rename_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Helpers",
                        "new_name": "Tools",
                    },
                ),
                step("xlide_list_modules", {"file_path": "${fixture}"}),
            ],
            [
                {"path": "modules", "contains": "Tools"},
                {"path": "modules", "not_contains": "Helpers"},
            ],
        )
    )

    out.append(
        case(
            "power-query.a-query-loads-onto-a-sheet-and-comes-off-again",
            "Loading needs the column names, because the connection has to name them and "
            "knowing them means running the query, which nothing here does. Excel settles "
            "them against the real result on the first refresh.",
            "plain_workbook",
            [
                step(
                    "xlide_write_query",
                    {
                        "file_path": "${fixture}",
                        "action": "load",
                        "query_name": "Numbers",
                        "columns": ["Value"],
                        "cell": "A1",
                    },
                ),
                step(
                    "xlide_read_query",
                    {"file_path": "${fixture}", "query_name": "Numbers"},
                ),
            ],
            [
                {"path": "load_target", "equals": "table"},
                # A loaded query has a connection, and therefore refresh settings.
                {"path": "refresh", "type": "object"},
            ],
        )
    )
    out.append(
        case(
            "power-query.load-without-column-names-is-refused",
            "A connection with no columns is one Excel cannot refresh. Guessing them here "
            "would be inventing data, so the caller is asked for what it expects.",
            "plain_workbook",
            [
                step(
                    "xlide_write_query",
                    {
                        "file_path": "${fixture}",
                        "action": "load",
                        "query_name": "Numbers",
                    },
                    error_contains="columns is required",
                )
            ],
        )
    )

    # ------------------------------------------------------- the other halves
    out.append(
        case(
            "power-query.removing-a-loaded-query-takes-its-plumbing",
            "A query loaded onto a sheet is four things: the definition, a connection, a "
            "query table and the table itself. Removing only the definition leaves a "
            "connection pointing at a query that no longer exists, which Excel meets on the "
            "next refresh rather than on open, so nothing says so at the time.",
            "loaded_query",
            [
                step(
                    "xlide_write_query",
                    {
                        "file_path": "${fixture}",
                        "action": "remove",
                        "query_name": "Numbers",
                    },
                )
            ],
            [{"path": "unloaded_from_sheet", "equals": True}],
        )
    )
    out.append(
        case(
            "modules.deleting-one-warns-about-the-buttons-that-called-it",
            "Nothing rewrites an OnAction, so a button keeps naming a procedure that has "
            "gone and the user finds out by clicking it. Deleting such a module is a "
            "legitimate thing to do, so this warns rather than refusing.",
            "shapes_workbook",
            [
                step(
                    "xlide_delete_module",
                    {"file_path": "${fixture}", "module_name": "Module1"},
                )
            ],
            [
                {"path": "shapes_now_calling_nothing", "at_least": 2},
                {"path": "warning", "contains": "xlide_set_shape_macro"},
            ],
        )
    )

    # -------------------------------------------------------------- form life
    out.append(
        case(
            "forms.a-userforms-module-is-not-a-class",
            "A UserForm's module is a class as far as its text goes; what makes it a form is "
            "the designer storage beside it. Reading the kind from the text alone calls every "
            "form a class, and a class is a thing this server will happily rename.",
            "workbook_with_a_form",
            [step("xlide_list_modules", {"file_path": "${fixture}"})],
            [{"path": "modules", "contains": '"kind": "userform"'}],
        )
    )
    out.append(
        case(
            "forms.renaming-a-userforms-module-is-refused",
            "Renaming only the module leaves a storage with no module, which the host does "
            "not show, and a module with no storage, which is a class. The form has silently "
            "gone and the file still opens, so nothing tells the user.",
            "workbook_with_a_form",
            [
                step(
                    "xlide_rename_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Wizard",
                        "new_name": "Renamed",
                    },
                    error_contains="design",
                )
            ],
        )
    )
    out.append(
        case(
            "forms.deleting-a-userforms-module-is-refused",
            "Deleting only the module leaves the designer storage behind entirely, with the "
            "same silent result.",
            "workbook_with_a_form",
            [
                step(
                    "xlide_delete_module",
                    {"file_path": "${fixture}", "module_name": "Wizard"},
                    error_contains="orphaned",
                )
            ],
        )
    )
    out.append(
        case(
            "forms.creating-one-writes-both-halves",
            "A module with no storage is a class rather than a form, which is why creating a "
            "form is its own tool rather than a write of an empty module.",
            "workbook",
            [
                step(
                    "xlide_manage_form",
                    {
                        "file_path": "${fixture}",
                        "action": "create",
                        "form_name": "Setup",
                        "caption": "Set it up",
                    },
                ),
                step("xlide_list_modules", {"file_path": "${fixture}"}),
            ],
            [{"path": "modules", "contains": '"name": "Setup"'},
             {"path": "modules", "contains": '"kind": "userform"'}],
        )
    )
    out.append(
        case(
            "forms.an-access-design-renames-with-its-code",
            "Access moves the design and its module together, which is why renaming is "
            "allowed there and refused for a UserForm. The same operation is not the same "
            "operation on every host.",
            "access_designs",
            [
                step(
                    "xlide_manage_form",
                    {
                        "file_path": "${fixture}",
                        "action": "rename",
                        "form_name": "Summary",
                        "new_name": "Totals",
                    },
                ),
                step("xlide_list_modules", {"file_path": "${fixture}"}),
            ],
            [
                {"path": "modules", "contains": "Form_Totals"},
                {"path": "modules", "not_contains": "Form_Summary"},
            ],
        )
    )

    # ----------------------------------------------------------------- shapes
    out.append(
        case(
            "shapes.a-button-carries-the-macro-it-runs",
            "The link an agent is usually actually after. A button on a sheet starts the "
            "automation, and nothing else in the surface shows that a Sub is reachable from "
            "one; an agent renaming that Sub has no way to know an OnAction names it, "
            "because nothing rewrites an OnAction.",
            "shapes_workbook",
            [step("xlide_list_shapes", {"file_path": "${fixture}"})],
            [
                {"path": "shape_count", "equals": 3},
                {"path": "macros_run_by_shapes", "at_least": 2},
                {"path": "macros_run_by_shapes", "contains": "Module1.DoTheThing"},
                {"path": "sheets[0].shapes", "contains": '"kind": "button"'},
                {"path": "sheets[0].shapes", "contains": '"cells": "B2:C3"'},
            ],
        )
    )
    out.append(
        case(
            "shapes.a-visible-control-is-not-reported-hidden",
            "Excel marks the DrawingML twin of every form control hidden, visible or not. "
            "Reading that flag off the twin reports every button on every sheet as hidden, "
            "which is worse than not reporting it at all; the VML style is the answer.",
            "shapes_workbook",
            [step("xlide_list_shapes", {"file_path": "${fixture}"})],
            [{"path": "sheets[0].shapes", "not_contains": '"hidden": true'}],
        )
    )
    out.append(
        case(
            "shapes.a-check-box-carries-its-linked-cell-and-no-macro",
            "A control that runs nothing must not be given a macro, and the cell it writes "
            "to is how a reader works out what it is for.",
            "shapes_workbook",
            [step("xlide_list_shapes", {"file_path": "${fixture}", "sheet": "controls"})],
            [
                {"path": "sheets[0].sheet", "equals": "Controls"},
                {"path": "sheets[0].shapes", "contains": '"linked_cell": "$D$6"'},
                {"path": "sheets[0].shapes", "contains": '"kind": "checkBox"'},
            ],
        )
    )

    out.append(
        case(
            "shapes.a-macro-can-be-repointed-in-every-part-that-holds-it",
            "A Forms button stores its macro twice, in its VML shape and in the sheet's "
            "controls entry, and Excel reads both. Writing one and not the other leaves a "
            "button whose behaviour depends on which copy Excel happens to trust.",
            "shapes_workbook",
            [
                step(
                    "xlide_set_shape_macro",
                    {
                        "file_path": "${fixture}",
                        "sheet": "Controls",
                        "shape_name": "RunButton",
                        "macro": "Module1.Renamed",
                    },
                ),
                step("xlide_list_shapes", {"file_path": "${fixture}"}),
            ],
            [
                {"path": "shape_count", "equals": 3},
                {"path": "macros_run_by_shapes", "contains": "Module1.Renamed"},
            ],
        )
    )
    out.append(
        case(
            "shapes.an-empty-macro-clears-the-link",
            "Unlinking a button is a thing people do, and it is not the same as deleting it.",
            "shapes_workbook",
            [
                step(
                    "xlide_set_shape_macro",
                    {
                        "file_path": "${fixture}",
                        "sheet": "Controls",
                        "shape_name": "GoShape",
                        "macro": "",
                    },
                ),
                step("xlide_list_shapes", {"file_path": "${fixture}"}),
            ],
            [
                {"path": "shape_count", "equals": 3},
                {"path": "macros_run_by_shapes", "not_contains": "GoShape"},
            ],
        )
    )

    # ------------------------------------------------------------ git changes
    out.append(
        case(
            "git-changes.reports-a-modified-module-with-its-diff",
            "git diff cannot answer this: an Office file is one binary blob, so a commit "
            "that changed a line and one that replaced the project look identical. Reading "
            "the blob at the revision and diffing the VBA inside it is the only way a "
            "workbook can be reviewed at all.",
            "committed_workbook",
            [
                step(
                    "xlide_write_module",
                    {
                        "file_path": "${fixture}",
                        "module_name": "Helpers",
                        "source": HELPERS.replace("a + b", "a + b + 1"),
                    },
                ),
                step("xlide_git_changes", {"file_path": "${fixture}"}),
            ],
            [
                {"path": "changed", "equals": 1},
                {"path": "changes", "contains": '"status": "modified"'},
                {"path": "changes", "contains": '"kind": "module"'},
                {"path": "changes", "contains": "AddNums = a + b + 1"},
            ],
            requires="git",
        )
    )
    out.append(
        case(
            "git-changes.reports-a-power-query-change-too",
            "A workbook's logic can move entirely in its M with no module touched. A "
            "comparison that only looked at VBA would call that commit empty, which is a "
            "wrong answer rather than a missing one.",
            "committed_workbook",
            [
                step(
                    "xlide_write_query",
                    {
                        "file_path": "${fixture}",
                        "action": "set",
                        "query_name": "Orders",
                        "formula": "let Source = {1..10} in Source",
                    },
                ),
                step("xlide_git_changes", {"file_path": "${fixture}"}),
            ],
            [
                {"path": "changed", "equals": 1},
                {"path": "changes", "contains": '"kind": "query"'},
                {"path": "changes", "contains": "let Source = {1..10} in Source"},
            ],
            requires="git",
        )
    )
    out.append(
        case(
            "git-changes.says-what-it-does-not-compare",
            "Cell values are not compared, and a reviewer who does not know that reads an "
            "empty report as an unchanged workbook. The scope travels with every answer "
            "rather than living in the documentation.",
            "committed_workbook",
            [step("xlide_git_changes", {"file_path": "${fixture}"})],
            [{"path": "covers", "contains": "cell values are not compared"}],
            requires="git",
        )
    )
    out.append(
        case(
            "git-changes.an-unchanged-file-reports-nothing",
            "The common case has to be quiet, or nobody calls it before committing.",
            "committed_workbook",
            [step("xlide_git_changes", {"file_path": "${fixture}"})],
            [
                {"path": "changed", "equals": 0},
                {"path": "verdict", "equals": "no changes"},
            ],
            requires="git",
        )
    )
    out.append(
        case(
            "git-changes.outside-a-repository-says-so",
            "There is no revision to compare against, and the refusal says what to call "
            "instead rather than leaving the agent to guess the tool is broken.",
            "workbook",
            [
                step(
                    "xlide_git_changes",
                    {"file_path": "${fixture}"},
                    error_contains="not inside a git repository",
                )
            ],
            requires="git",
        )
    )

    # ----------------------------------------------------------------- bounds
    out.append(
        case(
            "bounds.a-listing-that-stops-short-says-so",
            "A tool result is model context, and a listing that silently stops short reads "
            "as a complete answer: an agent reports the workbook has 300 modules, or that "
            "the module it was asked about does not exist. The count stays the project's, "
            "and a note says what was withheld.",
            "crowded_workbook",
            [step("xlide_list_modules", {"file_path": "${fixture}"})],
            [
                {"path": "count", "at_least": 320},
                {"path": "note", "contains": "are not"},
            ],
        )
    )
    out.append(
        case(
            "bounds.a-direct-read-is-not-bounded-by-a-listing",
            "Bounding a listing must not put anything out of reach. A module past the limit "
            "is still read by name, or the bound has hidden it rather than summarized it.",
            "crowded_workbook",
            [
                step(
                    "xlide_read_module",
                    {"file_path": "${fixture}", "module_name": "Mod319"},
                )
            ],
            [{"path": "module", "equals": "Mod319"}],
        )
    )

    # ----------------------------------------------------------- the boundary
    out.append(
        case(
            "workspace.refuses-a-path-outside-the-roots",
            "A path argument is untrusted input. Reach is bounded by configuration, not by "
            "what the caller asks for.",
            None,
            [
                step(
                    "xlide_list_modules",
                    {"file_path": "${outside}/elsewhere.xlsm"},
                    error_contains="outside this server's workspace",
                )
            ],
        )
    )
    out.append(
        case(
            "workspace.refuses-traversal-out-of-a-root",
            "The same boundary, reached the other way.",
            "workbook",
            [
                step(
                    "xlide_list_modules",
                    {"file_path": "${fixture}/../../elsewhere.xlsm"},
                    error_contains="outside this server's workspace",
                )
            ],
        )
    )

    return out


def build() -> dict[str, Any]:
    from xlide_mcp import __version__

    return {
        "conformance_version": __version__,
        "reference_implementation": "python",
        "note": (
            "Behaviour every implementation in this repository owes. The Python suite runs "
            "this file, so a case that is wrong fails there before a port ever sees it."
        ),
        "requires": {
            "files": "Reads and writes the Office file. No Office installation, any platform.",
            "office": "Runs VBA in a desktop application. Windows, with the application.",
            "git": "Reads and writes the Office file, and needs git on the PATH.",
            "live": "Talks to a running xlide_vbide session inside the Visual Basic Editor.",
        },
        "fixtures": FIXTURES,
        "case_count": len(cases()),
        "cases": cases(),
    }


def render(corpus: dict[str, Any]) -> str:
    return json.dumps(corpus, indent=2) + "\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Export the conformance corpus.")
    parser.add_argument("--check", action="store_true", help="Fail if the file is out of date.")
    parser.add_argument("--out", type=Path, default=CONFORMANCE_PATH)
    args = parser.parse_args(argv)

    rendered = render(build())
    if args.check:
        if not args.out.is_file() or args.out.read_text(encoding="utf-8") != rendered:
            print(
                f"{args.out} is out of date. Regenerate: python tools/export_conformance.py",
                file=sys.stderr,
            )
            return 1
        print(f"{args.out.name} is current.")
        return 0

    args.out.parent.mkdir(parents=True, exist_ok=True)
    # Written LF on every platform. The repository normalizes to LF, so an
    # artifact regenerated on Windows with native line endings would read as
    # modified the moment it was written.
    args.out.write_text(rendered, encoding="utf-8", newline="\n")
    print(f"Wrote {args.out} ({len(cases())} cases).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
