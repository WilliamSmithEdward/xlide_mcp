"""The VB6 fixture's files, in one place.

Both `test_vb6.py` and the conformance fixture in `conftest.py` build the same
project, and the conformance corpus describes these exact contents. Two copies of
a fixture drift, and a corpus that describes one while the suite runs the other
proves nothing.
"""

from __future__ import annotations

CRLF = "\r\n"

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

# A form keeps a designer block ahead of its attributes. It is what VB draws the
# form from, it is not code, and a read must not hand it over as something to edit.
VB6_FORM = CRLF.join(
    [
        "VERSION 5.00",
        "Begin VB.Form Form1 ",
        '   Caption         =   "Demo"',
        "   ClientHeight    =   3000",
        "End",
        'Attribute VB_Name = "Form1"',
        "Attribute VB_GlobalNameSpace = False",
        "Option Explicit",
        "",
        "Private Sub Form_Load()",
        "    Debug.Print AddNums(1, 2)",
        "End Sub",
        "",
    ]
)

VB6_MANIFEST = CRLF.join(
    [
        "Type=Exe",
        "Form=Form1.frm",
        "Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
        "..\\..\\Windows\\System32\\stdole2.tlb#OLE Automation",
        "Module=Helpers; Helpers.bas",
        'Startup="Form1"',
        'Name="DemoProject"',
        "MajorVer=1",
        "",
    ]
)
