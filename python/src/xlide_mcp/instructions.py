"""What the calling model is told before it uses any of this.

MCP hands `instructions` to the client at initialize, so this is the one piece of
text every agent sees whether or not the user pasted anything into a config file.
It is short on purpose. The rules here are the ones that stop real damage: the file
is the source of truth, exports are copies, analysis is a build gate, and the
guarded operations are the user's decision rather than the agent's.

Each tool carries its own description for the detail; this says how they fit
together and what not to do.
"""

from __future__ import annotations

SERVER_INSTRUCTIONS = """\
Read and change the VBA, UserForms, Power Query and worksheet cells inside Office \
files: Excel (.xlsm, .xlsb, .xlam, .xls), Word (.docm, .dotm, .doc), PowerPoint \
(.pptm, .potm) and Access (.accdb, .mdb), plus Power Query and sheets in plain \
.xlsx, and Visual Basic 6 projects (.vbp).

A .vbp's modules are files on disk, so its writes land as they are made and it is \
never protected or signed. Everything else works the same way.

Reading and writing the file needs no Office installation and works on any platform. \
Running macros and tests needs Windows with the application installed.

Workflow

1. xlide_list_projects when the user has not named a file. Pass absolute paths after that.
2. xlide_project_info once per file. It lists modules, forms, queries and sheets, and \
says whether the project is password-protected or digitally signed.
3. xlide_read_module to read. Its result carries a content_token; pass that back as \
expected_content_token to xlide_write_module and the write is refused if the module \
changed after your read.
4. xlide_write_module to change code. It replaces the module's whole source, so send \
all of it. The attribute header is managed for you: send the body.
5. xlide_analyze after every VBA change. Treat any error-severity problem as a build \
failure: fix it and analyze again until it is clean.
6. xlide_run_tests when the file has tests and the change touches tested code.

Rules

- The VBA inside the file is the only source of truth for it. Exported .bas and .cls \
files are copies and go stale. Never report a module as changed because you wrote an \
exported copy.
- Do not unzip, rezip or hex-edit an Office file, and do not touch vbaProject.bin. \
Those edits corrupt the file or drop its code silently.
- Do not drive Office with COM automation of your own, and never open, close or kill an \
Office application the user is running. xlide_run_macro and xlide_run_tests use an \
instance this server owns and holds a deadline over.
- Read the code you are about to change, and change only what the task needs.
- Ask the user before anything hard to undo: deleting a module, a form or a query, \
overwriting cells that hold data, or writing to a project that is password-protected or \
digitally signed. A signed project loses its signature on any macro change.
- Keep VBA source ASCII unless the user asks otherwise.
- To show the user what you changed, use the diff a write returns: it is taken from \
the file read back after saving, so it is what the file now holds rather than what you \
meant to write. xlide_git_changes does the same against a git revision, and covers \
Power Query as well as VBA. Neither compares cell values.
- A write fails while the file is open in its Office application. Tell the user; do not \
close the application for them.
- Access runs its compiled project, so a module written here takes effect when Access \
next opens the database.
- Cell values read from the file are what Excel last calculated. A formula you write, \
and anything depending on a cell you write, keeps its old result until Excel next opens \
the workbook. Do not report a result Excel has not calculated.
- Report what you changed, how you verified it, and what the user still has to do.
"""
