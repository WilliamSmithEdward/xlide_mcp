# Changelog

Notable changes to `xlide-mcp`. The version lives in
`python/src/xlide_mcp/_version.py`, and a `v*.*.*` tag releases it.

## [0.1.0]

First release. An MCP server for the VBA, UserForms, Power Query and worksheet
cells inside Office files, and for Visual Basic 6 projects.

### Added

- Forty tools over three layers. The file layer needs no Office installation and
  runs anywhere; the execution layer runs macros, tests and compile checks in a
  desktop application the server owns and holds a deadline over; the live layer
  reads a running `xlide_vbide` session inside the Visual Basic Editor.
- Excel, Word, PowerPoint and Access, plus `.vbp` projects read as projects so
  their modules analyze together.
- Worksheet cells read and written in the OOXML package, and, for `.xlsb` and
  `.xls`, through Excel, with `source` and `recalculated` on every result saying
  which answered.
- A generated contract under `contract/`: the tool surface, and 68 conformance
  cases pinning what the answers mean. Every implementation in the repository is
  verified against them, and the Python suite runs them too, so a case that is
  wrong fails before a port is built on it.
- What changed inside a file, three ways. A write reports a unified diff of what
  the file now holds, taken from the read-back rather than from what was sent.
  `xlide_git_changes` compares against any git revision, one entry per module and
  per query. `xlide-mcp --textconv` is a git textconv driver, so `git diff`,
  `git show` and `git log -p` render a workbook as its VBA and M instead of
  reporting that two binaries differ.

### The guards, each with a conformance case behind it

- A path argument resolves, symlinks included, before the containment check, and
  is refused outside the workspace roots.
- A read returns a content token; a write with a stale one is refused and the
  refusal carries the current token.
- Nothing opens, closes or touches an Office application the user is running.
- A form's module cannot be renamed or deleted on its own, because a form is a
  designer storage and a module that have to move together.
- Removing a loaded Power Query takes its connection and table with it.
- Deleting a module warns about the buttons that called it, because nothing
  rewrites an OnAction.
- No result is claimed that was not observed: a cell written to the package
  answers `recalculated: false`, because nothing there calculates anything.

### Known limits

- Adding or deleting a shape is not offered. It means creating a drawing part, a
  content-type override and a relationship, and for a Forms control four parts
  that have to agree, which is knowledge that belongs in pyOpenVBA rather than
  here. `xlide_set_shape_macro` covers pointing an existing one at a macro.
- A UserForm cannot be renamed or deleted, on any host but Access, for the same
  reason: nothing here can move the designer storage with the module.
- Writing cells to a legacy `.xls` is refused. Saving one means choosing a
  format, and the wrong choice drops what that format cannot hold, silently.
- A diff of a file covers its VBA, its Power Query and its sheet inventory, not
  its cell values. Both the rendered text and every comparison result say so,
  because a reader who does not know the scope takes an empty diff for an
  unchanged workbook.

[0.1.0]: https://github.com/WilliamSmithEdward/xlide_mcp/releases/tag/v0.1.0
