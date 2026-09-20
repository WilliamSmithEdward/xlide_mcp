# Changelog

Notable changes to `xlide-mcp`. The version lives in
`python/src/xlide_mcp/_version.py`, and a `v*.*.*` tag releases it.

## [1.0.2]

A findable name in the MCP registry. No behaviour change: the tool surface, the
conformance corpus and the contract digest are identical to 1.0.0.

- The registry listing is now
  `io.github.WilliamSmithEdward/xlide-excel-office-vba-mcp`. Registry search
  matches the name and nothing else, so the 1.0.1 listing was invisible to
  anyone searching for excel, office or vba, which is everyone who wants this.
  Measured before changing it: `vba` returned two servers, both matching on
  substrings of somebody's username, and neither was this one.
- The `server.json` name and the `mcp-name` marker in the PyPI description move
  together, because the registry proves ownership by finding the second inside
  the first's package.

The 1.0.1 listing under `xlide-mcp` stays where it is. The registry cannot
unpublish a server yet, so it remains, frozen at 1.0.1, pointing at the same
package.

## [1.0.1]

Registry metadata. No behaviour change: the tool surface, the conformance corpus
and the contract digest are identical to 1.0.0.

- Listed in the [MCP registry](https://registry.modelcontextprotocol.io) as
  `io.github.WilliamSmithEdward/xlide-mcp`, which is the index the major clients
  and most third-party directories read from.
- `server.json` describes the package for that listing, and the PyPI description
  carries the `mcp-name` marker the registry proves ownership with.
- Pushing a tag now registers the release as well as publishing it, authenticated
  by OIDC, so there is no token to store. The registry step runs after PyPI and
  waits for the new version to be served, because ownership is proved by reading
  the published description.
- A test keeps the three copies of that name and version in step. Getting them
  out of step fails the registry publish on a release already gone to PyPI,
  which cannot be taken back.

## [1.0.0]

First release. An MCP server for the VBA, UserForms, Power Query, worksheet
cells and document surface inside Office files, and for Visual Basic 6 projects.

### Added

- Forty-nine tools over three layers. The file layer needs no Office
  installation and runs anywhere; the execution layer runs macros, tests and
  compile checks in a desktop application the server owns and holds a deadline
  over; the live layer reads a running `xlide_vbide` session inside the Visual
  Basic Editor.
- Excel, Word, PowerPoint and Access, plus `.vbp` projects read as projects so
  their modules analyze together.
- The VBA project: read, write, rename, delete, search, list procedures, and a
  static analyzer with 119 diagnostics measured against each host's own object
  model.
- The document surface, through pyOfficeEditor: cells and formulas, fonts,
  fills, borders, alignment, number formats and merging; sheets added, removed,
  renamed, moved, hidden and protected; rows and columns inserted, deleted,
  sized, hidden and grouped, with every reference in the workbook following or
  breaking exactly as Excel breaks it; tables, defined names, data validation,
  conditional formatting, hyperlinks, and how a sheet prints.
- Worksheet cells read and written in the OOXML package, and, for `.xlsb` and
  `.xls`, through Excel, with `source` and `recalculated` on every result saying
  which answered.
- Power Query beside the VBA: read, write, rename, remove, and loaded onto a
  sheet or taken back off.
- What changed inside a file, three ways. A write reports a unified diff of what
  the file now holds, taken from the read-back rather than from what was sent.
  `xlide_git_changes` compares against any git revision, one entry per module
  and per query. `xlide-mcp --textconv` is a git textconv driver, so `git diff`,
  `git show` and `git log -p` render a workbook as its VBA and M instead of
  reporting that two binaries differ.
- A generated contract under `contract/`: the tool surface, and 77 conformance
  cases pinning what the answers mean. Every implementation in the repository is
  verified against them, and the Python suite runs them too, so a case that is
  wrong fails before a port is built on it.

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
- Formatting changes only what was asked for, so making a header row bold does
  not flatten the number formats under it.
- Hiding the last visible sheet is refused, because Excel will not open the
  result.
- A conditional rule with no paint is refused, because Excel stores one happily
  and it highlights nothing.
- No result is claimed that was not observed: a cell written to the package
  answers `recalculated: false`, because nothing there calculates anything.

### Known limits

- Adding or deleting a shape is not offered. `xlide_set_shape_macro` covers
  pointing an existing one at a macro.
- A UserForm cannot be renamed or deleted, on any host but Access, because
  nothing here can move the designer storage with the module.
- Writing cells to a legacy `.xls` is refused. Saving one means choosing a
  format, and the wrong choice drops what that format cannot hold, silently.
- A diff of a file covers its VBA, its Power Query and its sheet inventory, not
  its cell values. Both the rendered text and every comparison result say so,
  because a reader who does not know the scope takes an empty diff for an
  unchanged workbook.
- The document surface is Excel only. Word, PowerPoint and Access reach their
  VBA, forms and catalogs, and their document surfaces follow pyOfficeEditor.

[1.0.2]: https://github.com/WilliamSmithEdward/xlide_mcp/releases/tag/v1.0.2
[1.0.1]: https://github.com/WilliamSmithEdward/xlide_mcp/releases/tag/v1.0.1
[1.0.0]: https://github.com/WilliamSmithEdward/xlide_mcp/releases/tag/v1.0.0
