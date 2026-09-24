# Changelog

Notable changes to `xlide-mcp`. The version lives in
`python/src/xlide_mcp/_version.py`, and a `v*.*.*` tag releases it.

## [1.1.0]

Follows pyOpenVBA 6.1.2, pyOfficeEditor 0.3.0, pyVBAanalysis 2.2.1 and
pyVBAharness 1.1.5. Fifty-eight tools, nine of them new, and 98 conformance
cases, 21 of them new.

### Fixed

- `xlide_doctor` raised for everyone with the `live` extra. pyVBAharness 1.1.3
  renamed the field it reads, and nothing in the suite called the tool. It reads
  either name now, and a test calls it.
- A write to a digitally signed project went ahead without
  `allow_invalidate_signature`, although every description said it needed it.
  pyOpenVBA drops a stale signature and warns, and a warning after the save is
  not asking first. The write is refused now, before anything is written. A
  signature Office keeps in the parts beside vbaProject.bin, which is where a
  .xlsm, .docm or .pptm keeps one, was not seen at all until pyOpenVBA 6.1, so
  such a file read as unsigned.
- A macro-enabled file saved before its first macro, which has no vbaProject.bin,
  failed every VBA tool with an error that named nothing. The listings answer
  empty and say why, a read of something named says the file has no project, and
  the document surface works as before (#1).
- A locked file named its holder as a guess, "most likely open in Excel". The
  refusal names the process Windows says holds it, and says what frees it, which
  differs: Excel locks only a workbook open for editing, Word and PowerPoint lock
  one open read-only too, and a sync client or a backup agent is neither
  (#4).
- A file that another process held without even sharing it for reading failed
  with "Error executing tool". It says who holds it.
- The live tools sent a session's token through `HTTP_PROXY` or the system proxy
  whenever one was set, because urllib proxies 127.0.0.1 like any other host.
  And a closed session whose port something else had taken failed
  `xlide_live_sessions` and `xlide_doctor` with a bare error. They talk to the
  port directly now, and a stranger on it reads as a session that has closed.
- A write to a password-protected Access project failed with "Error executing
  tool" rather than the refusal every other host gives, because Access refuses
  with an error of its own. It names `allow_protected` now.

### Added

- The first module, form or import into a .xlsm, .docm or .pptm with no VBA
  project gives it the project its application makes for a first macro: in
  Excel a document module for the workbook and one per sheet.
  `xlide_create_project` gives an existing .docm the empty project Word keeps,
  and says why Excel and PowerPoint take none: neither writes a project until
  it holds code.
- `xlide_manage_reference` adds and removes type library references. Excel,
  Word, PowerPoint and Access are known by name anywhere; on Windows any library
  the VBA editor's References dialog lists resolves by its description or by the
  name code writes, with its GUID, version and path read from the registry
  rather than recalled. The host's own library and Microsoft Forms while a
  UserForm exists are refused (#2).
- The analyzer is given the project's references, so `Dim doc As Word.Document`
  in a workbook with no reference to Word is the error it is in Excel, and code
  that does reference Word is checked against Word's model too.
- `xlide_manage_shape` adds a Forms control, an AutoShape, a text box, a line or
  a picture, placed at a cell or in points, and removes one with every part it
  has. `xlide_list_shapes` reports each shape's position in points, what a
  control holds, and a chart's series (#2).
- `xlide_add_chart` charts a block of cells, as Excel's Insert Chart does.
- `xlide_read_cells` takes `calculate`, which works every formula out with
  pyOfficeEditor's formula engine, in memory, and names each cell it could not
  work out with the reason; `include='text'`, what each cell shows under its
  number format; and `include='rich_text'`, the font runs of text in several
  fonts, which `xlide_write_cells` now writes.
- `xlide_evaluate_formula` works out what a formula would give in a cell,
  without writing it.
- `xlide_manage_filter` sets, reapplies and clears autofilters on a sheet or a
  table, with criteria spelled as VBA's Range.AutoFilter spells them, and hides
  the rows they filter out: Excel does not apply a filter when it opens a
  workbook, it shows the rows as the file marks them.
- `xlide_manage_comment` lists, writes, answers, resolves and removes notes and
  threaded comments.
- `xlide_format_cells` applies a named cell style, and `xlide_list_sheets`
  reports pivot tables and chart sheets.
- `xlide_is_open`, `xlide_open_in_app` and `xlide_close_in_app`: whether a file
  is open in its Office application, in any running instance, read-only or for
  editing, with unsaved work or not; opening it for the user, read-only or in a
  new instance, behind their other windows if asked; and closing it. A copy
  with unsaved work is closed only with `save_changes` or `discard_changes`,
  and `end_process` ends the process holding a file that will not let go, only
  when that process is the file's own application. Where the user was in a
  file this closed is where it reopens. Opening follows the user's Trust Center
  macro settings, as a file they open themselves does (#4, #5).
- Every write tells a running XLIDE for VS Code what changed, and a module write
  carries its before and after, so XLIDE can mark it as an agent edit with a diff
  and Keep and Revert. docs/xlide-vscode-bridge.md is the protocol; with no XLIDE
  listening nothing changes (#3).

### Changed

- Every shape write goes through pyOfficeEditor, which keeps a Forms control's
  four parts in agreement, apart from the macro on a control Excel 2007 saved:
  that control lives only in VML, where pyOfficeEditor does not look, and its
  macro is still written here. `xlide_set_shape_macro` answers as before, and
  every shape conformance case passed before and after the swap. The reader
  stays until pyOfficeEditor's shapes carry the cells an anchor covers, alt
  text, the hidden flag and ActiveX controls.
- A form created here references Microsoft Forms, which pyOpenVBA 6.1.2 adds
  with a first UserForm as the VBA editor does; without it a form's own event
  procedures did not compile. The result lists the references added.
- `xlide_analyze_source` with a `.vbp` file_path analyzed against no host
  rather than refusing the host it named.
- Writes land in place, as before: the same file, its size and modification time
  changed and nothing left beside it, which a VS Code file watcher reports as a
  change (#1). A test now holds them to it.

### Known limits

- Word and PowerPoint shapes are not reached: pyOfficeEditor's document surface
  is Excel's, and its Word and PowerPoint surfaces come later.
- A shape cannot be moved, resized or renamed in place. A chart, and a group
  holding a picture, a chart or a control, cannot be removed: pyOfficeEditor
  leaves their parts behind (WilliamSmithEdward/pyOfficeEditor#3,
  WilliamSmithEdward/pyOfficeEditor#5). Nor can a member of a group, or a
  control Excel 2007 saved, which pyOfficeEditor does not reach; that control
  still takes a macro.
- The XLIDE side of the bridge is not in a released XLIDE yet:
  WilliamSmithEdward/xlide_vscode#91.

## [1.0.3]

Copy and packaging. No behaviour change: the tool surface, the conformance
corpus and the contract digest are identical to 1.0.0.

- The registry description drops the enumeration that listed cells beside the
  VBA and says "the document around it" instead, matching the READMEs. Directory
  sites copy that line when they ingest the listing, and one had already taken
  the older wording, so leaving it would have spread further with every site
  that picked it up next.
- Roots can be given positionally as well as with `--root`. A launcher that
  mounts the caller's folders somewhere of its own choosing appends them as
  plain arguments and cannot repeat a flag in front of each, which made
  `--root /a /b` exit on the second path rather than start.
- The Dockerfile splits the root out of the entrypoint and into `CMD`, so a
  runner passing its own arguments replaces it. A plain `docker run` is
  unchanged.
- More PyPI keywords: macro, xlsm, vb6, spreadsheet.

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

[1.1.0]: https://github.com/WilliamSmithEdward/xlide_mcp/releases/tag/v1.1.0
[1.0.3]: https://github.com/WilliamSmithEdward/xlide_mcp/releases/tag/v1.0.3
[1.0.2]: https://github.com/WilliamSmithEdward/xlide_mcp/releases/tag/v1.0.2
[1.0.1]: https://github.com/WilliamSmithEdward/xlide_mcp/releases/tag/v1.0.1
[1.0.0]: https://github.com/WilliamSmithEdward/xlide_mcp/releases/tag/v1.0.0
