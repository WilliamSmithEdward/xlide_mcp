# xlide-mcp

<!-- The MCP registry proves ownership of a PyPI package by finding this
     string in the published description. It must match `name` in the
     repository's server.json. -->
<!-- mcp-name: io.github.WilliamSmithEdward/xlide-excel-office-vba-mcp -->

**An MCP server for the inside of an Office file.** Read, write, analyze and
test the VBA in Excel, Word, PowerPoint and Access, and edit the document around
it. Visual Basic 6 projects open the same way.

VBA, Power Query and OOXML worksheet edits need no Office installation and run on
Windows, macOS and Linux. Reading `.xlsb` and `.xls` cells, writing `.xlsb` cells,
and running macros or tests need Windows with the desktop application.

```text
xlide_list_projects                -> Budget.xlsm
xlide_project_info   Budget.xlsm   -> 4 modules, 2 sheets, 1 query, not signed
xlide_read_module    Helpers       -> the source, and a content token
xlide_write_module   Helpers       -> guarded by that token
xlide_analyze        Budget.xlsm   -> 0 errors, 2 warnings
xlide_run_tests      Budget.xlsm   -> 12 passed
```

## Why

An agent asked to fix a macro works from a copied snippet with no idea what else
is in the project, or asks the user to export the modules and paste them back
afterwards. Both treat the Office file as opaque. The VBA project, the form
designs and the M code are all readable and writable without opening the
application at all.

## Install

```bash
pip install xlide-mcp              # reads and writes files, any platform
pip install "xlide-mcp[live]"      # adds running macros and tests, Windows
```

Or install nothing and run it through [uv](https://github.com/astral-sh/uv),
which fetches the package on first use and brings its own Python:

```bash
uvx --from "xlide-mcp[live]" xlide-mcp --root /path/to/your/files
```

The `live` extra is safe to ask for on every platform: what it pulls in is marked
`sys_platform == 'win32'`, so off Windows it resolves to nothing.

## Run

```bash
xlide-mcp --root /path/to/your/files
```

Point a client at it. For Claude Desktop, Claude Code, or any client that
launches a server over stdio:

```json
{
  "mcpServers": {
    "xlide": {
      "command": "xlide-mcp",
      "args": ["--root", "/path/to/your/files"]
    }
  }
}
```

For VS Code, put this in `.vscode/mcp.json` at the workspace root. VS Code uses
`servers` here; the portable `.mcp.json` format above uses `mcpServers`.

```json
{
  "servers": {
    "xlide": {
      "type": "stdio",
      "command": "uvx",
      "args": ["--from", "xlide-mcp[live]", "xlide-mcp", "--root", "${workspaceFolder}"]
    }
  }
}
```

Run `MCP: List Servers` in VS Code to start `xlide` and inspect its output if it
does not register. The server prints its version and allowed roots to stderr on
startup. See [VS Code's MCP guide](https://code.visualstudio.com/docs/agent-customization/mcp-servers)
for the configuration format and commands.

With uv instead, so that nothing has to be installed first:

```json
{
  "mcpServers": {
    "xlide": {
      "command": "uvx",
      "args": [
        "--from", "xlide-mcp[live]",
        "xlide-mcp", "--root", "/path/to/your/files"
      ]
    }
  }
}
```

uvx takes the newest published version unless you pin it, as `xlide-mcp@1.1.0`.

`--root` is the security boundary. Every path a tool accepts is resolved,
symlinks included, and refused unless it lands inside a root.

| Setting | Flag | Environment |
|---|---|---|
| Workspace roots | `--root` (repeatable) | `XLIDE_MCP_ROOTS` |
| Refuse every write | `--read-only` | `XLIDE_MCP_READ_ONLY` |
| Allow any absolute path | `--allow-outside-roots` | `XLIDE_MCP_ALLOW_OUTSIDE_ROOTS` |
| Default run deadline | `--timeout` | `XLIDE_MCP_TIMEOUT` |

`--transport streamable-http --port 8765` serves HTTP instead of stdio. It binds
to loopback unless told otherwise: this server reads and writes files, and a
default that listened on every interface would hand that reach to the network.

Call `xlide_doctor` first from a new client. It reports the workspace roots,
which layers are installed, which Office applications the machine has, and
whether each has the Trust Center setting that module injection needs.

## What it does

**Files** - the VBA project and the document it lives in. Most file tools need no
Office installation; binary Excel grids use Excel on Windows.

*The code project*

| | |
|---|---|
| Discover | `xlide_list_projects`, `xlide_project_info`, `xlide_validate_project`, `xlide_create_project`, `xlide_doctor` |
| Modules | `xlide_list_modules`, `xlide_read_module`, `xlide_write_module`, `xlide_edit_module`, `xlide_rename_module`, `xlide_delete_module`, `xlide_list_procedures`, `xlide_search_modules` |
| Analysis | `xlide_analyze`, `xlide_analyze_source`, `xlide_rules` |
| Forms | `xlide_list_forms`, `xlide_read_form`, `xlide_manage_form`, `xlide_edit_form` |
| References and catalog | `xlide_list_references`, `xlide_manage_reference`, `xlide_access_catalog`, `xlide_read_access_query` |
| Source control | `xlide_export_modules`, `xlide_import_modules`, `xlide_git_changes` |

`xlide_edit_module` accepts `preview_only=true` to check a batch of line edits and
return its proposed diff without saving. The preview's content token can be used
to apply the same edits if the module has not changed.
`xlide_search_modules` returns tokens for matched modules, so its line numbers
can also be passed directly to a guarded edit.
Broad searches can be paged with `offset` and `next_offset`; `total_match_count`
reports how many matching lines exist across all pages.
`xlide_list_procedures` returns a content token for guarded edits at the listed
lines. Reading an empty module returns its empty body and token.
Writing an identical module or applying an unchanged module import skips the save;
the response reports `saved=false`.
Applying an unchanged export leaves the existing `.bas` and `.cls` files untouched.
Long Power Query formulas can be read by line; the read returns a token that
`xlide_write_query` can use to refuse a stale change.
Large form designs can be read in pages with `offset` and `next_offset`.
Office file, VBA module and Power Query lists also return `next_offset` when a
later page is available.
For crowded drawing layers, pass `sheet` to `xlide_list_shapes` and follow its
`next_offset` to read later shapes.
Worksheet lists mark unreadable visibility or pivot metadata as unknown, and formatted or rich
text cell reads name the cell when decoding fails instead of returning a blank.
An Excel-backed worksheet survey refuses malformed output instead of dropping sheets.
An Excel-backed cell read refuses a partial grid instead of labeling it as the full range.
Form summary lists use the same `offset` and `next_offset` paging.
Form reads distinguish missing properties or sections from a read failure, and
an unreadable Access report collection fails instead of appearing empty.
Procedure lists also return `next_offset` for long modules.
Structural validation problems can be read with `offset` and `next_offset`.
Access catalog lists return a `next_offsets` entry for each included collection.
Long saved query SQL is previewed in the catalog; `xlide_read_access_query`
returns the full text in character pages with a content token.
The `list` actions for tables, names, validation, conditional formats, hyperlinks
and comments also accept `offset` and `max_results` and return `next_offset`.
Merging cells keeps only the top-left value. `xlide_format_cells` refuses to clear
other values or formulas unless `allow_overwrite=true` is passed.

*The document around it*

| | |
|---|---|
| Power Query | `xlide_list_queries`, `xlide_read_query`, `xlide_write_query` |
| Sheets and cells | `xlide_list_sheets`, `xlide_read_cells`, `xlide_write_cells`, `xlide_evaluate_formula`, `xlide_format_cells` |
| Structure | `xlide_manage_sheet`, `xlide_manage_rows_columns` |
| Tables and names | `xlide_manage_table`, `xlide_manage_name`, `xlide_manage_filter` |
| Rules, links and notes | `xlide_manage_validation`, `xlide_manage_conditional_format`, `xlide_manage_hyperlink`, `xlide_manage_comment`, `xlide_page_setup` |
| Shapes and charts | `xlide_list_shapes`, `xlide_manage_shape`, `xlide_set_shape_macro`, `xlide_add_chart` |

**Execution** - Windows with the desktop application.

`xlide_run_macro`, `xlide_run_vba`, `xlide_run_tests`, `xlide_compile_check`

**Your Office applications** - Windows with the desktop application, and only on
the file named.

`xlide_is_open`, `xlide_open_in_app`, `xlide_close_in_app`

**Live editor** - a running
[xlide_vbide](https://github.com/WilliamSmithEdward/xlide_vbide) session inside
the Visual Basic Editor.

`xlide_live_sessions`, `xlide_live_state`, `xlide_live_request`,
`xlide_live_read_module`

`xlide_live_read_module` accepts line ranges for long unsaved source and returns
a content token for comparing that source with the saved module.
Both analysis tools return `next_offset` when findings exceed a page, while
keeping their full severity counts.

### Formats

| Host | Extensions |
|---|---|
| Excel | `.xlsm` `.xlsb` `.xlam` `.xls`, and `.xlsx` for everything except VBA |
| Word | `.docm` `.dotm` `.doc` |
| PowerPoint | `.pptm` `.potm` |
| Access | `.accdb` `.mdb` |
| Visual Basic 6 | `.vbp` |

VBA reads and writes in all of them. Power Query is available in `.xlsx`, `.xlsm`,
`.xlsb` and `.xlam`; package-based worksheet edits are available in `.xlsx`,
`.xlsm` and `.xlam`. A `.xlsb` keeps its grid in binary records and a `.xls`
inside a compound file; on Windows with Excel, both can be read for values, and
`.xlsb` cells can be written. This server does not write `.xls` cells.

A recognized extension outside those sets is listed with the reason it cannot be
opened. Nothing drops out of a listing without saying why.

## Seeing what changed

An Office file is one binary blob to git, so a commit that changed a line of VBA
and one that replaced the whole project are the same three words: `Binary files
differ`.

`xlide_git_changes` reports what changed since any revision, one entry per module
and query, each with a unified diff. `xlide-mcp --textconv` is a git textconv
driver that does the same for git itself:

```bash
echo '*.xlsm binary diff=vba' >> .gitattributes
git config diff.vba.textconv "xlide-mcp --textconv"
git config diff.vba.cachetextconv true
```

```diff
 Public Sub Greet()
-    MsgBox "hello"
+    MsgBox "hello, world"
+    Debug.Print Now
 End Sub
```

That is `git diff` on a `.xlsm`, and `git show` and `git log -p` convert too.
Write `binary diff=vba`, not `diff=vba` alone: the `binary` macro is
`-diff -merge -text` and the later `diff=vba` overrides only its `-diff`, so the
file keeps `-text` and git never applies end-of-line conversion to a container it
would corrupt.

Both routes cover VBA, Power Query and the sheet inventory. Cell values are not
included, and the rendered text says so on its first line.

Inside VS Code beside XLIDE, every write also tells a running XLIDE what
changed, so a module an agent wrote is marked in XLIDE's project tree with its
diff and Keep and Revert, as XLIDE's own agent edits are. It needs an XLIDE that
listens for this; the protocol is
[docs/xlide-vscode-bridge.md](https://github.com/WilliamSmithEdward/xlide_mcp/blob/main/docs/xlide-vscode-bridge.md).

## The rules it works by

These are in the server's own instructions, so every agent that connects reads
them whether or not the user configured anything.

- The VBA inside the file is the only source of truth for it. Exported `.bas`
  and `.cls` files are copies and go stale.
- A read returns a content token. Pass it back on the write, and the write is
  refused if anything changed the module in between.
- Analysis after every change, and an error is a build failure.
- A run happens in an Office instance the server created and can therefore
  terminate. The user's own applications are touched only through the three
  tools above, on the file named: unsaved work is closed only when the call
  says to save or discard it, and a process is ended only when asked.
- Anything hard to undo is the user's decision: deleting a module, overwriting
  cells that hold data, writing to a project that is signed or
  password-protected. Cell overwrites and writes to signed or protected projects
  are refused until the call carries the flag
  that allows them.
- A cell value read from the package is what Excel last calculated. A formula
  written there has no result in the file until Excel next opens it.
  `calculate=true` works results out with pyOfficeEditor's formula engine, which
  names any cell it could not.

## Built on

| | |
|---|---|
| [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA) | Reads and writes VBA, UserForms and Power Query inside Office files, in pure Python. |
| [pyOfficeEditor](https://github.com/WilliamSmithEdward/pyOfficeEditor) | The document surface: cells, formulas, formatting, tables, validation, rows and columns. |
| [pyVBAanalysis](https://github.com/WilliamSmithEdward/pyVBAanalysis) | The static analyzer: 131 diagnostics, measured against each host's object model. |
| [pyVBAharness](https://github.com/WilliamSmithEdward/pyVBAharness) | Runs VBA in desktop Office under a supervisor that enforces a deadline. |
| [XLIDE for VS Code](https://github.com/WilliamSmithEdward/xlide_vscode) | Where the tool surface, the content-token guard and the agent instructions come from. |

## Working on it

This package is the reference implementation in a repository that will hold
others. See
[AGENTS.md](https://github.com/WilliamSmithEdward/xlide_mcp/blob/main/AGENTS.md)
for how to change a tool,
[contract/](https://github.com/WilliamSmithEdward/xlide_mcp/blob/main/contract/README.md)
for the generated tool surface and conformance corpus every implementation is
verified against, and
[docs/porting.md](https://github.com/WilliamSmithEdward/xlide_mcp/blob/main/docs/porting.md)
for building one in another language.

```bash
pip install -e ".[dev,live]"
python -m pytest                 # the file layer, no Office needed
python -m pytest -m live         # the rest, real Office, Windows only
python -m ruff check src tests tools
```

## License

[MIT](https://github.com/WilliamSmithEdward/xlide_mcp/blob/main/LICENSE).
