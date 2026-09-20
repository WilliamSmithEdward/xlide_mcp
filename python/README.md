# xlide-mcp

<!-- The MCP registry proves ownership of a PyPI package by finding this
     string in the published description. It must match `name` in the
     repository's server.json. -->
<!-- mcp-name: io.github.WilliamSmithEdward/xlide-excel-office-vba-mcp -->

**An MCP server for the code inside Office files.** The VBA, the UserForms, the
Power Query and the worksheet cells in Excel, Word, PowerPoint and Access
documents, plus Visual Basic 6 projects, reachable by any agent that speaks the
Model Context Protocol.

Reading and writing the file needs no Office installation and works on Windows,
macOS and Linux. Running macros and tests needs Windows with the desktop
application.

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

uvx takes the newest published version unless you pin it, as `xlide-mcp@1.0.0`.

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

**Files** - no Office installation, any platform.

| | |
|---|---|
| Discover | `xlide_list_projects`, `xlide_project_info`, `xlide_validate_project`, `xlide_create_project`, `xlide_doctor` |
| Catalog | `xlide_list_references`, `xlide_access_catalog` |
| Modules | `xlide_list_modules`, `xlide_read_module`, `xlide_write_module`, `xlide_rename_module`, `xlide_delete_module`, `xlide_list_procedures`, `xlide_search_modules` |
| Analysis | `xlide_analyze`, `xlide_analyze_source`, `xlide_rules` |
| Forms | `xlide_list_forms`, `xlide_read_form`, `xlide_manage_form`, `xlide_edit_form` |
| Power Query | `xlide_list_queries`, `xlide_read_query`, `xlide_write_query` |
| Cells | `xlide_list_sheets`, `xlide_read_cells`, `xlide_write_cells`, `xlide_format_cells` |
| Workbook structure | `xlide_manage_sheet`, `xlide_manage_rows_columns` |
| Tables and names | `xlide_manage_table`, `xlide_manage_name` |
| Rules and links | `xlide_manage_validation`, `xlide_manage_conditional_format`, `xlide_manage_hyperlink`, `xlide_page_setup` |
| Shapes | `xlide_list_shapes`, `xlide_set_shape_macro` |
| Source control | `xlide_export_modules`, `xlide_import_modules`, `xlide_git_changes` |

**Execution** - Windows with the desktop application.

`xlide_run_macro`, `xlide_run_vba`, `xlide_run_tests`, `xlide_compile_check`

**Live editor** - a running
[xlide_vbide](https://github.com/WilliamSmithEdward/xlide_vbide) session inside
the Visual Basic Editor.

`xlide_live_sessions`, `xlide_live_state`, `xlide_live_request`,
`xlide_live_read_module`

### Formats

| Host | VBA | Power Query | Cells |
|---|---|---|---|
| Excel | `.xlsm` `.xlsb` `.xlam` `.xls` | `.xlsx` `.xlsm` `.xlsb` `.xlam` | `.xlsx` `.xlsm` `.xlam`, and `.xlsb` `.xls` through Excel |
| Word | `.docm` `.dotm` `.doc` | - | - |
| PowerPoint | `.pptm` `.potm` | - | - |
| Access | `.accdb` `.mdb` | - | - |
| Visual Basic 6 | `.vbp` | - | - |

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

## The rules it works by

These are in the server's own instructions, so every agent that connects reads
them whether or not the user configured anything.

- The VBA inside the file is the only source of truth for it. Exported `.bas`
  and `.cls` files are copies and go stale.
- A read returns a content token. Pass it back on the write, and the write is
  refused if anything changed the module in between.
- Analysis after every change, and an error is a build failure.
- Nothing opens, closes or touches an Office application the user is running. A
  run happens in an instance the server created and can therefore terminate.
- Anything hard to undo is the user's decision: deleting a module, overwriting
  cells that hold data, writing to a project that is signed or
  password-protected.
- A cell value read from the package is what Excel last calculated. A formula
  written there has no result until Excel next opens the workbook, and the
  result says so.

## Built on

| | |
|---|---|
| [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA) | Reads and writes VBA, UserForms and Power Query inside Office files, in pure Python. |
| [pyOfficeEditor](https://github.com/WilliamSmithEdward/pyOfficeEditor) | The document surface: cells, formulas, formatting, tables, validation, rows and columns. |
| [pyVBAanalysis](https://github.com/WilliamSmithEdward/pyVBAanalysis) | The static analyzer: 119 diagnostics, measured against each host's object model. |
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
