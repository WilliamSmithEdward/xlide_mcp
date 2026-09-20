# xlide-mcp

**An MCP server for the code inside Office files.** The VBA, the UserForms, the
Power Query and the worksheet cells in Excel, Word, PowerPoint and Access
documents, plus Visual Basic 6 projects, reachable by any agent that speaks the
Model Context Protocol.

Reading and writing the file needs no Office installation and works on Windows,
macOS and Linux. Running macros and tests needs Windows with the desktop
application.

```
xlide_list_projects                -> Budget.xlsm
xlide_project_info   Budget.xlsm   -> 4 modules, 2 sheets, 1 query, not signed
xlide_read_module    Helpers       -> the source, and a content token
xlide_write_module   Helpers       -> guarded by that token
xlide_analyze        Budget.xlsm   -> 0 errors, 2 warnings
xlide_run_tests      Budget.xlsm   -> 12 passed
```

## Why

An agent asked to fix a macro has, until now, had two bad options: work from a
copied snippet with no idea what else is in the project, or ask the user to
export the modules and paste them back afterwards. Both treat the Office file as
opaque. It is not: the VBA project, the form designs and the M code are all
readable and writable without opening the application at all.

This server is the surface that makes that reachable, and it sets the rules that
stop it being dangerous: a write is refused if the module changed since it was
read, analysis is a build gate rather than a suggestion, a run happens in an
application the server owns and holds a deadline over, and everything that cannot
be undone is the user's decision rather than the agent's.

## Implementations

| Directory | Language | Status |
|---|---|---|
| [`python/`](python/) | Python 3.10+ | Reference implementation |

Other languages go in sibling directories. Python is normative: it tracks the
upstream libraries, and ports track it. See [docs/porting.md](docs/porting.md).

## Install and run

```bash
cd python
pip install -e ".[live]"          # omit [live] off Windows
xlide-mcp --root /path/to/your/files
```

Point a client at it. For Claude Desktop, Claude Code or any client that launches
a server over stdio:

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

`--root` is the security boundary. Every path a tool accepts is resolved,
symlinks included, and refused unless it lands inside a root. Add `--read-only`
to allow reads and analysis and refuse every write.

Call `xlide_doctor` first from a new client: it reports the workspace roots, which
layers are installed, which Office applications this machine has, and whether
each one has the Trust Center setting that module injection needs.

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
| Cells | `xlide_list_sheets`, `xlide_read_cells`, `xlide_write_cells` |
| Shapes | `xlide_list_shapes`, `xlide_set_shape_macro` |
| Source control | `xlide_export_modules`, `xlide_import_modules`, `xlide_git_changes` |

**Execution** - Windows with the desktop application.

`xlide_run_macro`, `xlide_run_vba`, `xlide_run_tests`, `xlide_compile_check`

**Live editor** - a running [xlide_vbide](https://github.com/WilliamSmithEdward/xlide_vbide)
session inside the Visual Basic Editor.

`xlide_live_sessions`, `xlide_live_state`, `xlide_live_request`, `xlide_live_read_module`

### Formats

| Host | VBA | Power Query | Cells |
|---|---|---|---|
| Excel | `.xlsm` `.xlsb` `.xlam` `.xls` | `.xlsx` `.xlsm` `.xlsb` `.xlam` | `.xlsx` `.xlsm` `.xlam` |
| Word | `.docm` `.dotm` `.doc` | - | - |
| PowerPoint | `.pptm` `.potm` | - | - |
| Access | `.accdb` `.mdb` | - | - |
| Visual Basic 6 | `.vbp` | - | - |

A recognized extension outside those sets is listed with the reason it cannot be
opened, rather than left silently out of a listing.

## The rules it works by

These are in the server's own instructions, so every agent that connects reads
them whether or not the user configured anything.

- The VBA inside the file is the only source of truth for it. Exported `.bas` and
  `.cls` files are copies and go stale.
- A read returns a content token. Pass it back on the write, and the write is
  refused if anything changed the module in between.
- Analysis after every change, and an error is a build failure.
- Nothing opens, closes or touches an Office application the user is running. A
  run happens in an instance the server created and can therefore terminate.
- Anything hard to undo is the user's decision: deleting a module, overwriting
  cells that hold data, writing to a project that is signed or password-protected.
- A cell value is what Excel last calculated. A formula written here has no result
  until Excel next opens the workbook, and the result says so.

## Built on

| | |
|---|---|
| [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA) | Reads and writes VBA, UserForms and Power Query inside Office files, in pure Python. |
| [pyVBAanalysis](https://github.com/WilliamSmithEdward/pyVBAanalysis) | The static analyzer: 119 diagnostics, measured against each host's object model. |
| [pyVBAharness](https://github.com/WilliamSmithEdward/pyVBAharness) | Runs VBA in desktop Office under a supervisor that enforces a deadline. |
| [XLIDE for VS Code](https://github.com/WilliamSmithEdward/xlide_vscode) | Where the tool surface, the content-token guard and the agent instructions come from. |
| [xlide for the VBE](https://github.com/WilliamSmithEdward/xlide_vbide) | The live editor session the `xlide_live_*` tools talk to. |

## Contributing

```bash
cd python
pip install -e ".[dev,live]"
python -m pytest              # 229 tests, no Office needed
python -m pytest -m live      # 12 more, real Office, Windows only
python -m ruff check src tests tools
```

The generated contract under [`contract/`](contract/) must stay current; the test
suite fails if it does not. After changing a tool:

```bash
python tools/export_contract.py
python tools/export_conformance.py
```

## License

MIT.
