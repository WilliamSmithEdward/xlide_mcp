# xlide-mcp

[![xlide-excel-word-powerpoint-access-office-vba-mcp MCP server – quality and maintenance score on Glama](https://glama.ai/mcp/servers/WilliamSmithEdward/xlide_mcp/badges/card.svg)](https://glama.ai/mcp/servers/WilliamSmithEdward/xlide_mcp)

[![Python Version](https://img.shields.io/pypi/pyversions/xlide-mcp.svg)](https://pypi.org/project/xlide-mcp/)
[![CI](https://github.com/WilliamSmithEdward/xlide_mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/WilliamSmithEdward/xlide_mcp/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/WilliamSmithEdward/xlide_mcp/blob/main/LICENSE)
[![Downloads](https://static.pepy.tech/badge/xlide-mcp/month)](https://pepy.tech/project/xlide-mcp)

**An MCP server for the inside of an Office file.** Read, write, analyze and
test the VBA in Excel, Word, PowerPoint and Access, and edit the document around
it. Visual Basic 6 projects open the same way.

Reading and writing needs no Office installation and runs on Windows, macOS and
Linux. Running macros and tests needs Windows with the desktop application.

```
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

This server makes that reachable, and bounds it. A write is refused if the
module changed since it was read. An analysis error fails the change. A run
happens in an application the server created and holds a deadline over. Anything
that cannot be undone is the user's decision.

## Implementations

| Directory | Language | Status |
|---|---|---|
| [`python/`](python/) | Python 3.10+ | Reference implementation |

Other languages go in sibling directories. Python is normative: it tracks the
upstream libraries, and ports track it. See [docs/porting.md](docs/porting.md).

## Install and run

```bash
pip install xlide-mcp              # reads and writes files, any platform
pip install "xlide-mcp[live]"      # adds running macros and tests, Windows
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

Or install nothing and let [uv](https://github.com/astral-sh/uv) fetch it on
first run. uv is one binary and installs its own Python, so a machine with
neither can still run this:

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

The `live` extra is safe to ask for on every platform: what it pulls in is marked
`sys_platform == 'win32'`, so off Windows it resolves to nothing and the same
configuration works everywhere. Drop the `--from` pair for the file layer alone.
uvx takes the newest published version unless you pin it, as `xlide-mcp@1.1.0`.

There is a `Dockerfile` for the file layer, which is the part that needs no
Office installation. Mount the folder holding the files at `/workspace`, because
that is the root the container is bounded to:

```bash
docker build -t xlide-mcp .
docker run --rm -i -v "$PWD:/workspace" xlide-mcp
```

`--root` is the security boundary. Every path a tool accepts is resolved,
symlinks included, and refused unless it lands inside a root. Add `--read-only`
to allow reads and analysis and refuse every write.

Call `xlide_doctor` first from a new client: it reports the workspace roots, which
layers are installed, which Office applications this machine has, and whether
each one has the Trust Center setting that module injection needs.

## What it does

**Files** - no Office installation, any platform. Two halves: the code project,
and the document it lives in.

*The code project*

| | |
|---|---|
| Discover | `xlide_list_projects`, `xlide_project_info`, `xlide_validate_project`, `xlide_create_project`, `xlide_doctor` |
| Modules | `xlide_list_modules`, `xlide_read_module`, `xlide_write_module`, `xlide_rename_module`, `xlide_delete_module`, `xlide_list_procedures`, `xlide_search_modules` |
| Analysis | `xlide_analyze`, `xlide_analyze_source`, `xlide_rules` |
| Forms | `xlide_list_forms`, `xlide_read_form`, `xlide_manage_form`, `xlide_edit_form` |
| References and catalog | `xlide_list_references`, `xlide_manage_reference`, `xlide_access_catalog` |
| Source control | `xlide_export_modules`, `xlide_import_modules`, `xlide_git_changes` |

*The document around it*

| | |
|---|---|
| Power Query | `xlide_list_queries`, `xlide_read_query`, `xlide_write_query` (set, rename, remove, load, unload) |
| Sheets and cells | `xlide_list_sheets`, `xlide_read_cells`, `xlide_write_cells`, `xlide_evaluate_formula`, `xlide_format_cells` |
| Structure | `xlide_manage_sheet`, `xlide_manage_rows_columns` |
| Tables and names | `xlide_manage_table`, `xlide_manage_name`, `xlide_manage_filter` |
| Rules, links and notes | `xlide_manage_validation`, `xlide_manage_conditional_format`, `xlide_manage_hyperlink`, `xlide_manage_comment`, `xlide_page_setup` |
| Shapes and charts | `xlide_list_shapes`, `xlide_manage_shape`, `xlide_set_shape_macro`, `xlide_add_chart` |

**Execution** - Windows with the desktop application.

`xlide_run_macro`, `xlide_run_vba`, `xlide_run_tests`, `xlide_compile_check`

**Your Office applications** - Windows with the desktop application. The one
place this server acts inside an application the user is running, and only on
the file named.

`xlide_is_open`, `xlide_open_in_app`, `xlide_close_in_app`

**Live editor** - a running [xlide_vbide](https://github.com/WilliamSmithEdward/xlide_vbide)
session inside the Visual Basic Editor.

`xlide_live_sessions`, `xlide_live_state`, `xlide_live_request`, `xlide_live_read_module`

### Formats

| Host | Extensions |
|---|---|
| Excel | `.xlsm` `.xlsb` `.xlam` `.xls`, and `.xlsx` for everything except VBA |
| Word | `.docm` `.dotm` `.doc` |
| PowerPoint | `.pptm` `.potm` |
| Access | `.accdb` `.mdb` |
| Visual Basic 6 | `.vbp` |

VBA reads and writes in all of them. Power Query and the document surface are
Excel's, and they live in the OOXML package, so they come from `.xlsx`, `.xlsm`
and `.xlam`. A `.xlsx` has no VBA project by design and is listed anyway, since
its queries and its sheets are fully reachable.

A recognized extension outside those sets is listed with the reason it cannot be
opened. Nothing drops out of a listing without saying why.

Excel, Word and PowerPoint write no VBA project into a macro-enabled file until
its first macro exists, so a `.xlsm` nobody has written code in yet has none.
That is an ordinary file: the listings answer empty, and the first module
written gives it the project its application would have made.

A `.xlsb` keeps its grid in binary records and a `.xls` inside a compound file,
neither of them OOXML. On Windows with Excel, those go through Excel, and the
result says `source: excel` and `recalculated: true`, because opening the
workbook is what produced the values.

## Seeing what changed

An Office file is one binary blob to git, so a commit that changed a line of VBA
and one that replaced the whole project are the same three words: `Binary files
differ`. Two ways out, both reading the file through the same renderer.

`xlide_git_changes` reports what changed since any revision, one entry per module
and query, each with a unified diff. Call it before committing, or to review what
an agent just did.

`xlide-mcp --textconv` is a git textconv driver. Wire it up once and the file
diffs as text everywhere git looks:

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

That is `git diff` on a `.xlsm`. `git show` and `git log -p` convert too.

Inside VS Code beside [XLIDE](https://github.com/WilliamSmithEdward/xlide_vscode),
there is a third way. Every write tells a running XLIDE what changed, and a
module write carries its before and after, so XLIDE marks the module as an agent
edit in its project tree, opens the diff, and offers Keep and Revert, as it does
for its own agent tools. It needs an XLIDE that listens for this;
[docs/xlide-vscode-bridge.md](docs/xlide-vscode-bridge.md) is the protocol. With
no XLIDE running, nothing changes.

`binary diff=vba` rather than `diff=vba` alone. The `binary` macro means
`-diff -merge -text`, and the later `diff=vba` overrides only its `-diff`, so the
file keeps `-text` and git never applies end-of-line conversion to a container it
would corrupt.

In VS Code, any side of a diff that comes out of git history renders through the
driver, because the git extension reads blobs with `git show --textconv`.
Comparing two revisions of a workbook therefore shows VBA. The Source Control
panel's working-tree diff does not: the right-hand side there is the file on
disk, still binary.

Both routes render VBA, Power Query and the sheet inventory. Cell values are not
included, and the first line of every rendered file says so, because a reader
who does not know the scope takes an empty diff for an unchanged workbook. git
still stores the blob either way, so merges stay binary.

## The rules it works by

These are in the server's own instructions, so every agent that connects reads
them whether or not the user configured anything.

- The VBA inside the file is the only source of truth for it. Exported `.bas` and
  `.cls` files are copies and go stale.
- A read returns a content token. Pass it back on the write, and the write is
  refused if anything changed the module in between.
- Analysis after every change, and an error is a build failure.
- A run happens in an Office instance the server created and can therefore
  terminate. The user's own applications are touched only through the three
  tools above, on the file named: unsaved work is closed only when the call says
  to save or discard it, and a process is ended only when asked, and only the
  one Windows names as holding the file.
- Anything hard to undo is the user's decision: deleting a module, overwriting
  cells that hold data, writing to a project that is signed or password-protected.
  The last two are refused until the call carries the flag that allows them.
- A cell value is what Excel last calculated. A formula written here has no result
  in the file until Excel next opens it. `calculate=true` works results out with
  pyOfficeEditor's formula engine, which names any cell it could not.

## Built on

| | |
|---|---|
| [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA) | Reads and writes VBA, UserForms and Power Query inside Office files, in pure Python. |
| [pyOfficeEditor](https://github.com/WilliamSmithEdward/pyOfficeEditor) | The document surface: cells, formulas, formatting, tables, validation, rows and columns. |
| [pyVBAanalysis](https://github.com/WilliamSmithEdward/pyVBAanalysis) | The static analyzer: 131 diagnostics, measured against each host's object model. |
| [pyVBAharness](https://github.com/WilliamSmithEdward/pyVBAharness) | Runs VBA in desktop Office under a supervisor that enforces a deadline. |
| [XLIDE for VS Code](https://github.com/WilliamSmithEdward/xlide_vscode) | Where the tool surface, the content-token guard and the agent instructions come from. |
| [xlide for the VBE](https://github.com/WilliamSmithEdward/xlide_vbide) | The live editor session the `xlide_live_*` tools talk to. |

## Contributing

```bash
cd python
pip install -e ".[dev,live]"
python -m pytest              # the file layer, no Office needed
python -m pytest -m live      # the rest, real Office, Windows only
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
