# xlide-mcp (Python)

The reference implementation. See the [repository README](../README.md) for what
the server is and [docs/porting.md](../docs/porting.md) for how ports relate to
this one.

## Install

```bash
pip install -e ".[dev,live]"     # omit live off Windows
```

`live` brings pyvbaharness and pywin32, which are what run VBA in a desktop
application. Everything else works without them.

## Run

```bash
xlide-mcp --root /path/to/files            # stdio, the usual case
xlide-mcp --root /work --read-only         # reads and analysis only
xlide-mcp --transport streamable-http --port 8765
python -m xlide_mcp --help
```

| Setting | Flag | Environment |
|---|---|---|
| Workspace roots | `--root` (repeatable) | `XLIDE_MCP_ROOTS` |
| Refuse writes | `--read-only` | `XLIDE_MCP_READ_ONLY` |
| Allow any absolute path | `--allow-outside-roots` | `XLIDE_MCP_ALLOW_OUTSIDE_ROOTS` |
| Default run deadline | `--timeout` | `XLIDE_MCP_TIMEOUT` |

An HTTP transport binds to loopback unless told otherwise. This server reads and
writes files; a default that listened on every interface would hand that reach to
the network.

## Layout

```
src/xlide_mcp/
  server.py         builds the MCPServer and registers every tool group
  instructions.py   what the calling model is told at initialize
  config.py         settings, and the workspace roots that bound every path
  paths.py          resolving a caller's path, or refusing it with the reason
  hosts.py          extension -> host, and what can be done with each
  project.py        the VBA project: modules, kinds, guarded saves
  tokens.py         content tokens, the guard on a stale write
  xlsx.py           worksheet cells, read and written in the OOXML package
  shapes.py         the drawing layer: buttons, shapes, and the macros they run
  errors.py         the one error type, and the helpers that build its message
  tools/
    discovery.py    list, summarize, validate, create, doctor
    modules.py      read, write, rename, delete, search, list procedures
    analysis.py     the build gate, and the rule catalogue
    catalog.py      project references, and an Access database's tables
    forms.py        UserForm and Access designs
    powerquery.py   the M code beside the VBA
    sheets.py       cells, formulas, and the shapes on a sheet
    sync.py         export and import .bas/.cls, previewed
    vcs.py          what changed inside the file since a git revision
    execution.py    macros, tests and compile checks in real Office
    live.py         a running xlide_vbide session in the VBE
tools/
  export_contract.py     writes ../contract/tool-surface.json
  export_conformance.py  writes ../contract/conformance.json
```

Tool groups are split by what they reach, because that is also how they fail: the
file layer works anywhere, execution needs Windows with the application, and the
live layer needs the VBE add-in running.

## Test

```bash
python -m pytest                 # 171 tests, no Office needed
python -m pytest -m live         # 11 more, real Office, Windows only
python -m ruff check src tests tools
```

The live tests are opt-in because they start desktop applications, and because
only one Office session runs at a time per machine: interleaving them with the
rest produces failures that are contention, not defects.

Three things in the suite are worth knowing about:

- `test_conformance.py` runs `../contract/conformance.json`, the corpus every
  implementation in the repository owes. A case that is wrong fails here before
  any port is built against it.
- `test_contract.py` fails if either generated contract file is out of date, so a
  tool cannot change without the artifacts changing with it.
- `test_execution_live.py` ends with the gate that matters for the cell-write
  path: Excel opens a workbook this server spliced, recalculates it, and reads the
  formula back as it was typed.

## After changing a tool

```bash
python tools/export_contract.py
python tools/export_conformance.py
```

The diff under `contract/` is the work list for every other implementation.
