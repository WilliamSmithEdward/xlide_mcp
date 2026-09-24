# Working on xlide_mcp

An MCP server for the inside of an Office file: the VBA in Excel, Word,
PowerPoint and Access, and the document around it. Visual Basic 6 projects open
the same way.
[README.md](README.md) says what it does. This says how to change it.

Adopt these first:

- `F:\GitHub\RIDM_Recursive_Invariant_Discovery_Model\RIDM.MD`
- `F:\GitHub\AI_Best_Practices\docs\agentic_ai_programming_best_practices.md`
- `F:\GitHub\AI_Best_Practices\docs\ai_smells_for_agents_to_avoid.md`
- `F:\GitHub\AI_Best_Practices\docs\ui_ux_guidelines_for_agents.md`

## The one rule that shapes everything else

**Python is normative. Every other implementation ports from it.**

`python/` tracks the upstream libraries. A sibling directory per language tracks
`python/`. Changes flow one way: implement in Python, regenerate the contract, and
the diff under `contract/` is the work list for every port. A port that fixes
something Python has wrong fixes it in Python too, or the fix is lost at the next
sync.

This is not a preference about languages. The server is a thin layer over four
libraries that hold the measured knowledge of the Office file formats:

| Library | What it knows |
|---|---|
| [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA) | How to read and write VBA, UserForm designs and Power Query inside the containers. |
| [pyOfficeEditor](https://github.com/WilliamSmithEdward/pyOfficeEditor) | The document surface: cells, formulas, formatting, tables, validation, rows and columns. |
| [pyVBAanalysis](https://github.com/WilliamSmithEdward/pyVBAanalysis) | 131 diagnostics, each measured against its host's object model. |
| [pyVBAharness](https://github.com/WilliamSmithEdward/pyVBAharness) | How to run VBA in desktop Office without wedging on a dialog. |

The split between the first two is the file itself: pyOpenVBA edits the VBA
project, pyOfficeEditor edits the document it lives in.

A port reimplements the *server*. It does not reimplement that knowledge, and it
is never the place a format discovery lands: that belongs upstream, and reaches
here through a version bump.

### One debt against that rule, half paid

`shapes.py` and `xlsx.py` held format knowledge this server should not own: the
worksheet grid, and the drawing layer where a button keeps the macro it runs.
Both were ported from XLIDE because no library reached either.

**The grid is done.** pyOfficeEditor covers it, so `cells.py` is a thin adapter
over it and the tools read and write through that. The swap was provable rather
than hopeful, which is the point of the corpus: every cell conformance case
passed before and after, unchanged. What the adapter owns is translation, not
format knowledge - an integer where the file stores a double, a leading `=` on a
formula, a JSON-safe date - because those are answers the corpus pins.

**Shapes are not.** `shapes.py` still reads the drawing layer itself, and still
reaches into `xlsx.py` for the package surface underneath it: `part_text`,
`set_part_text`, `part_relationships`, `sheets`. That is the only reason
`xlsx.py` still exists. When `shapes.py` becomes a pyOpenVBA adapter, both go in
one deletion rather than two risky trims.

Until then, nothing new goes into either. Adding or deleting a shape means
creating a drawing part, a content-type override and a relationship, and for a
Forms control four parts that have to agree - exactly the knowledge that belongs
upstream. `xlide_set_shape_macro` covers what an agent is usually after: pointing
a clickable thing at a Sub it just wrote.

## Layout

```
contract/          tool-surface.json, conformance.json   generated, normative
docs/porting.md    how a port is built and verified
python/            the reference implementation
  tests/           the suite; the ones marked live need Windows with Office
  tools/           the two contract exporters
<language>/        a port
```

Inside `python/src/xlide_mcp/`:

```
server.py         builds the MCPServer and registers every tool group
instructions.py   what the calling model is told at initialize
config.py         settings, and the workspace roots that bound every path
paths.py          resolving a caller's path, or refusing it with the reason
hosts.py          extension -> host, and what can be done with each
project.py        the VBA project: modules, kinds, guarded saves
tokens.py         content tokens, the guard on a stale write
textual.py        the file as text: what a diff of it reads
cells.py          worksheet cells, through pyOfficeEditor
xlsx.py           the OOXML package surface the drawing layer still needs
grid.py           the same, through Excel, for the formats that are not OOXML
shapes.py         the drawing layer: buttons, shapes, and the macros they run
vb6.py            a .vbp read as a project, through the same surface
locks.py          which process holds a locked file, and what frees it
office_apps.py    a file in the user's own Office application, from a worker process
xlide_vscode.py   telling a running XLIDE for VS Code what a tool changed
errors.py         the one error type, and the helpers that build its message
tools/
  discovery.py    list, summarize, validate, create, doctor
  modules.py      read, write, rename, delete, search, list procedures
  analysis.py     the build gate, and the rule catalogue
  catalog.py      project references, and an Access database's tables
  forms.py        UserForm and Access designs
  powerquery.py   the M code beside the VBA
  sheets.py       cells, formulas, and the shapes on a sheet
  formatting.py   how a range looks: fonts, fills, borders, merging
  structure.py    sheets, and the rows and columns in them
  features.py     tables, names, validation, rules, links, page setup
  sync.py         export and import .bas/.cls, previewed
  vcs.py          what changed inside the file since a git revision
  execution.py    macros, tests and compile checks in real Office
  office.py       the file in the user's own Office: is it open, open it, close it
  live.py         a running xlide_vbide session in the VBE
```

Tool groups are split by what they reach, because that is also how they fail: the
file layer works anywhere, execution and the office tools need Windows with the
application, and the live layer needs the VBE add-in running. XLIDE for VS Code
is not a layer: when it is running, every write tells it what changed, and when
it is not, nothing notices (docs/xlide-vscode-bridge.md).

## Before you change a tool

Read [contract/README.md](contract/README.md) for what the generated files mean.
Then, in order:

1. Change the Python implementation.
2. `cd python && python -m pytest` - no Office needed.
3. `python -m ruff check src tests tools`
4. `python tools/export_contract.py && python tools/export_conformance.py`
5. A new or renamed tool goes in both READMEs: [README.md](README.md) and
   [python/README.md](python/README.md), which is the PyPI page.
6. On Windows with Office: `python -m pytest -m live`.

Steps 4 and 5 are not optional and cannot be skipped quietly. The contract is
derived from the running server precisely so it cannot drift from it, and
`tests/test_contract.py` fails when either artifact is stale.
`tests/test_docs.py` does the same for the documents: it reads the tool names and
the conformance counts out of the contract, so a tool nobody documented, a name
left behind by a rename, and a count that no longer matches all fail the build.

## What this server owes its callers

The tool descriptions and the server instructions are the product, not
documentation of it. They are what the calling model reads to decide what to do,
and a behaviour change that leaves them stale has shipped a lie. Both are in the
contract for that reason.

Rules the surface holds to, each of which has a conformance case behind it:

- **A path argument is untrusted input.** It resolves, symlinks included, before
  the containment check, and is refused outside the workspace roots.
- **A write is guarded.** A read returns a content token; a write with a stale one
  is refused and the refusal carries the current token, so recovery is one step.
- **Nothing touches an application the user is running, unless asked to.** A run
  happens in an instance the server created, which is the only reason it can
  enforce a deadline by terminating it. The three office tools are the exception
  the user asked for, and they act on the one file named: a copy holding unsaved
  work is closed only with save_changes or discard_changes, an application is
  quit only if this server started it and nothing else is open in it, and a
  process is ended only with end_process, only the one Windows names as holding
  the file, and only if it is that file's own application.
- **Hard-to-undo things are the user's decision.** Deleting a module, overwriting
  cells that hold data, writing to a signed or password-protected project. Each
  of the last two is a refusal until the flag that allows it is passed: a
  warning after the write is not asking first.
- **No result is claimed that was not observed.** Nothing in the files layer
  calculates anything, so a cell write answers `recalculated: false`. An
  implementation that reported a computed value it did not compute would be
  lying to the user through the agent.

A refusal is as much of the contract as a success: 23 of the 84 conformance cases
assert a failure message. An error message is the whole of what the calling agent
has to work with, so each one names what was refused and what to do instead.

## Prose

Plain ASCII, no em dashes, no AI tells. This applies to tool descriptions and
error messages as much as to documentation, because an agent reads them and a
user reads what the agent repeats back.

## Releasing

The version lives in exactly one place, `python/src/xlide_mcp/_version.py`. The
packaging metadata, the server's `version` field and the generated contract all
read from it. Bump that line, regenerate the contract, tag `v*.*.*`.

A release title is the tag and nothing else. Commit subjects follow the
convention already in `git log`, not an older style still visible in it.

Pushing a `v*.*.*` tag builds the package, publishes it to PyPI through Trusted
Publishing and cuts the GitHub release. The workflow refuses a tag that does not
match `_version.py`, and refuses to publish at all if the generated contract is
out of date, because a release whose tool surface does not match its artifacts is
one every port would be verified against wrongly.

```bash
# from the repository root, with CHANGELOG.md written for the version
git tag v1.0.0 && git push origin v1.0.0
```
