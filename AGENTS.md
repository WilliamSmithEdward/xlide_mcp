# Working on xlide_mcp

An MCP server for the code inside Office files: the VBA, the UserForms, the Power
Query and the worksheet cells in Excel, Word, PowerPoint and Access documents,
plus Visual Basic 6 projects.
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

This is not a preference about languages. The server is a thin layer over three
libraries that hold the measured knowledge of the Office file formats:

| Library | What it knows |
|---|---|
| [pyOpenVBA](https://github.com/WilliamSmithEdward/pyOpenVBA) | How to read and write VBA, UserForm designs and Power Query inside the containers. |
| [pyVBAanalysis](https://github.com/WilliamSmithEdward/pyVBAanalysis) | 119 diagnostics, each measured against its host's object model. |
| [pyVBAharness](https://github.com/WilliamSmithEdward/pyVBAharness) | How to run VBA in desktop Office without wedging on a dialog. |

A port reimplements the *server*. It does not reimplement that knowledge, and it
is never the place a format discovery lands: that belongs upstream, and reaches
here through a version bump.

## Layout

```
contract/          tool-surface.json, conformance.json   generated, normative
docs/porting.md    how a port is built and verified
python/            the reference implementation
  src/xlide_mcp/   the server
  tests/           201 without Office, 11 more with
  tools/           the two contract exporters
<language>/        a port
```

## Before you change a tool

Read [python/README.md](python/README.md) for the module layout and
[contract/README.md](contract/README.md) for what the generated files mean.

Then, in order:

1. Change the Python implementation.
2. `cd python && python -m pytest` - 201 tests, no Office needed.
3. `python -m ruff check src tests tools`
4. `python tools/export_contract.py && python tools/export_conformance.py`
5. On Windows with Office: `python -m pytest -m live` - 11 more.

Step 4 is not optional and cannot be skipped quietly: `tests/test_contract.py`
fails when either artifact is stale, and so does CI. The contract is derived from
the running server precisely so it cannot drift from it.

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
- **Nothing touches an application the user is running.** A run happens in an
  instance the server created, which is the only reason it can enforce a deadline
  by terminating it.
- **Hard-to-undo things are the user's decision.** Deleting a module, overwriting
  cells that hold data, writing to a signed or password-protected project.
- **No result is claimed that was not observed.** Nothing in the files layer
  calculates anything, so a cell write answers `recalculated: false`. An
  implementation that reported a computed value it did not compute would be
  lying to the user through the agent.

A refusal is as much of the contract as a success: 14 of the 52 conformance cases
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
