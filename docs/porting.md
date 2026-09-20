# Porting to another language

The repository is laid out one directory per implementation. `python/` is the
reference; a new language gets a sibling directory and nothing else moves.

```
xlide_mcp/
  contract/          tool-surface.json, conformance.json   <- generated, normative
  docs/              this file
  python/            the reference implementation
  <language>/        a port
```

## The rule

**Python tracks the upstream libraries. Every other implementation tracks
Python.**

Changes flow one way. A behaviour is implemented in Python first, the contract is
regenerated, and ports follow. A port that fixes something Python has wrong fixes
it in Python too, or the fix is lost at the next sync.

This is not a preference about languages. The Python server is a thin layer over
pyOpenVBA, pyOfficeEditor, pyvbaanalysis and pyvbaharness, which is where the
measured knowledge of the Office file formats actually lives. A port reimplements
the *server*; it does not reimplement that knowledge, and it cannot be the place
a format discovery lands.

## What a port owes

1. **Every tool in `tool-surface.json`**, with the same names, the same argument
   names and types, and the same descriptions. The descriptions are what the
   calling model reads; rewriting them changes the product.
2. **Every case in `conformance.json`** that its layer supports, passing.
3. **The workspace boundary.** Path arguments resolve, symlinks and all, before
   the containment check. Two conformance cases cover it and they are the two
   worth failing the build over.
4. **The same refusals.** 20 of the 77 conformance cases assert a failure
   message. A port that succeeds where the contract refuses has a different
   product, not a lenient one.

A port does not owe the internal structure. How it opens a container, whether it
has an equivalent of `project.py`, what its modules are called: all of that is
free. The contract is the surface and the meaning of the answers.

## Layers, and what a port may leave out

A port may implement one layer and stop. A tool it does not implement should
still be registered and answer with what is missing: a tool that is absent reads
to an agent as "this server cannot do that", and sends it looking for another way
in, which is usually worse than the thing being unavailable.

| Layer | Tools | Needs |
|---|---|---|
| Files | discovery, modules, analysis, catalog, forms, Power Query, sheets, shapes, sync, git changes | Nothing. Any platform, git on the PATH for `xlide_git_changes`. |
| Execution | `xlide_run_macro`, `xlide_run_vba`, `xlide_run_tests`, `xlide_compile_check` | Windows, desktop Office. |
| Live | `xlide_live_*` | A running `xlide_vbide` session. |

The files layer is the one that matters. A port with only that layer is useful;
a port with only the others is not.

## Building a port

1. Read `contract/tool-surface.json` and generate or hand-write the tool
   registrations. The descriptions are copied verbatim.
2. Write a conformance runner. It is small on purpose: build the fixtures,
   substitute the placeholders, call the tools, walk the assertion paths. The
   Python one is `python/tests/test_conformance.py`, about 200 lines including
   its comments, and it is the specification of the runner as much as
   `contract/README.md` is.
3. Run the corpus. Everything at `requires: files` should pass before anything
   else is attempted.
4. Add the port's own tests for what the corpus does not reach. The corpus pins
   the contract; it is not a substitute for testing the code.

## Keeping in sync

When upstream moves:

1. Update the pinned versions in `python/pyproject.toml` and run the Python suite.
2. Fix whatever broke, in Python.
3. Regenerate: `python tools/export_contract.py` and
   `python tools/export_conformance.py`.
4. The diff in `contract/` is the work list for every port.
   `generated_with.upstream` in `tool-surface.json` says which library versions
   the surface now reflects.

`python -m pytest tests/test_contract.py` fails when the contract is out of date,
so step 3 cannot be forgotten quietly.

## What a port must not do

- Silently widen a refusal. If a case says a path outside the roots is refused,
  a port that allows it with a warning has broken the boundary.
- Invent a tool. A tool that exists in one implementation and not the others is
  how the surface stops being one surface. Add it to Python first.
- Claim a result it has not observed. `recalculated: false` on a cell write is
  not a formality: nothing in the files layer calculates anything, and an
  implementation that reports a computed value it did not compute is lying to
  the user through the agent.
