# The contract

Two generated files. Between them they say what every implementation in this
repository must expose and what its answers must mean.

| File | What it pins | Regenerate with |
|---|---|---|
| `tool-surface.json` | Every tool: name, description, argument schema, and whether it writes. | `python tools/export_contract.py` |
| `conformance.json` | What the answers mean: 42 cases, each a short script of calls and assertions. | `python tools/export_conformance.py` |

Both are **derived from the Python implementation**, never hand-edited. A
hand-maintained contract drifts the first time someone adds an argument and
forgets the document, and a port verified against a drifted contract is verified
against nothing. `python -m pytest tests/test_contract.py` fails if either file is
out of date, so the drift cannot reach a commit.

## Why Python is the reference

The Python server is a thin layer over three libraries that do the hard part:

- [pyOpenVBA](https://pypi.org/project/pyOpenVBA/) reads and writes the VBA, the
  UserForms and the Power Query inside Office files.
- [pyvbaanalysis](https://pypi.org/project/pyvbaanalysis/) is the static analyzer.
- [pyvbaharness](https://pypi.org/project/pyvbaharness/) runs VBA in a desktop
  application under a deadline.

Those libraries are where the measured knowledge of the file formats lives, and
they move. The Python implementation tracks them; every other implementation
tracks the Python implementation. `tool-surface.json` records the upstream
versions it was generated against under `generated_with`, so "we are behind" is a
diff rather than a feeling.

## `tool-surface.json`

```json
{
  "contract_version": "0.1.0",
  "reference_implementation": "python",
  "digest": "sha256:...",
  "tool_count": 36,
  "tools": [ { "name": "...", "description": "...", "input_schema": {...}, "annotations": {...} } ],
  "generated_with": {
    "upstream": { "pyOpenVBA": "5.2.4", "pyvbaanalysis": "2.1.1", "pyvbaharness": "1.1.2" }
  }
}
```

`digest` covers the server and its tools and nothing else, so it means "this tool
surface". A port pins it and detects any change to the surface in one comparison.

`generated_with` records which upstream libraries the export ran against. It sits
outside the digest and outside the drift check on purpose: it describes the
machine that ran the export, not the surface. A developer's checkout is routinely
ahead of the released versions CI installs, and a check that failed on that would
be a check nobody could keep green. The drift check still reports the difference,
it just does not fail on it.

The tool **descriptions are part of the contract**, not documentation of it. They
are what the calling model reads to decide which tool to use and how; a port that
reimplements the behaviour and rewrites the prose has changed the product.

## `conformance.json`

The tool surface says what the arguments are. The corpus says what the answers
mean, which is the part a port gets wrong.

```json
{
  "id": "analyze.positions-match-what-a-read-returns",
  "why": "Analysis measures a source that carries the attribute header; a read strips it...",
  "requires": "files",
  "fixture": "workbook",
  "steps": [ { "tool": "xlide_write_module", "arguments": {...} },
             { "tool": "xlide_analyze", "arguments": {...} } ],
  "expect": [ { "path": "problems[0].line", "equals": 5 } ]
}
```

Each case carries a `why`. A case without a stated reason to exist is a case
nobody can decide whether to change, and the Python suite refuses one.

**`requires`** says what a case needs, so a runner can select what it can run:

| Value | Needs |
|---|---|
| `files` | Nothing. Reads and writes the Office file on any platform. |
| `git` | git on the PATH, for the cases that diff against a revision. |
| `office` | Windows with the desktop application installed. |
| `live` | A running `xlide_vbide` session inside the Visual Basic Editor. |

**Fixtures** are described, not shipped. `workbook` is a macro-enabled workbook
built from the application's own template with one named module in it; a runner
builds it with whatever that language binds to. Shipping the binaries instead
would pin one version of one template forever.

One fixture is the exception and says so: `shapes_workbook` carries a
`repository_path`, because a Forms-toolbar button lives in four parts that only
Excel writes in agreement, and no library can build one. Copy that file.

**Placeholders** inside arguments:

| Placeholder | Means |
|---|---|
| `${fixture}` | The path of the case's fixture file. |
| `${folder:name}` | A scratch folder under the workspace. |
| `${step[N].path}` | A value from an earlier step's result, 0-based. |
| `${outside}` | A real path outside the workspace roots. |

**Assertions** over the last step's result, all optional:
`equals`, `not_equals`, `contains`, `not_contains`, `starts_with`,
`at_least` (a count, or a collection's length), and
`type` (`string`, `number`, `boolean`, `array`, `object`, `null`).
`path` walks `a.b[0].c`; a missing step reads as null, so absence is expressible.

A step may instead declare `error_contains`, which requires that step to fail
with that text in the message. Fourteen of the cases are refusals: what a server
does when asked for something it should not do is as much of the contract as what
it does when asked for something it should.

See [docs/porting.md](../docs/porting.md) for how to build a port against these.
