"""Coverage-guided fuzzing of the text this server parses itself.

The Office containers are read by pyOpenVBA, pyOfficeEditor and
pyVBAanalysis, which fuzz their own readers. What remains here:

  xml   the drawing layer's tag scanner and entity decoder (xlsx.py), on
        arbitrary text: scanning ends, and a malformed part is refused with
        a ToolError.
  a1    cell and range references (xlsx.py): read, or refused with a
        ToolError.
  vbp   a VB6 project manifest (vb6.py), written to a scratch folder with
        one module beside it: read, or refused with a ToolError.

A ToolError is the server's answer to bad input; anything else is a finding.

    python fuzz/fuzz_server.py <target> [libFuzzer options] [corpus dirs]
    python fuzz/fuzz_server.py xml -max_total_time=60 tests/fuzz_corpus/xml

Run from python/. The .github/workflows/fuzz.yml workflow runs each target
from tests/fuzz_corpus/<target>; a finding becomes a seed there, which
tests/test_fuzz_corpus.py replays on every CI run.
"""

import contextlib
import sys
import tempfile
from pathlib import Path

import atheris

with atheris.instrument_imports():
    from mcp.server.mcpserver.exceptions import ToolError

    from xlide_mcp import Settings, vb6, xlsx

SCRATCH = Path(tempfile.mkdtemp(prefix="xlide-fuzz-"))
(SCRATCH / "Module1.bas").write_bytes(
    b'Attribute VB_Name = "Module1"\r\nOption Explicit\r\nSub Main()\r\nEnd Sub\r\n'
)


def fuzz_xml(data):
    text = data.decode("utf-8", errors="replace")
    position = 0
    try:
        while True:
            tag = xlsx.next_tag(text, position)
            if tag is None:
                break
            if tag.end <= position:
                raise AssertionError("the tag scanner did not advance")
            position = tag.end
        xlsx.decode_xml(text)
    except ToolError:
        pass


def fuzz_a1(data):
    text = data.decode("utf-8", errors="replace")
    for parse in (xlsx.parse_cell_ref, xlsx.parse_range, xlsx.formula_for_file):
        with contextlib.suppress(ToolError):
            parse(text)


def fuzz_vbp(data):
    manifest = SCRATCH / "Project1.vbp"
    manifest.write_bytes(data)
    try:
        # Rooted at the scratch folder, so a module line that escapes it is
        # refused with ToolError, which this target accepts.
        project = vb6.Vb6Project(manifest, settings=Settings(roots=(SCRATCH.resolve(),)))
        project.module_names()
        _ = project.references
        _ = project.name
        project.validate()
    except ToolError:
        pass


TARGETS = {"xml": fuzz_xml, "a1": fuzz_a1, "vbp": fuzz_vbp}


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in TARGETS:
        sys.exit(f"usage: fuzz_server.py <{'|'.join(TARGETS)}> [libFuzzer options]")
    atheris.Setup([sys.argv[0], *sys.argv[2:]], TARGETS[sys.argv[1]])
    atheris.Fuzz()


if __name__ == "__main__":
    main()
