"""The fuzz corpus, replayed through the fuzz targets on every run.

tests/fuzz_corpus/<target> seeds fuzz/fuzz_server.py; a fuzz finding joins it
as a regression seed. A ToolError is the server's answer to bad input;
anything else escaping is a finding.
"""

from __future__ import annotations

import contextlib
from pathlib import Path

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from xlide_mcp import vb6, xlsx

CORPUS = Path(__file__).parent / "fuzz_corpus"
SEEDS = [
    (target, path)
    for target in ("xml", "a1", "vbp")
    for path in sorted((CORPUS / target).iterdir())
]


def _text(path: Path) -> str:
    return path.read_bytes().decode("utf-8", errors="replace")


@pytest.mark.parametrize(("target", "path"), SEEDS, ids=[f"{t}/{p.name}" for t, p in SEEDS])
def test_a_seed_is_read_or_refused_with_a_tool_error(
    target: str, path: Path, tmp_path: Path
) -> None:
    if target == "xml":
        text = _text(path)
        position = 0
        try:
            while (tag := xlsx.next_tag(text, position)) is not None:
                assert tag.end > position
                position = tag.end
            xlsx.decode_xml(text)
        except ToolError:
            pass
    elif target == "a1":
        for parse in (xlsx.parse_cell_ref, xlsx.parse_range, xlsx.formula_for_file):
            with contextlib.suppress(ToolError):
                parse(_text(path))
    else:
        (tmp_path / "Module1.bas").write_bytes(
            b'Attribute VB_Name = "Module1"\r\nOption Explicit\r\nSub Main()\r\nEnd Sub\r\n'
        )
        manifest = tmp_path / "Project1.vbp"
        manifest.write_bytes(path.read_bytes())
        try:
            project = vb6.Vb6Project(manifest)
            project.module_names()
            project.validate()
        except ToolError:
            pass
