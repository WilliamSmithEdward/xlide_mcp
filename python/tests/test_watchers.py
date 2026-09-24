"""What a write looks like to a file watcher.

xlide_vscode follows a workbook's project tree through a file watcher, which
reports a changed file as `change` when the same path's size and modification
time move. A write that deleted the target and made a new one, or left a temp
file beside it, would show up as something else (xlide_mcp#1). Every save here
serializes in memory and writes over the same file, so the file is the same
one afterwards.
"""

from __future__ import annotations

import os
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any

from conftest import SAMPLE_MODULE


def _changed_the_same_file(path: Path, write: Callable[[], Any]) -> None:
    before = path.stat()
    siblings = sorted(p.name for p in path.parent.iterdir())
    # Modification times are coarse on some file systems; step past a tick.
    time.sleep(0.05)
    write()
    after = path.stat()
    assert after.st_mtime_ns > before.st_mtime_ns, "the modification time moved"
    if os.name == "nt" or before.st_ino:
        assert after.st_ino == before.st_ino, "the same file, not a new one in its place"
    assert sorted(p.name for p in path.parent.iterdir()) == siblings, "nothing left beside it"


def test_a_module_write_changes_the_file_in_place(
    call: Callable[..., Any], workbook: Path
) -> None:
    _changed_the_same_file(
        workbook,
        lambda: call(
            "xlide_write_module",
            file_path=str(workbook),
            module_name="Helpers",
            source=SAMPLE_MODULE + "\n' changed\n",
        ),
    )


def test_a_cell_write_changes_the_file_in_place(
    call: Callable[..., Any], plain_workbook: Path
) -> None:
    _changed_the_same_file(
        plain_workbook,
        lambda: call(
            "xlide_write_cells",
            file_path=str(plain_workbook),
            sheet="Sheet1",
            start_cell="A1",
            data=[["changed"]],
        ),
    )
