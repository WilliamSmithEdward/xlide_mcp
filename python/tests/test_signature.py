"""A digitally signed project: the write is the user's decision, and it is asked for.

Excel, Word and PowerPoint keep the signature of a zip-based file in parts beside
vbaProject.bin, not in its streams. Until pyOpenVBA 6.1 nothing looked there, so
every such file read as unsigned and a write kept a signature that no longer
covered the code. 6.1 sees the parts and drops them on a save that changes the
code, as Office does, with a warning. A warning after the fact is not asking
first, so this server refuses the write until allow_invalidate_signature says the
user agreed. The `signed_workbook` fixture is in conftest.py, because the
conformance corpus names it too.
"""

from __future__ import annotations

import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import SAMPLE_MODULE, SIGNATURE_PART, ToolFailure


def test_a_signature_beside_the_project_is_seen(
    call: Callable[..., Any], signed_workbook: Path
) -> None:
    info = call("xlide_project_info", file_path=str(signed_workbook))
    assert info["digitally_signed"] is True
    assert any("allow_invalidate_signature" in caution for caution in info["cautions"])


def test_a_write_to_a_signed_project_is_refused_and_writes_nothing(
    call: Callable[..., Any], signed_workbook: Path
) -> None:
    before = signed_workbook.read_bytes()
    with pytest.raises(ToolFailure) as refusal:
        call(
            "xlide_write_module",
            file_path=str(signed_workbook),
            module_name="Helpers",
            source=SAMPLE_MODULE + "\n' changed\n",
        )
    assert "digitally signed" in refusal.value.message
    assert "allow_invalidate_signature=true" in refusal.value.message
    assert signed_workbook.read_bytes() == before


def test_with_the_users_agreement_the_signature_goes_with_the_write(
    call: Callable[..., Any], signed_workbook: Path
) -> None:
    written = call(
        "xlide_write_module",
        file_path=str(signed_workbook),
        module_name="Helpers",
        source=SAMPLE_MODULE + "\n' changed\n",
        allow_invalidate_signature=True,
    )
    assert written["saved"] is True
    assert call("xlide_project_info", file_path=str(signed_workbook))["digitally_signed"] is False
    with zipfile.ZipFile(signed_workbook) as package:
        assert SIGNATURE_PART not in package.namelist()
