"""The Excel-backed sheet survey must not turn malformed output into a partial list."""

from __future__ import annotations

import pytest

from xlide_mcp.errors import ToolError
from xlide_mcp.grid import _parse_sheet_survey


def test_sheet_survey_reads_visible_and_hidden_sheets() -> None:
    sheets = _parse_sheet_survey("Data\tA1:C4\tTrue\nArchive\tA1\tFalse\n")
    assert [(s.name, s.used_range, s.hidden) for s in sheets] == [
        ("Data", "A1:C4", False),
        ("Archive", "A1", True),
    ]


@pytest.mark.parametrize(
    "raw",
    [
        "Data\tA1:C4\tTrue\nBroken\tA1\n",
        "Data\tA1:C4\tTrue\nBroken\tA1\tFalse\textra\n",
    ],
)
def test_sheet_survey_refuses_a_partial_listing(raw: str) -> None:
    with pytest.raises(ToolError) as refusal:
        _parse_sheet_survey(raw)
    assert "malformed worksheet details on line 2" in str(refusal.value)


def test_sheet_survey_refuses_an_unknown_visibility_state() -> None:
    with pytest.raises(ToolError) as refusal:
        _parse_sheet_survey("Data\tA1:C4\tmaybe\n")
    assert "unknown visibility state" in str(refusal.value)
    assert "Data" in str(refusal.value)
