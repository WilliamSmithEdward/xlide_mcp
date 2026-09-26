"""Results stay a size an agent can think with, and say when they stopped short.

A tool result is model context. A legacy project with 400 modules is ordinary and
a data-entry form with 400 controls is not absurd; measured before this, that form
came back as 112 KB in one call, which is most of what an agent has to work with.

The bound matters less than the note. A listing that silently stops short reads as
a complete answer, and an agent acts on it: it reports the workbook has 300
modules, or that the control it was asked about does not exist.
"""

from __future__ import annotations

import json
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

from xlide_mcp.tools._common import MAX_ITEMS, bound

MODULE_BODY = "Option Explicit\r\n\r\n" + "".join(
    f"Public Sub Proc{i}()\r\nEnd Sub\r\n" for i in range(10)
)


@pytest.fixture
def crowded(workspace: Path) -> Path:
    """More modules and controls than any bound allows."""
    import pyopenvba

    over = MAX_ITEMS["modules"] + 20
    path = workspace / "Crowded.xlsm"
    with pyopenvba.ExcelFile.create_new(path) as book:
        for i in range(over):
            book.vba_project().add_module(f"Mod{i:03d}", MODULE_BODY)
        form = book.add_form("Monster", width=600, height=400)
        for c in range(MAX_ITEMS["controls"] + 20):
            form.add_control("CommandButton", f"Btn{c:03d}", left=6, top=6)
        book.save()
    return path


def test_bound_passes_a_short_list_through_untouched() -> None:
    items = [1, 2, 3]
    shown, note = bound(items, "modules")
    assert shown == items
    assert note == ""


def test_bound_reports_what_it_withheld() -> None:
    items = list(range(MAX_ITEMS["modules"] + 5))
    shown, note = bound(items, "modules", "Narrow it.")

    assert len(shown) == MAX_ITEMS["modules"]
    assert str(len(items)) in note
    assert "5 are not" in note
    assert "Narrow it." in note


def test_a_module_listing_is_bounded_and_says_so(
    call: Callable[..., Any], crowded: Path
) -> None:
    result = call("xlide_list_modules", file_path=str(crowded))

    assert len(result["modules"]) == MAX_ITEMS["modules"]
    # The count is of the project, not of what was returned: an agent that
    # reported "this workbook has 300 modules" would be wrong.
    assert result["count"] > MAX_ITEMS["modules"]
    assert "note" in result


def test_module_pages_reach_modules_after_the_bound(
    call: Callable[..., Any], crowded: Path
) -> None:
    first = call("xlide_list_modules", file_path=str(crowded), max_results=1)
    assert first["next_offset"] == 1
    later = call(
        "xlide_list_modules", file_path=str(crowded),
        offset=first["count"] - 1, max_results=1,
    )
    assert later["modules"][0]["name"] == "Monster"
    assert later["next_offset"] is None


def test_the_one_shot_summary_is_bounded_too(
    call: Callable[..., Any], crowded: Path
) -> None:
    result = call("xlide_project_info", file_path=str(crowded))

    assert len(result["modules"]) == MAX_ITEMS["modules"]
    assert result["module_count"] > MAX_ITEMS["modules"]
    assert "modules_note" in result


def test_a_form_with_too_many_controls_is_bounded(
    call: Callable[..., Any], crowded: Path
) -> None:
    """A control carrying its properties costs far more than a module summary, so
    this one stops on size well before it reaches the count. That is the point of
    measuring size rather than counting: the ceiling is what a result costs."""
    result = call("xlide_read_form", file_path=str(crowded), form_name="Monster")

    shown = len(result["controls"])
    assert 0 < shown < MAX_ITEMS["controls"]
    assert result["control_count"] > shown
    assert str(result["control_count"]) in result["note"]


def test_form_control_pages_reach_the_end(
    call: Callable[..., Any], crowded: Path
) -> None:
    first = call(
        "xlide_read_form", file_path=str(crowded), form_name="Monster",
        include_properties=False, max_controls=20,
    )
    assert len(first["controls"]) == 20
    assert first["next_offset"] == 20
    second = call(
        "xlide_read_form", file_path=str(crowded), form_name="Monster",
        include_properties=False, max_controls=20, offset=first["next_offset"],
    )
    assert second["controls"][0]["name"] != first["controls"][0]["name"]
    last = call(
        "xlide_read_form", file_path=str(crowded), form_name="Monster",
        include_properties=False, max_controls=1, offset=first["control_count"] - 1,
    )
    assert last["controls"][0]["name"] == "Btn319"
    assert last["next_offset"] is None


def test_the_size_ceiling_binds_before_the_count_for_costly_items() -> None:
    """Two listings of the same length, one of which costs ten times the other."""
    cheap = [{"name": f"Mod{i}", "kind": "standard"} for i in range(200)]
    costly = [{"name": f"Ctl{i}", "properties": {f"p{k}": "x" * 20 for k in range(12)}}
              for i in range(200)]

    kept_cheap, note_cheap = bound(cheap, "modules")
    kept_costly, note_costly = bound(costly, "controls")

    assert kept_cheap == cheap and note_cheap == ""
    assert len(kept_costly) < len(costly)
    assert "are not" in note_costly


def test_the_crowded_results_fit_in_a_usable_answer(
    call: Callable[..., Any], crowded: Path
) -> None:
    """The measurement that motivated the bound. 112 KB for one form's controls
    was most of an agent's context spent on one call."""
    for name, arguments in (
        ("xlide_project_info", {}),
        ("xlide_list_modules", {}),
        ("xlide_read_form", {"form_name": "Monster"}),
    ):
        result = call(name, file_path=str(crowded), **arguments)
        size = len(json.dumps(result))
        assert size < 100_000, f"{name} answered with {size:,} characters"


def test_a_bounded_read_still_finds_what_was_asked_for(
    call: Callable[..., Any], crowded: Path
) -> None:
    """Bounding a listing must not bound the tools that answer a direct question,
    or a module past the limit becomes unreachable."""
    last = f"Mod{MAX_ITEMS['modules'] + 19:03d}"
    read = call("xlide_read_module", file_path=str(crowded), module_name=last)
    assert read["module"] == last

    hits = call("xlide_search_modules", file_path=str(crowded), query=last)
    assert hits["match_count"] == 0 or hits["matches"][0]["module"] == last
