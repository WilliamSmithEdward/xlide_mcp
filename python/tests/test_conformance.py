"""Run the repo's conformance corpus against this implementation.

The corpus at contract/conformance.json is what every implementation in this
repository owes. Python is the reference, so it runs the corpus too: a case that
is wrong fails here, before any port is built against it.

The runner is deliberately small. Another language's runner has to reproduce it,
and a vocabulary that needs a parser is a vocabulary ports will implement three
different ways.
"""

from __future__ import annotations

import json
import re
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest
from conftest import ToolFailure

CORPUS_PATH = Path(__file__).resolve().parents[2] / "contract" / "conformance.json"


def load_corpus() -> dict[str, Any]:
    if not CORPUS_PATH.is_file():
        pytest.skip(f"{CORPUS_PATH} is missing; run python tools/export_conformance.py")
    return json.loads(CORPUS_PATH.read_text(encoding="utf-8"))


CORPUS = load_corpus() if CORPUS_PATH.is_file() else {"cases": []}
CASES = [c for c in CORPUS.get("cases", []) if c.get("requires", "files") == "files"]


def test_the_corpus_is_present_and_whole() -> None:
    assert CORPUS["cases"], "the corpus is empty"
    assert CORPUS["case_count"] == len(CORPUS["cases"])
    identifiers = [case["id"] for case in CORPUS["cases"]]
    assert len(identifiers) == len(set(identifiers)), "case ids must be unique"
    for case in CORPUS["cases"]:
        assert case["why"], f"{case['id']} has no stated reason to exist"
        assert case["steps"], f"{case['id']} has no steps"


@pytest.mark.parametrize("case", CASES, ids=lambda c: c["id"])
def test_conformance_case(
    case: dict[str, Any],
    call: Callable[..., Any],
    workspace: Path,
    request: pytest.FixtureRequest,
) -> None:
    fixture_path = ""
    if case.get("fixture"):
        fixture_path = str(request.getfixturevalue(case["fixture"]))

    results: list[Any] = []
    last: Any = None
    for index, step in enumerate(case["steps"]):
        arguments = _substitute(step["arguments"], fixture_path, workspace, results)
        expected_error = step.get("error_contains")
        try:
            last = call(step["tool"], **arguments)
        except ToolFailure as failure:
            if not expected_error:
                raise AssertionError(
                    f"{case['id']} step {index} ({step['tool']}) failed: {failure.message}"
                ) from failure
            assert expected_error in failure.message, (
                f"{case['id']} step {index} refused with the wrong message.\n"
                f"  wanted: {expected_error}\n  got:    {failure.message}"
            )
            return
        if expected_error:
            raise AssertionError(
                f"{case['id']} step {index} ({step['tool']}) was expected to fail with "
                f"{expected_error!r}, and succeeded."
            )
        results.append(last)

    for assertion in case.get("expect", []):
        _check(case["id"], assertion, last, fixture_path, workspace, results)


# ------------------------------------------------------------- substitution

_STEP_REF = re.compile(r"^\$\{step\[(\d+)\]\.(.+)\}$")
_FOLDER_REF = re.compile(r"^\$\{folder:([A-Za-z0-9_-]+)\}$")


def _substitute(value: Any, fixture: str, workspace: Path, results: list[Any]) -> Any:
    if isinstance(value, dict):
        return {k: _substitute(v, fixture, workspace, results) for k, v in value.items()}
    if isinstance(value, list):
        return [_substitute(v, fixture, workspace, results) for v in value]
    if not isinstance(value, str):
        return value
    return _substitute_text(value, fixture, workspace, results)


def _substitute_text(text: str, fixture: str, workspace: Path, results: list[Any]) -> Any:
    match = _STEP_REF.match(text)
    if match:
        return _dig(results[int(match.group(1))], match.group(2))
    match = _FOLDER_REF.match(text)
    if match:
        return str(workspace / match.group(1))
    if "${fixture}" in text:
        text = text.replace("${fixture}", fixture)
    if "${outside}" in text:
        # A path that is real and definitely not under the workspace root.
        text = text.replace("${outside}", str(workspace.parent.parent))
    return text


# ---------------------------------------------------------------- assertions

_INDEX = re.compile(r"\[(\d+)\]")


def _dig(value: Any, path: str) -> Any:
    """Walk `a.b[0].c`. A missing step is None, not an exception: an assertion
    that something is absent has to be expressible."""
    current = value
    for raw in path.split("."):
        name, *indexes = _INDEX.split(raw)
        indexes = [i for i in indexes if i != ""]
        if name:
            if not isinstance(current, dict) or name not in current:
                return None
            current = current[name]
        for index in indexes:
            if not isinstance(current, (list, tuple)) or int(index) >= len(current):
                return None
            current = current[int(index)]
    return current


_TYPES: dict[str, Any] = {
    "string": str,
    "number": (int, float),
    "boolean": bool,
    "array": list,
    "object": dict,
}


def _check(
    case_id: str,
    assertion: dict[str, Any],
    result: Any,
    fixture: str,
    workspace: Path,
    results: list[Any],
) -> None:
    path = assertion["path"]
    actual = _dig(result, path)
    where = f"{case_id}: {path}"

    if "equals" in assertion:
        wanted = _substitute(assertion["equals"], fixture, workspace, results)
        assert actual == wanted, f"{where} is {actual!r}, wanted {wanted!r}"
    if "not_equals" in assertion:
        assert actual != assertion["not_equals"], f"{where} is {actual!r}, wanted anything else"
    if "contains" in assertion:
        assert actual is not None, f"{where} is absent"
        assert assertion["contains"] in _searchable(actual), (
            f"{where} does not contain {assertion['contains']!r}"
        )
    if "not_contains" in assertion:
        assert assertion["not_contains"] not in _searchable(actual), (
            f"{where} contains {assertion['not_contains']!r} and should not"
        )
    if "starts_with" in assertion:
        assert isinstance(actual, str) and actual.startswith(assertion["starts_with"]), (
            f"{where} is {actual!r}, wanted a string starting {assertion['starts_with']!r}"
        )
    if "at_least" in assertion:
        # A count answers directly; a collection answers with its length.
        size = actual if isinstance(actual, int) else len(actual or [])
        assert size >= assertion["at_least"], (
            f"{where} measures {size}, wanted at least {assertion['at_least']}"
        )
    if "type" in assertion:
        wanted = assertion["type"]
        if wanted == "null":
            assert actual is None, f"{where} is {actual!r}, wanted null"
        else:
            assert isinstance(actual, _TYPES[wanted]), f"{where} is {actual!r}, wanted {wanted}"


def _searchable(value: Any) -> str:
    """Containment over a string, or over a list's stringified members."""
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        return json.dumps(value)
    return json.dumps(value, default=str)
