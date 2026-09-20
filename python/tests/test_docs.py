"""The documents are checked against the contract rather than against memory.

Two READMEs list the tool surface, one for the repository and one for the package
page on PyPI, and both are written by hand because each is read by a different
person arriving cold. A hand-written list of a generated thing drifts: when these
tests were added, four documents stated a conformance-case count and three of them
were wrong, in a repository whose whole discipline is that drift fails the build.

So the counts and the tool names are derived here too. A tool added without being
documented fails, a tool named in a document that no longer exists fails, and a
count that no longer matches the corpus fails, each with the current value in the
message.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
CONTRACT = REPO_ROOT / "contract"

# The repository, the package and the VS Code and VBE editors share the prefix
# and are not tools.
NOT_TOOLS = {"xlide_mcp", "xlide_vbide", "xlide_vscode"}

# A released changelog entry is history. It may name a tool that has since been
# renamed or removed, and rewriting it to keep a test green would be falsifying
# the record, so it is not scanned for tool names.
DOCS_WITH_LIVE_TOOL_NAMES = (
    "README.md",
    "AGENTS.md",
    "contract/README.md",
    "docs/porting.md",
    "python/README.md",
)

# Where the tool surface is listed for a reader: the repository landing page and
# the package page.
DOCS_LISTING_EVERY_TOOL = ("README.md", "python/README.md")

# A trailing `*` is prose for a family, as in `xlide_live_*`, and is checked as a
# prefix rather than as a name.
TOOL_NAME = re.compile(r"xlide_[a-z][a-z0-9_]*\*?")


def _read(relative: str) -> str:
    return (REPO_ROOT / relative).read_text(encoding="utf-8")


def _contract_tool_names() -> set[str]:
    surface = json.loads((CONTRACT / "tool-surface.json").read_text(encoding="utf-8"))
    return {tool["name"] for tool in surface["tools"]}


def _conformance_counts() -> tuple[int, int]:
    """Total cases, and the number of them that assert a refusal."""
    corpus = json.loads((CONTRACT / "conformance.json").read_text(encoding="utf-8"))
    cases = corpus["cases"]
    refusals = [c for c in cases if any("error_contains" in step for step in c["steps"])]
    return len(cases), len(refusals)


def _names_in(doc: str) -> set[str]:
    return set(TOOL_NAME.findall(_read(doc))) - NOT_TOOLS


def test_every_tool_is_listed_where_a_reader_looks_for_it() -> None:
    tools = _contract_tool_names()
    for doc in DOCS_LISTING_EVERY_TOOL:
        named = _names_in(doc)
        families = tuple(name[:-1] for name in named if name.endswith("*"))
        covered = named | {tool for tool in tools if tool.startswith(families)}
        missing = sorted(tools - covered)
        assert not missing, (
            f"{doc} does not mention {', '.join(missing)}. A tool an agent can call "
            "and a reader cannot find is a tool nobody knows about."
        )


def _is_known(name: str, tools: set[str]) -> bool:
    if name.endswith("*"):
        return any(tool.startswith(name[:-1]) for tool in tools)
    return name in tools


def test_no_document_names_a_tool_that_does_not_exist() -> None:
    tools = _contract_tool_names()
    for doc in DOCS_WITH_LIVE_TOOL_NAMES:
        unknown = sorted(name for name in _names_in(doc) if not _is_known(name, tools))
        assert not unknown, (
            f"{doc} names {', '.join(unknown)}, which the contract does not. "
            "Renaming a tool means renaming it in the documents too."
        )


def test_every_relative_link_points_at_something_that_exists() -> None:
    # The package page had a `../README.md` in it once, which resolves in the
    # repository and 404s on PyPI, where the reader is one directory that does
    # not exist. Links off the package page are absolute for that reason; the
    # ones inside the repository have to resolve.
    link = re.compile(r"\]\(([^)#]+?)(?:#[^)]*)?\)")
    for doc in DOCS_WITH_LIVE_TOOL_NAMES:
        source = REPO_ROOT / doc
        for target in link.findall(_read(doc)):
            if "://" in target or target.startswith(("mailto:", "#")):
                continue
            resolved = (source.parent / target).resolve()
            assert resolved.exists(), f"{doc} links to {target}, which does not exist."
            assert REPO_ROOT in resolved.parents or resolved == REPO_ROOT, (
                f"{doc} links to {target}, which is outside the repository."
            )


def test_the_registry_entry_agrees_with_the_package() -> None:
    """`server.json` is what the MCP registry publishes, and three of its fields
    restate something the package already says.

    The one that matters is the ownership marker. The registry proves a PyPI
    package belongs to a server by fetching the published description and
    finding `mcp-name: <server name>` in it. That marker lives in the package
    README and the name lives in server.json, which is the same two-part write
    that breaks everything else here: change one, and the registry publish fails
    on a release that has already gone out to PyPI and cannot be taken back.
    """
    from xlide_mcp import __version__

    entry = json.loads((REPO_ROOT / "server.json").read_text(encoding="utf-8"))
    package = entry["packages"][0]

    assert entry["version"] == __version__, (
        f"server.json says version {entry['version']}, the package is {__version__}."
    )
    assert package["version"] == __version__, (
        f"server.json's package says {package['version']}, the package is {__version__}."
    )
    assert package["identifier"] == "xlide-mcp"

    readme = _read("python/README.md")
    marker = f"mcp-name: {entry['name']}"
    assert marker in readme, (
        f"python/README.md is the published PyPI description and must carry "
        f"'{marker}', or the registry cannot prove the package is this server's."
    )


def test_a_stated_conformance_count_matches_the_corpus() -> None:
    total, refusals = _conformance_counts()
    # Collapse the wrapping first: these sentences are hard-wrapped, and a count
    # split across two lines would otherwise read as absent rather than wrong.
    pattern = re.compile(r"(\d+) of the (\d+) conformance cases")
    plain = re.compile(r"(\d+) conformance cases")

    for doc in DOCS_WITH_LIVE_TOOL_NAMES:
        text = " ".join(_read(doc).split())

        for stated_refusals, stated_total in pattern.findall(text):
            assert (int(stated_refusals), int(stated_total)) == (refusals, total), (
                f"{doc} says {stated_refusals} of {stated_total} conformance cases "
                f"assert a refusal. The corpus has {refusals} of {total}."
            )

        for stated_total in plain.findall(re.sub(pattern.pattern, "", text)):
            assert int(stated_total) == total, (
                f"{doc} says {stated_total} conformance cases. The corpus has {total}."
            )
