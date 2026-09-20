"""Write the repo's normative tool contract out of the running Python server.

The contract is derived, never hand-maintained. A hand-written one drifts from the
implementation the first time someone adds an argument and forgets the document,
and a port verified against a drifted contract is verified against nothing.

    python tools/export_contract.py          # write ../contract/tool-surface.json
    python tools/export_contract.py --check   # fail if the file is out of date

`--check` is what CI runs. It is the mechanism that keeps the promise: change a
tool without regenerating, and the build stops.

The upstream versions recorded here are the point of the whole file. The Python
implementation tracks pyOpenVBA, pyvbaanalysis and pyvbaharness; a port tracks the
Python implementation. Stamping which upstream versions a contract reflects is
what turns "we are behind" into a diff someone can act on.
"""

from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
import sys
from importlib import metadata
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[2]
CONTRACT_PATH = REPO_ROOT / "contract" / "tool-surface.json"

# Distribution name to import name. They differ for pyOpenVBA, and the import
# name is what actually answers.
UPSTREAM = {
    "pyOpenVBA": "pyopenvba",
    "pyvbaanalysis": "pyvbaanalysis",
    "pyvbaharness": "pyvbaharness",
}


def build_contract() -> dict[str, Any]:
    from xlide_mcp import Settings, __version__, build_server
    from xlide_mcp.instructions import SERVER_INSTRUCTIONS

    server = build_server(Settings())
    tools = asyncio.run(server.list_tools())

    exported: list[dict[str, Any]] = []
    for tool in sorted(tools, key=lambda t: t.name):
        annotations = tool.annotations
        exported.append(
            {
                "name": tool.name,
                "title": tool.title or "",
                "description": tool.description or "",
                "input_schema": tool.input_schema or {},
                "annotations": {
                    "read_only": bool(annotations and annotations.read_only_hint),
                    "destructive": bool(annotations and annotations.destructive_hint),
                    "idempotent": bool(annotations and annotations.idempotent_hint),
                },
            }
        )

    body: dict[str, Any] = {
        "contract_version": __version__,
        "reference_implementation": "python",
        "server": {"name": server.name, "instructions": SERVER_INSTRUCTIONS},
        "tool_count": len(exported),
        "tools": exported,
        # Which upstream libraries this build was generated against. Recorded so a
        # port can tell whether it is behind, and deliberately outside the digest
        # and outside --check: it describes the machine that ran the export, not
        # the surface. A developer's editable checkout is routinely ahead of the
        # released versions CI installs, and a drift check that failed on that
        # would be a check nobody could keep green.
        "generated_with": {
            "upstream": {
                dist: _version(dist, module) for dist, module in sorted(UPSTREAM.items())
            }
        },
    }
    body["digest"] = _digest(body)
    return body


def _version(distribution: str, module_name: str) -> str:
    """The version actually in use.

    An editable install keeps the distribution metadata from whenever it was
    installed, so for a sibling checkout that moves on, metadata says 1.2.0 while
    the code being exercised is 2.1.1. The imported module is the one that ran.
    """
    try:
        imported = __import__(module_name)
    except ImportError:
        pass
    else:
        declared = getattr(imported, "__version__", "")
        if declared:
            return str(declared)
    try:
        return metadata.version(distribution)
    except metadata.PackageNotFoundError:
        return "not installed"


#  Keys that describe the export rather than the surface. A port pins the digest
#  to mean "this tool surface", so a value that moves without the surface moving
#  must not be in it.
_NOT_NORMATIVE = frozenset({"digest", "generated_with", "contract_version"})


def _digest(body: dict[str, Any]) -> str:
    """A hash of the surface itself: the server and its tools, nothing else."""
    payload = {key: value for key, value in body.items() if key not in _NOT_NORMATIVE}
    encoded = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return "sha256:" + hashlib.sha256(encoded).hexdigest()


def _comparable(contract: dict[str, Any]) -> dict[str, Any]:
    """The part of the contract a drift check is about: the surface, not the export."""
    return {key: value for key, value in contract.items() if key != "generated_with"}


def render(contract: dict[str, Any]) -> str:
    return json.dumps(contract, indent=2, sort_keys=True) + "\n"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="Do not write. Exit 1 if the checked-in contract is out of date.",
    )
    parser.add_argument("--out", type=Path, default=CONTRACT_PATH)
    args = parser.parse_args(argv)

    contract = build_contract()
    rendered = render(contract)

    if args.check:
        if not args.out.is_file():
            print(
                f"{args.out} does not exist. Run: python tools/export_contract.py",
                file=sys.stderr,
            )
            return 1
        try:
            current = json.loads(args.out.read_text(encoding="utf-8"))
        except ValueError as exc:
            print(f"{args.out} is not readable JSON: {exc}", file=sys.stderr)
            return 1
        if _comparable(current) != _comparable(contract):
            print(
                f"{args.out} is out of date: the tool surface changed.\n"
                "Regenerate it with: python tools/export_contract.py\n"
                "Then port the change to every other implementation in this repo.",
                file=sys.stderr,
            )
            return 1
        if args.out.read_text(encoding="utf-8") != rendered:
            # Same surface, different machine. Worth saying, never worth failing.
            print(
                f"{args.out.name} was generated against different upstream versions than "
                "this machine has. The surface itself is current.",
                file=sys.stderr,
            )
        print(f"{args.out.name} is current ({contract['tool_count']} tools).")
        return 0

    args.out.parent.mkdir(parents=True, exist_ok=True)
    # Written LF on every platform. The repository normalizes to LF, so an
    # artifact regenerated on Windows with native line endings would read as
    # modified the moment it was written.
    args.out.write_text(rendered, encoding="utf-8", newline="\n")
    print(f"Wrote {args.out} ({contract['tool_count']} tools, {contract['digest']}).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
