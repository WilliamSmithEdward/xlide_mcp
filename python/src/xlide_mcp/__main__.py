"""Command line: `xlide-mcp` and `python -m xlide_mcp`.

stdio is the default, because that is how an MCP client launches a server it owns.
The HTTP transports are there for a client that connects to a server someone else
started, and they bind to loopback unless told otherwise: this server reads and
writes files, and a default that listened on every interface would be handing that
reach to the network.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
from collections.abc import Sequence
from pathlib import Path

from .config import MAX_TIMEOUT_SECONDS, Settings, roots_from_argv
from .errors import ToolError
from .hosts import host_info
from .server import __version__, build_server


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlide-mcp",
        description=(
            "MCP server for the VBA, forms, Power Query and cells inside Office files."
        ),
        epilog=(
            "Every path a tool accepts is resolved inside the roots, so --root is how you "
            "say which folder an agent may reach."
        ),
    )
    parser.add_argument("--version", action="version", version=f"xlide-mcp {__version__}")
    parser.add_argument(
        "roots",
        nargs="*",
        metavar="PATH",
        help=(
            "Folders the server may reach, the same as --root and combined with it. "
            "Positional because a launcher that mounts the caller's folders somewhere "
            "of its own choosing appends them as plain arguments, and repeating a flag "
            "in front of each one is not something it can do."
        ),
    )
    parser.add_argument(
        "--textconv",
        metavar="PATH",
        help=(
            "Print the file's VBA and Power Query as text and exit, instead of starting "
            "the server. This is a git textconv driver: point one at it and git diff, "
            "git show and git log -p render an Office file as source rather than "
            "reporting that two binaries differ."
        ),
    )
    parser.add_argument(
        "--root",
        action="append",
        metavar="PATH",
        help=(
            "A folder the server may reach. Repeatable, and each value may hold several "
            "paths separated by the platform's path separator. Defaults to the working "
            "directory."
        ),
    )
    parser.add_argument(
        "--read-only",
        action="store_true",
        help="Refuse every tool that changes a file. Reads and analysis still work.",
    )
    parser.add_argument(
        "--allow-outside-roots",
        action="store_true",
        help="Accept absolute paths anywhere on this machine. Off by default.",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        metavar="SECONDS",
        help=f"Default deadline for a macro or test run. Held at {MAX_TIMEOUT_SECONDS:g}s.",
    )
    parser.add_argument(
        "--transport",
        choices=("stdio", "streamable-http", "sse"),
        default="stdio",
        help="How the client reaches this server. Default: stdio.",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Interface for an HTTP transport. Default: loopback only.",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8765,
        help="Port for an HTTP transport. Default: 8765.",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.textconv is not None:
        return textconv(args.textconv)

    settings = Settings.from_environment()
    roots = roots_from_argv(list(args.root or []) + list(args.roots or []))
    if roots:
        settings = settings.with_roots(roots)
    if args.read_only:
        settings = _replace(settings, read_only=True)
    if args.allow_outside_roots:
        settings = _replace(settings, allow_outside_roots=True)
    if args.timeout is not None:
        settings = _replace(
            settings, default_timeout=max(1.0, min(MAX_TIMEOUT_SECONDS, args.timeout))
        )

    server = build_server(settings)
    if args.transport == "stdio":
        # Anything written to stdout is protocol, so a stray print would corrupt
        # the stream. Startup goes to stderr, where a client shows it as a log.
        print(f"xlide-mcp {__version__} | {settings.describe()}", file=sys.stderr)
        server.run("stdio")
        return 0

    print(
        f"xlide-mcp {__version__} on http://{args.host}:{args.port} | {settings.describe()}",
        file=sys.stderr,
    )
    server.run(args.transport, host=args.host, port=args.port)
    return 0


def textconv(raw: str) -> int:
    """Render one file for git, and never fail.

    git aborts the whole diff on a non-zero exit, so every path here ends in a
    printed document and a zero. It also hands over a temporary file with no
    useful name when the blob comes from history, which is why the extension is
    sniffed from the bytes when the name does not carry one.

    The workspace roots deliberately do not apply. This is not the server: it is
    a local filter git runs on a file the user already has open in their own
    repository, and refusing it there would only break their diff.
    """
    from .hosts import sniff_extension
    from .textual import render

    path = Path(raw)
    if not path.exists():
        print(f"' xlide-mcp: no such file: {raw}")
        return 0

    info = None
    try:
        info = host_info(path)
    except ToolError:
        info = None
    if info is None:
        sniffed = sniff_extension(path)
        if sniffed is not None:
            # The reader is chosen by extension, and pyOpenVBA opens a path, so
            # the blob is copied to a name that says what it is.
            with tempfile.TemporaryDirectory() as directory:
                named = Path(directory) / f"blob{sniffed}"
                shutil.copyfile(path, named)
                sys.stdout.write(render(named))
            return 0

    sys.stdout.write(render(path, info))
    return 0


def _replace(settings: Settings, **changes: object) -> Settings:
    from dataclasses import replace

    return replace(settings, **changes)  # type: ignore[arg-type]


if __name__ == "__main__":
    raise SystemExit(main())
