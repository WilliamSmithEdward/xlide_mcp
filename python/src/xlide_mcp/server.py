"""Builds the MCP server and registers every tool on it.

Tools are grouped into modules by what they reach, because that is also how they
fail: the file layer works anywhere, the execution layer needs Windows with the
application installed, and the live layer needs an xlide_vbide session running.
Each group registers itself, so a group can be left out of a build without the
rest noticing.

Registration is unconditional even for the layers that cannot work on this
machine. A tool that is present and explains what is missing is more use to an
agent than a tool that is absent, which reads as "this server cannot do that at
all" and sends it looking for another way in.
"""

from __future__ import annotations

from mcp.server.mcpserver import MCPServer

from ._version import __version__
from .config import Settings
from .instructions import SERVER_INSTRUCTIONS
from .tools import (
    analysis,
    catalog,
    discovery,
    execution,
    features,
    formatting,
    forms,
    live,
    modules,
    powerquery,
    sheets,
    structure,
    sync,
    vcs,
)

_REGISTRARS = (
    discovery.register,
    modules.register,
    analysis.register,
    catalog.register,
    forms.register,
    powerquery.register,
    sheets.register,
    formatting.register,
    structure.register,
    features.register,
    sync.register,
    vcs.register,
    execution.register,
    live.register,
)


def build_server(settings: Settings | None = None) -> MCPServer:
    """A server with every tool registered, ready for any transport."""
    resolved = settings or Settings.from_environment()
    server = MCPServer(
        name="xlide",
        title="XLIDE: VBA and Office files",
        version=__version__,
        instructions=SERVER_INSTRUCTIONS,
        website_url="https://github.com/WilliamSmithEdward/xlide_mcp",
    )
    for register in _REGISTRARS:
        register(server, resolved)
    _register_resources(server, resolved)
    return server


def _register_resources(server: MCPServer, settings: Settings) -> None:
    """Things an agent may want to read once rather than call repeatedly."""

    @server.resource(
        "xlide://instructions",
        name="Agent instructions",
        title="How to work on Office files with this server",
        description=(
            "The workflow and the rules, as text the user can paste into CLAUDE.md, "
            "AGENTS.md or .github/copilot-instructions.md."
        ),
        mime_type="text/markdown",
    )
    def instructions_resource() -> str:
        return SERVER_INSTRUCTIONS

    @server.resource(
        "xlide://workspace",
        name="Workspace",
        title="What this server is allowed to reach",
        description="The roots every path argument is resolved inside, and the write mode.",
        mime_type="text/plain",
    )
    def workspace_resource() -> str:
        return settings.describe()
