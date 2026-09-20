"""xlide-mcp: an MCP server for the code inside Office files.

The VBA, the UserForms, the Power Query and the worksheet cells in Excel, Word,
PowerPoint and Access documents, reachable by any MCP-capable agent. Reading and
writing the file needs no Office installation and works on any platform; running
macros and tests needs Windows with the application installed.

    from xlide_mcp import build_server

    build_server().run("stdio")
"""

from __future__ import annotations

from .config import Settings
from .server import __version__, build_server

__all__ = ["Settings", "__version__", "build_server"]
