"""Tool groups, one module per layer of reach.

* discovery, modules, analysis, forms, powerquery, sheets, sync, vcs - the
  file itself.
  No Office installation, any platform.
* execution - macros, tests and compile checks in a desktop application this
  server owns. Windows only.
* live - a running xlide_vbide session inside the VBE. Windows, and the add-in.
"""

from __future__ import annotations

__all__ = [
    "analysis",
    "discovery",
    "execution",
    "forms",
    "live",
    "modules",
    "powerquery",
    "sheets",
    "sync",
    "vcs",
]
