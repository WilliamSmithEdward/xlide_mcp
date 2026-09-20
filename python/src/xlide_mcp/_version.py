"""The version, in one place.

Kept in a module of its own, with no imports, so the build backend can read it
from the source without importing the package: a version read by import would
need mcp and pydantic present before the package could be built.

`python/pyproject.toml` declares `version` dynamic and points at this attribute,
so a release bumps one line and the packaging metadata, the server's `version`
field and the generated contract all follow.
"""

__version__ = "0.1.0"
