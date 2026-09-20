# The file layer, in a container.
#
# This builds the server from the source in this repository rather than from the
# published wheel, so that what the image runs is what the commit says. Only the
# file layer is here: running macros and tests needs Windows with the desktop
# application, which a Linux container does not have and cannot pretend to.
#
# The workspace root is /workspace. Mount the folder holding the Office files
# there, because `--root` is the boundary every path argument is checked against
# and a server that could reach the whole filesystem would be a different
# product:
#
#     docker run --rm -i -v "$PWD:/workspace" xlide-mcp
#
# Glama builds this image to check that the server starts and answers an
# introspection request, so it has to run with no arguments and no mounted
# volume as well.

FROM python:3.12-slim

# Nothing in the file layer shells out except git, which xlide_git_changes uses
# to read a blob at a revision.
RUN apt-get update \
    && apt-get install --no-install-recommends -y git \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /src
COPY python/ /src/
RUN pip install --no-cache-dir . && rm -rf /src

# A non-root user, because this reads and writes files a caller mounts in.
RUN useradd --create-home --uid 1000 xlide
WORKDIR /workspace
RUN chown xlide:xlide /workspace
USER xlide

# stdio is how a client launches a server it owns. Startup logging goes to
# stderr; stdout is protocol and a stray byte there corrupts the stream.
ENTRYPOINT ["xlide-mcp", "--root", "/workspace"]
