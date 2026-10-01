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

# Pinned by digest to the multi-platform image index; Dependabot proposes new
# 3.12 releases.
FROM python:3.12.14-slim@sha256:f77ac9e44ae96ef2c90b8053ea08c31f8be030f824196b0ae4db6d462c84e51f

# Nothing in the file layer shells out except git, which xlide_git_changes uses
# to read a blob at a revision.
RUN apt-get update \
    && apt-get install --no-install-recommends -y git \
    && rm -rf /var/lib/apt/lists/*

# The dependencies come from a hash-locked file, so every build of the image
# installs the same ones; the package itself then installs without resolving.
WORKDIR /src
COPY .github/requirements/image.txt /tmp/image.txt
COPY python/ /src/
RUN pip install --no-cache-dir --require-hashes -r /tmp/image.txt \
    && pip wheel --no-cache-dir --no-deps --no-build-isolation --wheel-dir /tmp/wheel . \
    && pip install --no-cache-dir --no-deps /tmp/wheel/*.whl \
    && rm -rf /src /tmp/image.txt /tmp/wheel

# A non-root user, because this reads and writes files a caller mounts in.
RUN useradd --create-home --uid 1000 xlide
WORKDIR /workspace
RUN chown xlide:xlide /workspace
USER xlide

# stdio is how a client launches a server it owns. Startup logging goes to
# stderr; stdout is protocol and a stray byte there corrupts the stream.
#
# The root is split out of the entrypoint and into CMD so that it can be
# replaced. `docker run ... xlide-mcp` takes the default below; a runner that
# mounts the caller's folder somewhere of its own choosing passes its own
# `--root` as arguments, and Docker replaces CMD with them.
ENTRYPOINT ["xlide-mcp"]
CMD ["--root", "/workspace"]
