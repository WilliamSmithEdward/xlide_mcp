# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/xlide_mcp/security/advisories/new).
Only the maintainer sees it. Include the xlide-mcp version, the Python and
operating system versions, the Office file type, what you did and what
happened, and the smallest file or script that shows it, with credentials
and private data removed.

A confirmed vulnerability is fixed in a release on PyPI, and the advisory
is published with it, crediting you unless you ask otherwise.

## Supported versions

Only the latest release on PyPI receives security fixes. Older releases are
not maintained separately; update when a fix ships.

## Scope

The server reads and writes Office files that may be untrusted. A crafted
file that escapes a workspace root, changes an unintended file, bypasses a
guarded write, corrupts a project, or makes a read execute code is in
scope. So is a path that escapes its configured root through a symlink, or
a request that touches an Office process the server did not create unless
the caller explicitly chose one of the office tools. The module files a VB6
`.vbp` names are held to the roots too, after symlinks resolve: a project
naming one outside them is refused rather than read or written.

### Running macros

Running a macro runs VBA with the user's Office permissions. The run
supervisor enforces a deadline and owns its Office instance; it is not a
sandbox for an untrusted macro. Use only macros you trust, or run them in
an isolated environment.

### Running the container image

The `Dockerfile` builds the file layer from the source in this repository,
not from the published wheel. It runs as a non-root user with `--root
/workspace`, the boundary every path argument is checked against, so mount
only the folder that holds the Office files there:

```
docker run --rm -i -v "$PWD:/workspace" xlide-mcp
```

The image cannot run macros or tests, which need Windows with the desktop
application. Its only tool beyond Python is git, which `xlide_git_changes`
uses to read a file at a revision.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security and Malware scan also run daily, and again from the
Publish workflow before every release.

- **Code:** CodeQL with the `security-extended` queries, for Python and the
  GitHub Actions workflows, and Semgrep with the `p/default`, `p/python`,
  `p/security-audit`, `p/secrets` and `p/github-actions` rule sets, both
  over `python/src`, `python/tools`, `scripts/security` and `.github`.
  Semgrep ignores `nosemgrep` comments and gives every rule time to finish
  on every file. A scan warning or an incomplete scan fails Security as a
  finding does. Results go to the repository's code scanning.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** pip-audit checks the runtime dependencies, `mcp`,
  pyOpenVBA, pyOfficeEditor and pyvbaanalysis with everything they pull
  in, as a fresh install resolves them today
  (`.github/requirements/runtime.txt`, hash-locked and moved daily by
  Dependabot). Any known vulnerability fails Security. The tools the
  workflows install come from hash-locked files (see Pinning and updates).
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan every tracked file, and the wheel and sdist built from
  it with the hash-locked build tools, as a release builds them. ClamAV
  also flags any Office file that holds VBA and any encrypted file. YARA-X
  runs the full YARA Forge pack. A failed signature update, a scanner error
  or a scan warning fails Malware scan.
- **Fuzzing:** Atheris fuzzes the text this server parses itself
  (`python/fuzz/fuzz_server.py`): the drawing layer's tag scanner and entity
  decoder, cell and range references, and the VB6 project manifest. A
  `ToolError` is the expected answer to bad input; anything else is a
  finding. Each target starts from its seeds in `python/tests/fuzz_corpus`,
  which the test suite also replays. The Fuzz workflow runs on every change
  to `python/src`, the fuzz target or its corpus, for a minute per target,
  and daily for five. It is not a gate: a finding becomes a regression test
  with its fix. The Office containers are read by pyOpenVBA, pyOfficeEditor
  and pyVBAanalysis, which fuzz their own readers.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Some of its checks assume more than one maintainer, such as a second
  person approving every change, so a single-maintainer project cannot
  score full marks on them.

## Accepted findings

A finding is fixed, or accepted with a written reason in
[.github/security/accepted.toml](.github/security/accepted.toml) (CodeQL and
Semgrep) or
[.github/security/malware-accepted.toml](.github/security/malware-accepted.toml)
(ClamAV and YARA-X). An entry matches the tool, the rule, the file and
either the exact source line or, for a malware hit, the file's SHA-256, so
a changed file needs another review, and an entry that no longer matches
fails the report. zizmor keeps its exceptions in `.github/zizmor.yml` or
inline beside the line they excuse, each with its reason. The current
entries:

- Semgrep `subprocess-injection` on `subprocess.run(` in
  `python/src/xlide_mcp/office_apps.py`: the command is a fixed list (the
  current Python, `-m` and this package's `office_apps` module), the
  caller's request goes to it as JSON on stdin, and no shell is used.
- ClamAV `Heuristics.OLE2.ContainsMacros.VBA` on
  `python/tests/fixtures/shapes.xlsm`: a test workbook that holds VBA on
  purpose, a Forms button whose macro the shape tests read and repoint.
- zizmor `self-repository`, turned off in
  [.github/zizmor.yml](.github/zizmor.yml) until GitHub's documentation
  confirms the `$/` self-repository syntax for reusable workflows called
  from `publish.yml`.
- zizmor `superfluous-actions`, turned off in the same file: it asks for
  `gh release create` in place of `softprops/action-gh-release`, and the
  release path changes only once that change can be dry-run.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, scanner images to digests, Python tools to
hash-locked lock files, the runtime dependencies the tests and the
container image install to hash-locked lock files as well, and the YARA-X
engine and YARA Forge rules to a release and its SHA-256. The container
image's Python base is pinned to a digest. The `mcp-publisher` release the
Publish workflow downloads is pinned to a release and its SHA-256 and moved
by hand. ClamAV's signatures change too often to pin, so freshclam fetches
and verifies them on every run. The Semgrep rule sets are fetched from the
Semgrep registry on every run.

Dependabot proposes updates to the GitHub Actions, the scanner images, the
container image's Python base (3.12 releases only), the dependencies in
`python/pyproject.toml` and the lock files in `.github/requirements` once a
version is a week old, and at once for a security advisory. The Update YARA
rules workflow proposes new YARA pins each week. A minor or patch update,
and the YARA pull request, merges itself once CI, Security and Malware scan
pass; a third-party major version waits for review.

## Releases

Pushing a `v*.*.*` tag starts the Publish workflow. It checks that the tag
matches the version in `python/src/xlide_mcp/_version.py`, runs the tests,
checks that the generated tool contract and conformance files are current,
builds the wheel and sdist and checks their metadata. It runs Security and
Malware scan on the tagged commit, and uploads to PyPI through trusted
publishing only when the build and both scans pass, so no upload token is
stored anywhere. It then publishes the server to the MCP Registry, proving
the namespace with the workflow's OIDC identity, and creates the GitHub
release. Started by hand, it is a dry run that publishes nothing.

The GitHub release carries the wheel and sdist, the security and malware
reports and their raw SARIF results, each named `xlide-mcp-<version>-*`,
and the signed provenance bundle `xlide-mcp-<version>.sigstore.json`.
Releases up to v1.2.2 carry one security report covering all four scans,
and no provenance bundle.

### Verifying a download

Every file on PyPI carries PyPI's own provenance, which names this
repository's `publish.yml` as the publisher; the file's page on PyPI shows it.
Releases published after 2026-09-30 also carry a GitHub build provenance
attestation, which you can check against any copy of the file, from PyPI or
from the GitHub release:

```
pip download xlide-mcp --no-deps -d check
gh attestation verify check/<file> --owner WilliamSmithEdward
```

The output names the commit and workflow run that built the file. The
signed bundle is also attached to the GitHub release as
`xlide-mcp-<version>.sigstore.json`, so the check works without asking
GitHub for it: add `--bundle xlide-mcp-<version>.sigstore.json`.

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
