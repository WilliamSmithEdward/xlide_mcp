# Security policy

## Report a vulnerability

Use [GitHub private vulnerability reporting](https://github.com/WilliamSmithEdward/xlide_mcp/security/advisories/new). Please do not open a public issue for a possible vulnerability. Include the xlide-mcp version, Python and operating system versions, the Office file type, what you did, what happened, and a minimal file or script that reproduces it. Remove private data from a sample file before attaching it.

## Supported versions

Only the latest release on PyPI receives fixes. Security fixes ship in a new release rather than backports.

## Scope

The server reads and writes Office files that may be untrusted. A crafted file that escapes a workspace root, changes an unintended file, bypasses a guarded write, corrupts a project, or makes a read execute code is in scope. So is a path that escapes its configured root through a symlink, or a request that touches an Office process the server did not create unless the caller explicitly chose one of the office tools.

Running a macro runs VBA with the user's Office permissions. The run supervisor enforces a deadline and owns its Office instance; it is not a sandbox for an untrusted macro. Use only macros you trust, or run them in an isolated environment.

## Automated checks

CodeQL runs its security-extended queries over Python and GitHub Actions. Semgrep runs its Python, security, secrets and GitHub Actions rulesets. ClamAV and YARA-X scan every tracked file for malware and suspicious content. ClamAV fetches and verifies the current official signatures on every run and also flags any Office file with VBA and any encrypted file. YARA-X uses the full YARA Forge rule set, which gathers the public YARA collections, pinned to a release and its SHA-256; a weekly workflow proposes the next release in a pull request that the security workflow scans before it is merged. All four scan every push to main, every pull request, daily, and before a release is published. An unlisted finding, a scan warning, an incomplete scan, or an accepted finding that no longer matches fails the security workflow. A reviewer may list a false positive in [.github/security/accepted.toml](.github/security/accepted.toml) only with a specific source line, or for a ClamAV or YARA-X hit the file's SHA-256, and a reason. The one accepted today is the test workbook that holds VBA on purpose.

Every action is pinned to a commit SHA, the ClamAV image to a digest, and each downloaded tool to an exact release checked against its SHA-256.

The release workflow waits for the security workflow before publishing to PyPI or creating a GitHub release. Each release carries a versioned security report and the raw SARIF results as assets. GitHub secret scanning with push protection and Dependabot security updates are enabled. Dependabot also proposes package and workflow updates for review. Publishing uses PyPI Trusted Publishing; no PyPI upload token is stored in the repository.
