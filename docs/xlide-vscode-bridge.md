# Telling XLIDE for VS Code what a tool changed

When this server runs inside VS Code beside the XLIDE extension, a module an
agent writes through it should look like one an agent wrote through XLIDE's own
tools: marked in the project tree, with a diff of before and after, and Keep and
Revert on it. XLIDE already keeps that review for its own writes. It cannot see
this server's writes except as a file that changed on disk, which carries no
before, so this server sends what it has, over a channel XLIDE opens.

Everything here is best effort on the server's side. A write never waits more
than a second per XLIDE window, never fails because of one, and with no XLIDE
running costs a directory listing. The server's half is `xlide_vscode.py`.

## Finding XLIDE

XLIDE listens on a loopback port behind a token, and says where in one of two
ways. The server uses both.

1. **`XLIDE_VSCODE_API`**, when XLIDE starts this server itself, for instance as
   the MCP server it contributes: `http://127.0.0.1:{port}/{token}`. A value
   that is not an `http` URL on `127.0.0.1`, `localhost` or `::1` is ignored.
2. **A discovery file per VS Code window**, for a server the user configured
   themselves: `xlide-api-{pid}.json` in `%LOCALAPPDATA%\xlide_vscode` on
   Windows, and in `$XDG_STATE_HOME/xlide_vscode` (default
   `~/.local/state/xlide_vscode`) elsewhere.

   ```json
   {
     "pid": 4242,
     "port": 51234,
     "token": "c2d1...",
     "product": "xlide_vscode",
     "version": "10.8.0",
     "protocol": 1,
     "workspaceFolders": ["C:\\work\\budget"]
   }
   ```

   A window deletes its file when it closes. A file can outlive a window that
   was killed, so the server believes no file until the window answers
   `hello`. What it finds is kept for ten seconds.

This is the arrangement the VBE add-in, xlide_vbide, already uses for the
`xlide_live_*` tools: the token is the security model, the port is loopback
only, and anything running as the user can read the discovery file. The server
reaches the port directly and never through a proxy, which a port has to do on
purpose: Python's urllib, for one, sends 127.0.0.1 through `HTTP_PROXY` like
any other host.

## Routes

Every route is `{base}/{route}`, where base is `http://127.0.0.1:{port}/{token}`.
Bodies are JSON. The server sends `Content-Type: application/json` and reads a
JSON object back; anything else counts as no answer.

### `GET hello`

```json
{"product": "xlide_vscode", "protocol": 1}
```

The only proof a window is alive. Anything other than `product: xlide_vscode`
means the window is skipped.

### `POST agent-edit`

A module this server wrote, created or deleted.

```json
{
  "file": "C:\\work\\budget\\Budget.xlsm",
  "module": "Helpers",
  "before": "Public Sub Old()\r\nEnd Sub\r\n",
  "beforeExisted": true,
  "after": "Public Sub New()\r\nEnd Sub\r\n",
  "afterExists": true,
  "kind": "standard",
  "tool": "xlide_write_module",
  "server": "xlide_mcp",
  "serverVersion": "1.1.0",
  "protocol": 1
}
```

- `before` and `after` are the module body as `xlide_read_module` returns it:
  the `Attribute VB_*` header stripped. `after` is read back from the file after
  the save, so it is what XLIDE will find there.
- `beforeExisted: false` is a module the write created; its revert deletes it.
- `afterExists: false` is a module the write deleted, with its text in
  `before` and its kind in `kind` (`standard`, `class`, `document` or
  `userform`), so XLIDE could offer to put it back. Its own deletes keep no
  review, and it may treat this one the same way.
- Sent by `xlide_write_module`, `xlide_delete_module`, and by
  `xlide_import_modules` once per module it changed.

What XLIDE does with it is the same as for its own write: record the review
with the first `before` kept when a module already has one pending, mark the
module row, open the diff when `xlide.agent.showWriteDiffs` is on, and refresh
the project. Its revert guard compares the module's current text with `after`
under its own normalisation (line endings, trailing white space, leading blank
lines), so a write made after this one, by anyone, turns the revert into a
refusal rather than a loss.

Answer:

```json
{"shown": true, "review": "pending"}
```

`shown` is whether this window has the file in its tree. `review` says what
became of the notice. The server reports both in the tool's result, so the
agent can tell the user where to look.

### `POST module-renamed`

```json
{"file": "...", "from": "Helpers", "to": "Tools", "tool": "xlide_rename_module", ...}
```

A pending review moves with the module, as XLIDE's own rename moves it.

### `POST file-changed`

```json
{"file": "...", "what": "document", "tool": "", ...}
```

Anything else this server wrote: `vba`, `forms`, `references`, `queries` or
`document` (cells, formatting, shapes, charts, comments, filters). XLIDE keeps no
review for these; it refreshes the project straight away and records the file's
new stamp. That also covers the case its file watcher misses: the first check of
a project only records a stamp, so the first outside write to a project that is
only shown in the tree changes nothing on screen.

Every body also carries `server`, `serverVersion` and `protocol`.

## Versioning

`protocol` is 1. A change that an older XLIDE would misread gets a new number,
and XLIDE answers `hello` with the numbers it speaks. Adding a field is not such
a change: both sides ignore fields they do not know.
