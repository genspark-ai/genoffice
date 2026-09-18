# CLI / MCP matrix

Consolidated index of the `genoffice` command line and the MCP tools that run
the same commands. Derived from [packages/cli/README.md](../packages/cli/README.md),
the [tool table](../packages/cli/src/mcp/tools.ts), the
[deck tools](../packages/cli/src/mcp/deck.ts), and the
[MCP server](../packages/cli/src/mcp/server.ts). No new behavior is described
here; for option details see the CLI README and `genoffice <command> --help`.

Every command prints a one-line human summary by default or a single JSON
object with `--json`. See the CLI README for the envelope, warning, and exit
code contract (`0` ok, `1` usage, `2` file, `3` conversion failed, `4` app not
available).

## CLI command matrix

| CLI command                       | Purpose (from the CLI README)                              |
| --------------------------------- | ---------------------------------------------------------- |
| `genoffice info <file>`           | Metadata and structure summary of a document               |
| `genoffice convert <file> --to …` | Convert between pdf, docx, xlsx, pptx, md, html, csv       |
| `genoffice open <file>`           | Open a document in the GenOffice app for the user          |
| `genoffice create --type docx`    | New Word document from Markdown or restricted HTML         |
| `genoffice create --type xlsx`    | New workbook from a 2-D array, sheets object, CSV, or JSON |
| `genoffice create --type pptx`    | New deck from ops or from a deck spec directory            |
| `genoffice create --type pdf`     | Print a document to PDF through the hidden renderer        |
| `genoffice docs read`             | Blocks, text preview, comments, revisions, header/footer   |
| `genoffice docs apply`            | Edit a docx with a batch of ops                            |
| `genoffice docs check`            | Consistency checks on a docx (exit 0)                      |
| `genoffice sheet read`            | Cells, formulas, and sheet features of a worksheet         |
| `genoffice sheet apply`           | Edit an xlsx by cells or by workbook DSL ops               |
| `genoffice sheet check`           | Consistency checks on an xlsx (exit 0)                     |
| `genoffice slides read`           | Slide and element structure with durable ids and geometry  |
| `genoffice slides apply`          | Edit a pptx with a batch of ops                            |
| `genoffice slides audit`          | Layout audit: overflow, out-of-bounds, overlap             |
| `genoffice slides render`         | One PNG per slide through the app PDF export               |
| `genoffice slides check`          | Validate an outline.json or one page spec file             |
| `genoffice slides replace`        | Rebuild one slide from its page spec file                  |
| `genoffice render <file>`         | One PNG per page of any document                           |
| `genoffice guide <domain>`        | Op reference for docs, sheets, or slides                   |
| `genoffice capabilities`          | Which cloud features are configured (no network call)      |
| `genoffice search <query>`        | Web or image search through the configured provider        |
| `genoffice image <prompt>`        | Generate an image and save it                              |
| `genoffice media <source>`        | Describe or extract from an image, audio, or video file    |
| `genoffice mcp`                   | Serve the commands as MCP tools on stdio                   |
| `genoffice install-cli`           | Repeat the PATH install attempt                            |

Routes that need layout (`to pdf`, `docx` to `html`, `html` to `docx`,
`create --type pdf`) run inside the GenOffice binary through its hidden
`--headless-export` mode; see [headless-pdf-export.md](headless-pdf-export.md).

## MCP tool mapping

`genoffice mcp` serves one tool per command verb. Each tool takes its
parameters from the command option list, so the two surfaces cannot drift.
Ops, cell lists, specs, and Markdown are passed inline and land in a scratch
directory for the length of the call. `render` and `slides_render` return the
PNGs as image content. Results use the same JSON envelope `--json` prints.

| MCP tool         | CLI command it runs  |
| ---------------- | -------------------- |
| `info`           | `info`               |
| `convert`        | `convert`            |
| `create_docx`    | `create --type docx` |
| `create_xlsx`    | `create --type xlsx` |
| `create_pdf`     | `create --type pdf`  |
| `create_pptx`    | `create --type pptx` |
| `docs_read`      | `docs read`          |
| `docs_apply`     | `docs apply`         |
| `docs_check`     | `docs check`         |
| `sheet_read`     | `sheet read`         |
| `sheet_apply`    | `sheet apply`        |
| `sheet_check`    | `sheet check`        |
| `slides_read`    | `slides read`        |
| `slides_apply`   | `slides apply`       |
| `slides_audit`   | `slides audit`       |
| `slides_render`  | `slides render`      |
| `slides_check`   | `slides check`       |
| `slides_replace` | `slides replace`     |
| `render`         | `render`             |
| `guide`          | `guide`              |
| `capabilities`   | `capabilities`       |
| `search`         | `search`             |
| `image`          | `image`              |
| `media`          | `media`              |
| `open`           | `open`               |

The staged deck flow is the same state from either side (style sheet,
outline, one page spec file per slide). The CLI side is `slides check`,
`create --spec <dir>`, and `slides replace`; the MCP side is:

| MCP deck tool  | CLI equivalent                                  |
| -------------- | ----------------------------------------------- |
| `deck_start`   | Write style sheet + outline, run `slides check` |
| `deck_page`    | Write and check one page file                   |
| `deck_build`   | `create --type pptx --spec <dir>`               |
| `deck_replace` | `slides replace`                                |

Op references are also resources: `genoffice://guide/docs`,
`genoffice://guide/sheets`, `genoffice://guide/slides`,
`genoffice://guide/slides/design`, `genoffice://guide/slides/spec`.

## Path policy and audit log summary

From the Path policy and audit log section of the CLI README:

- `GENOFFICE_ALLOWED_ROOTS` (PATH-style list of directories) confines every
  file genoffice reads or writes to those trees; symlinks are resolved before
  the check. A path outside exits 2 with the roots in `detail.allowed_roots`.
  Unset means unrestricted.
- `apply --out` onto another existing file needs `--force`, like `create` and
  `convert`; editing in place never does. Unknown options are rejected.
- A file the running GenOffice shell has open in a tab is not rewritten in
  place (exit 2, `detail.gui_pid`); the shell publishes open tabs to
  `userData/open-documents.json` (`GENOFFICE_USER_DATA` overrides the
  location). `--force` writes anyway.
- Every executed command appends one JSON line (`ts`, `command`, `argv`,
  `status`, `code`, `output_path`, `ms`, `cwd`) to
  `~/.genoffice/cli-audit.jsonl`, rotated at 2 MB.
  `GENOFFICE_AUDIT_LOG=<path>` redirects it, `GENOFFICE_AUDIT_LOG=off`
  disables it.
