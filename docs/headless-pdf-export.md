# Headless PDF export

Every "export to PDF" path in GenOffice can be driven without a visible
editor window, behind one entry point:

```
<app binary> --headless-export <input-file> --to pdf --out <path> [--json]
```

The app creates no visible window, hides the macOS dock icon, writes exactly
one line to stdout and exits. It does **not** take the single-instance lock,
so it runs happily alongside a GUI instance.

```
$ GenOffice --headless-export report.docx --to pdf --out report.pdf --json
{"status":"ok","summary":"Exported /w/report.docx to /w/report.pdf","output_path":"/w/report.pdf"}
```

Without `--json` the same line is a plain sentence. Both `--flag value` and
`--flag=value` are accepted.

## Exit codes

| code | meaning                                                                         |
| ---- | ------------------------------------------------------------------------------- |
| 0    | the PDF was written                                                             |
| 1    | bad arguments (missing `--out`, a target other than `pdf`, no output directory) |
| 2    | input file problem (missing, not a file, extension no module can render)        |
| 3    | conversion failure (unreadable document, renderer crash, timeout)               |

On failure the JSON envelope is
`{"status":"error","summary":…,"error":…}` — one line, newlines folded out.

## Supported inputs

| extension                     | module   |
| ----------------------------- | -------- |
| `.docx`                       | docs     |
| `.xlsx` `.xlsm` `.xls` `.csv` | sheets   |
| `.pptx`                       | slides   |
| `.md` `.markdown`             | markdown |
| `.html` `.htm`                | html     |

## How it works

The grammar, the exit codes and the renderer-side waiting helpers live in
`packages/electron-utils/src/headless-export.ts`; the renderer half is
imported through the `@genoffice/electron-utils/headless-export` subpath so a
renderer bundle never pulls in `node:` builtins.

`apps/shell/src/main/headless-export.ts` validates the paths and routes by
extension to one hidden-window exporter per module
(`exportDocsPdfHeadless`, `exportSheetsPdfHeadless`, …). Each exporter:

1. creates a `show: false` BrowserWindow with that module's normal preload,
   sandbox and `backgroundThrottling: false`;
2. queues the input through the module's existing pending-open path, so the
   renderer runs its ordinary load pipeline;
3. lets the renderer wait until the document has settled (docs: pagination
   slice count stable twice in a row; slides: every picture decoded and the
   private Office fonts registered; sheets: the workbook fully preloaded) and
   then call **the same export function the File menu calls**, with the CLI's
   path in place of the save dialog;
4. resolves on the renderer's report, and destroys the window.

There is one export function per module — the GUI passes no path and gets a
dialog, headless passes one and does not — so GUI and CLI output cannot
drift. Measured on four documents, the two paths produce PDFs that differ
only in the embedded creation timestamp.

`isHeadlessMode()` (`packages/electron-utils/src/headless-mode.ts`) is set
before any editor module boots; it suppresses everything that would surface
UI after a successful write (opening the export in a tab, revealing it in the
file manager) and the slides recent-files entry.

## Behaviour worth knowing

- **An unreadable input is an error, not a blank page.** docs and slides
  answer a corrupt file with an untitled blank document; headless treats "a
  document is mounted but it never came from disk" as exit 3.
- **Password-protected documents** cannot be answered without a user, so they
  fail on the readiness timeout rather than immediately.
- **Large workbooks** normally wait for the user's "Full Load" click before
  PDF export is allowed. Headless makes that call itself.
- **Only the active sheet** is exported for spreadsheets, exactly as in the
  GUI.
