# Docker: headless batch conversion

A self-contained image that converts a whole tree of Office/Markdown/HTML
documents to PDF with GenOffice's own headless export pipeline
(`<app binary> --headless-export`, see `docs/headless-pdf-export.md`). It wraps
the official release deb — no GenOffice source is rebuilt — so every output is
what the desktop app's File ▸ Export would produce.

This is the batch-conversion slice of the Docker ask (genoffice#1808); it is
a CLI utility image, not the server-side/collaborative form discussed there.

## Build

```sh
docker build -t genoffice-batch packaging/docker
```

On Apple silicon / other non-amd64 hosts the build runs the amd64 deb under
qemu (`--platform linux/amd64` is picked automatically by the deb's arch when
using buildx); it is slower to build but the produced image is amd64, which is
what the release deb ships.

## Convert a directory

Mount the documents at `/data` (read-only is fine) and an empty directory at
`/output`. The output tree mirrors the input tree with `.pdf` extensions:

```sh
docker run --rm -v "$PWD/docs:/data:ro" -v "$PWD/pdf:/output" genoffice-batch
```

```
==> reports/q1.docx
{"status":"ok","summary":"Exported /data/reports/q1.docx to /output/reports/q1.pdf","output_path":"/output/reports/q1.pdf"}
...
{"status":"ok","total":7,"converted":7,"failed":0}
```

Exit codes: `0` all converted, `1` usage, `2` no convertible input found,
`3` some files failed (their `==> <file> (failed, exit N)` blocks carry the
upstream error envelopes).

## Knobs (environment variables)

| variable                                  | default   | meaning                                                                                                                                                                |
| ----------------------------------------- | --------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `GENOFFICE_INPUT_DIR`                     | `/data`   | input volume                                                                                                                                                           |
| `GENOFFICE_OUTPUT_DIR`                    | `/output` | output volume                                                                                                                                                          |
| `GENOFFICE_TO`                            | `pdf`     | export target — the image only wires up `pdf` (the app also exports docs to html and html to docx); anything else is rejected with exit 1 rather than silently ignored |
| `GENOFFICE_PARALLEL`                      | `1`       | concurrent conversions (each is its own Electron process; 2-4 helps on large hosts)                                                                                    |
| `GENOFFICE_INCLUDE` / `GENOFFICE_EXCLUDE` | —         | optional `grep -E` filters on `/data`-relative paths                                                                                                                   |
| `GENOFFICE_FLAGS`                         | —         | extra flags passed through to `--headless-export`                                                                                                                      |
| `GENOFFICE_XVFB`                          | `1`       | set `0` when you provide the display yourself (e.g. wrap with `xvfb-run`)                                                                                              |

Examples:

```sh
# only the reports folder, four conversions at a time
docker run --rm -v "$PWD/docs:/data:ro" -v "$PWD/pdf:/output" \
  -e GENOFFICE_INCLUDE='^reports/' -e GENOFFICE_PARALLEL=4 \
  genoffice-batch

# skip draft documents
docker run --rm -v "$PWD/docs:/data:ro" -v "$PWD/pdf:/output" \
  -e GENOFFICE_EXCLUDE='draft' genoffice-batch
```

## Behaviour worth knowing

- **Convertible inputs** are exactly the headless-export table:
  `.docx`, `.xlsx`/`.xlsm`/`.xls`/`.csv`/`.tsv`, `.pptx`, `.md`/`.markdown`,
  `.html`/`.htm`. PDF files in the input tree are deliberately not re-ingested.
- **Name collisions** (`a.md` and `a.docx` both want `a.pdf`) resolve to
  `a.md.pdf` / `a.docx.pdf` for the second claimant, deterministically.
- **Sandbox**: the container passes `--no-sandbox` because Chromium's setuid
  helper cannot work inside image layers; the process runs as uid 1000 (map
  your volumes accordingly, or override with `--user`).
- **Fonts**: Noto CJK/Core/Emoji are installed so CJK documents do not render
  as tofu; mount extra font files into `/usr/local/share/fonts` for anything
  exotic.
- **Updates are off by design**: the deb's `app-update.yml` is removed in the
  image (the same mechanism that disables the updater in fork/PR builds), so
  the image itself is the update channel — rebuild on a new release and bump
  `GENOFFICE_VERSION` + `GENOFFICE_DEB_SHA256` in the Dockerfile.
- One Electron boot per file is the price of wrapping the released binary; an
  in-process multi-file batch (one boot, N exports) would be an upstream
  headless-export feature, not a packaging concern.
