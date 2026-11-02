# @genoffice/doc2docx

Legacy Word 97-2003 (`.doc`) import, staged so every layer can be reviewed
and tested on its own. The converter reads the OLE2/CFB container locally and
rebuilds a `.docx` through the shared engine; no upload or external service is
involved.

## Stage 1 — text foundation

Internal only. Establishes the binary-format foundation the later stages build
on:

- **OLE2/CFB container reader**: validated wrapper over `cfb` with the OLE2
  signature/header checks, 256 MiB input and 128 MiB per-stream caps, and the
  `WordDocument` stream required.
- **FIB parser**: FibBase/FibRgW97/FibRgLw97/FibRgFcLcb97, the nFib 0x00C1–0x0112
  gate, encrypted-document refusal, `0Table`/`1Table` selection, and the CLX
  plus formatting table ranges.
- **CLX piece table**: Prc/Pcdt walk, Windows-1252 compressed and UTF-16LE
  pieces, CP monotonicity and per-piece WordDocument range validation.
- **Body text and paragraph boundaries**: Word control marks mapped to readable
  text (cell mark → tab, line/page break → newline, field instruction dropped
  with the result kept, Windows-1252 punctuation decoded).

API: `openLegacyDoc(bytes)` returns the FIB, piece table, streams and text;
`readLegacyDocText(bytes)` returns just `{ text, paragraphs }`.

## Stage 2 — formatting model

Internal only. Reads the direct character and paragraph formatting of the body
story on top of the stage-1 piece table:

- **PlcBteChpx / PlcBtePapx** — the FC → FKP page tables in the table stream.
- **CHPX / PAPX FKPs** — run FC ranges, the short (`cb != 0`) and long
  (`cb == 0`) PapxInFkp forms, paragraph style index (istd) and default
  paragraphs (`bOffset == 0`).
- **Sprm parsing** — all spra operand-size classes, variable-length operands,
  unknown sprms skipped structurally, huge PAPX read through the Data stream.
- **Model** — `readLegacyDocModel(bytes)` returns
  `{ fib, text, paragraphs, warnings }`; each paragraph has
  `ParagraphFormatting` and `FormattedRun[]` with `CharacterFormatting`.
  Only explicit sprms are surfaced — style-sheet inheritance is stage 3b.

## Stage 3a — .docx writer (core conversion)

Internal only (no CLI/UI wiring yet): `convertLegacyDocToDocx(bytes)` returns
`{ docx, warnings }`. The writer builds the document on docx-engine's blank
template (`buildBlankDocx`) and emits one `<w:p>` per model paragraph through
`generateParagraphXml`:

- **runs**: bold/italic/strike, underline kinds → `w:u`, font size, superscript/
  subscript, highlight colours, the 16 Word colour indices → hex;
- **paragraphs**: alignment, left/right/first-line indents, spacing before/
  after, line spacing (multiple → `w:line` in 240ths, otherwise an exact
  rule), keep-lines / keep-next / page-break-before;
- **warnings** name every known gap honestly: unresolved styles (for
  documents whose stylesheet could not be resolved), flattened tables, and
  footnotes/endnotes/headers/text boxes that are not converted yet.

## Stage 3b — style resolution

Internal only. Parses the Word stylesheet (STSH) and folds it into the
formatting model, so style-based formatting is no longer dropped:

- **STSH parser** — STSHI + STD entries (Word 97 `cbSTDBaseInFile` 0x0A and
  Word 2000+ 0x12 with StdfPost2000 skipped), paragraph/character/table/
  numbering styles, `basedOn`/`next` chains, the XstzName and the PAPX/CHPX
  UPX sprms; every offset/length check raises `bad_stsh`.
- **Resolver** — paragraph rules from the paragraph style chain (base →
  derived) plus the direct PAPX sprms; character rules from the paragraph
  style CHPX chain, the default character style (istd 10), the run's character
  style (`sprmCIstd`) chain and then the direct CHPX sprms.
  ToggleOperand 0x80 keeps the inherited value and 0x81 inverts it; cycles,
  missing styles and out-of-range indices are reported as warnings while the
  remaining layers still apply.
- **Model** — `readResolvedLegacyDocModel(bytes)` returns the same shape as
  stage 2 with the merged formatting; `readLegacyDocModel` keeps the direct
  formatting model unchanged, and `convertLegacyDocToDocx` uses the resolved
  model whenever the document has a stylesheet.
- **Writer** — style-derived formatting is written as direct formatting, so
  it renders in Word/LibreOffice; the .docx styles themselves are not emitted
  yet (that is a later polish item).

## Roadmap

- **Stage 3c**: tables (TAP), images, headers/footers, section geometry
  (page size/margins), lists and hyperlinks, and emitting real .docx styles
  instead of resolved direct formatting; tighter paragraph-count parity.
- **Stage 3d**: docs import wiring (`.doc` → `.docx` in File → Open /
  drag-and-drop and the CLI), as core + UI PRs.

## Format coverage

Word 97-2003 (nFib 0x00C1–0x0112), main document story; paragraph/character
styles are resolved into direct formatting. Footnotes, headers/footers,
textboxes, tables, images and emitted .docx styles are later stages. Encrypted
files and Word 6/95 documents are rejected with actionable `LegacyDocError`
codes (`not_cfb`, `truncated`, `bad_fib`, `unsupported_version`,
`encrypted`, `missing_stream`, `bad_clx`, `bad_piece_table`, `bad_fkp`,
`bad_sprm`, `bad_stsh`, `limits_exceeded`, `corrupt`).
