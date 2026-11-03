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
  subscript, highlight colours, hidden text → `w:vanish`, colours (sprmCCv RGB,
  else the 16 Word colour indices → hex), and `w:rFonts` ascii/hAnsi/eastAsia/cs
  resolved through the SttbfFfn font table (run → style chain → STSHI defaults);
- **paragraphs**: alignment, left/right/first-line indents (twips plus the
  character-unit `w:firstLineChars` twin), spacing before/after, line spacing
  (multiple → `w:line` in 240ths, negative → exact, positive → atLeast),
  keep-lines / keep-next / page-break-before, custom tab stops (`w:tabs` from
  sprmPChgTabs/sprmPChgTabsPapx) and paragraph borders (`w:pBdr` from the
  Brc80/Brc sprms);
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

## Stage 3c-1 — section geometry and headers/footers

Internal only. Reads the PlcfSed section table and its SEPX section sprms, and
the PlcfHdd header/footer stories (FIB pair 11), and writes them into the
generated package:

- **parseSections** — page size, margins (left/right/top/bottom), header and
  footer distances, gutter, portrait/landscape and the title-page flag.
  Section sprms are matched by their low 9-bit operation (producers vary the
  spra bits, e.g. LibreOffice writes sprmSDyaTop as 0x9023), each section
  starts from the blank template's geometry, and multi-column sections are
  reported as warnings. Malformed tables/SEPX bytes raise \`bad_section\`.
- **parseHeaderStories** — the six separator stories plus six stories per
  section (even header, odd header, even footer, odd footer, first header,
  first footer); story CPs are offsets from the header document at
  \`ccpText + ccpFtn\`, the guard paragraph mark is stripped, and empty stories
  stay empty. Story paragraphs carry their resolved run/paragraph formatting
  (a formatting read over the whole text stream) and are written with the
  body paragraph writer. Malformed tables raise \`bad_hdd\`; a document that
  declares header text without a header table gets a warning instead of failing.
- **Writer** — replaces the template's trailing \`sectPr\` with the last
  section's geometry (header/footer references, \`pgSz\` with
  \`w:orient="landscape"\`, \`pgMar\`, \`titlePg\`, \`docGrid\`), always
  writes the Word 97-2003 compatibility block into settings.xml (what Word
  itself writes when it saves a .doc as .docx) plus the document languages
  and Word 97 paragraph defaults into styles.xml, and adds
  \`word/headerN.xml\` / \`word/footerN.xml\` parts with the matching
  [Content_Types].xml overrides and document.xml.rels relationships. Even/odd
  headers need settings.xml (\`w:evenAndOddHeaders\`), which is added only when
  an even story exists; the odd stories are the default variant and the first
  stories need \`w:titlePg\`. Header/footer character formatting is plain text
  for now (tabs and line breaks are kept).
- **Limitations** — multi-section documents are flattened to their last
  section with an explicit warning; page-number format and footnote/endnote
  separator stories are not written; header/footer runs carry no bold/size
  etc. yet.

## Stage 3c-2 — tables

Internal only. Reconstructs Word 97-2003 tables from the body text marks (a
cell ends with a cell mark, a row with an extra empty cell mark) and the TAP
sprms of each row-end PAPX, and writes real `w:tbl` tables:

- **sprm sizing** — sprmTDefTable / sprmPChgTabs (spra 6) carry a u16 length
  that counts one of its own bytes, so the payload is the declared value minus
  one (POI-compatible); the generic spra-6 sprms keep their u8 length.
- **readLegacyDocBody** — splits the body at paragraph and cell marks into CP
  units, maps each unit to its PAPX entry, parses sprmTDefTable (rgdxaCenter
  cell boundaries plus 20-byte TC descriptors: merge flags, width, four
  BorderCodes) and the other TAP sprms (alignment, cantSplit, table header,
  row height, table borders, cell padding). Rows share one grid derived from
  the union of their cell boundaries; fMerged continuation cells fold into the
  first cell gridSpan (LibreOffice-flattened merged rows are recovered from
  their wider boundary span), vertical merges map to vMerge restart/continue,
  and multi-paragraph cells keep all their paragraphs. Cell character
  formatting is resolved from the CHPX runs by CP; cell paragraph formatting
  (indents, spacing, alignment, list membership) is resolved like body
  paragraphs. A row-end PAPX moved to the Data stream
  (sprmPHugePapx) is expanded before the TAP is read. Malformed TAP/TDefTable
  data degrades with a warning: the row (or the whole table when no grid can
  be built) is imported as plain paragraphs, a truncated sprm ends the TAP walk.
- **Writer** — emits `<w:tbl>` with tblW auto, tblBorders (a single 4-eighth
  point line set when any border is visible), tblLayout fixed, a tblGrid with
  gridCol widths, trPr (cantSplit, trHeight, tblHeader), tcPr (tcW dxa,
  gridSpan, vMerge, tcBorders) and the cell paragraphs with their resolved
  runs, list membership and pictures; tables interleave with body paragraphs in CP order. Table cell
  padding, cell shading and table styles (e.g. "Table Grid") are parsed but
  not written (borders come from the TC descriptors); nested tables are
  flattened with a warning.
- **Warnings** — malformed rows and nested tables are reported once per
  document.

## Stage 3c-3 — pictures (inline and floating)

Internal only. Reads the inline pictures of the body story from the CFB Data
stream and the floating OfficeArt pictures of the body and header stories,
and writes them as real .docx media parts with drawings:

- **parseInlineImages** — walks the body CPs, ignores picture characters inside
  field instructions (EMBED results are pictures), maps each remaining picture
  CP to its CHPX run and requires sprmCFSpec plus a signed sprmCPicLocation.
  The PICF at that Data offset is validated (lcb, cbHeader, bounds), the
  display size comes from dxaGoal * mx / 1000 (mx/my are unsigned 16-bit
  scaling factors) with xExt/yExt and plain dxaGoal as fallbacks, and the
  picture bytes come from the OfficeArtBlip record inside the PICF range (PNG,
  JPEG, DIB, WMF, EMF; see below), falling back to a raw PNG/JPEG signature
  scan. Corrupt or truncated picture data degrades per picture (warning, the
  picture is skipped); only a structurally invalid PICF header raises
  `bad_picture`.
- **parseFloatingPictures** (`drawings.ts`) — PlcSpaMom (FIB pair 40) and
  PlcSpaHdr (pair 41) give the FSPA anchors: the CP of the \x08 anchor
  character, the shape id, the rectangle in twips and the bx/by (margin / page
  / text) and wr (wrapping) flags. DggInfo (pair 50) is walked as OfficeArt
  records: the blip store (FBSE entries pointing at a blip record in the
  WordDocument stream, or embedding it) and one drawing container per story
  whose shape containers carry the pib property. Group shapes contribute their
  picture members, placed through the group's coordinate bounds and child
  anchors. Blip records decode PNG/JPEG/TIFF (UID(s) + tag + bytes), DIB (a
  BITMAPFILEHEADER is prepended, written as .bmp) and WMF/EMF (34-byte
  metafile header, deflate-compressed data inflated); PICT is skipped.
  Non-picture shapes (text boxes, autoshapes, lines) are skipped with one
  warning; a missing or truncated blip skips that picture.
- **Body model** — `readLegacyDocBody` emits `{ type: "image" }` blocks for a
  picture character that is a paragraph of its own; pictures that share a
  paragraph with text or sit in a table cell travel on the paragraph block
  (`pictures.inline`, written at the paragraph start or end), and floating
  pictures attach to the paragraph that holds their anchor (`pictures.floating`
  on body blocks and `paragraphPictures` on table cells). Header/footer
  stories carry theirs in `pictures` by paragraph index; a header story that
  holds nothing but an anchor is kept.
- **Writer** — one `word/media/imageN.<ext>` per distinct picture (png, jpeg,
  wmf, emf, bmp, tiff), `<Default Extension=...>` entries, image relationships
  in document.xml.rels and in `word/_rels/headerN.xml.rels` / footer rels for
  header pictures. Inline pictures are `<wp:inline>` runs; floating pictures
  are `<wp:anchor>` runs with positionH/positionV offsets relative to
  page/margin/column (paragraph), the FSPA extent, wrapSquare /
  wrapTopAndBottom / wrapNone from wr, and behindDoc from fBelowText. Header
  pictures are written without wrapping, as Word lays them out. docPr ids are
  unique across the document and its header parts.
- **Limitations** — pictures inside text boxes, footnotes and endnotes are not
  handled yet, tight/through wrapping is written as square, grouped
  non-picture members are dropped, and picture metadata (alt text, borders,
  hyperlinks, cropping) is not carried over.

## Stage 3c-4 — lists/numbering and HYPERLINK fields

Internal only. Reads the Word 97-2003 list tables and the HYPERLINK and
page-number fields and writes real numbering definitions, hyperlink runs and
live PAGE / NUMPAGES / SECTIONPAGES fields (fldChar begin, instrText,
separate, the cached result, end) in the body, table cells and header/footer
stories:

- **parseListTables** — PlfLst (FIB pair 73): cLst 28-byte LSTF records and the
  LVL array that follows them (one level for fSimpleList lists, nine
  otherwise). Each LVL gives start value, nfc numbering format, level text,
  indents from grpprlPapx and the character properties; placeholders in the
  Xst become %1..%9 and symbol-font bullet glyphs are mapped to text bullets.
  PlfLfo (pair 74) maps 1-based ilfo values to lsids; its LFOData LFOLVLs
  become w:lvlOverride entries (startOverride and/or a replacement w:lvl) and
  a paragraph's bullet/number kind follows the overriding level. Malformed tables degrade with a
  warning: a broken level grpprl keeps the level without its indents, a
  truncated level drops the remaining definitions, a bad table range drops
  that table.
- **Formatting model** — sprmPIlfo/sprmPIlvl already flow through the direct
  PAPX and style resolution, so resolved paragraphs carry `listId` and
  `listLevel`; the body builder turns them into `{ kind, numId: 100 + ilfo,
ilvl }` list items (a missing list definition is reported and the paragraph
  stays plain).
- **parseFields** — walks the body text field ranges (\x13 instruction
  \x14 result \x15) with a stack for nested fields, parses HYPERLINK
  instructions (quoted/bare URL, \l anchors, \o tooltips) into result CP
  ranges, and reports unsupported field types or malformed/unclosed fields as
  warnings while keeping the result as plain text.
- **Writer** — appends one `w:abstractNum` per used lsid (ids from 10, levels
  with w:start/w:numFmt/w:lvlText/w:lvlJc and the level indents) and one
  `w:num` per used ilfo (ids 100 + ilfo) to the blank template numbering.xml
  (CT_Numbering keeps abstractNum* before num*); numbered paragraphs get
  `w:numPr` with their direct formatting. Hyperlink result runs are wrapped in
  `w:hyperlink`: external URLs get a relationship with
  `TargetMode="External"`, bookmark links use `w:anchor`; the field
  instruction text never reaches the output, and documents without lists or
  fields keep the previous writer output.
- **Limitations** — unsupported field types are warned and left as plain
  text; EMF/WMF pictures, floating images, footnotes and emitted real .docx
  styles are still later stages.

## Roadmap

- **Stage 3c-5**: EMF/WMF conversion, floating images and pictures in
  headers/cells, footnotes/endnotes, emitted real .docx styles instead of
  resolved direct formatting, richer header/footer formatting, real
  multi-section output and table cell padding/shading/styles; tighter
  paragraph-count parity.
- **Stage 3d**: docs import wiring (`.doc` files through File → Open /
  drag-and-drop and the CLI), as core + UI PRs.

## Format coverage

Word 97-2003 (nFib 0x00C1–0x0112), main document story with section geometry,
header/footer stories, real tables, inline PNG/JPEG pictures, lists/numbering
and HYPERLINK fields; paragraph/character styles are resolved into direct
formatting. Footnotes, textboxes, EMF/WMF pictures, floating images and
emitted .docx styles are later stages, and multi-section documents are
flattened to one section. Encrypted files and Word 6/95 documents are rejected
with actionable `LegacyDocError` codes (`not_cfb`, `truncated`, `bad_fib`,
`unsupported_version`, `encrypted`, `missing_stream`, `bad_clx`,
`bad_piece_table`, `bad_fkp`, `bad_sprm`, `bad_stsh`, `bad_section`,
`bad_hdd`, `bad_table`, `bad_picture`, `bad_list`, `limits_exceeded`,
`corrupt`).

## Comments

Word 97-2003 annotations become `word/comments.xml`. `src/comments.ts` reads
PlcfandRef (anchor CP + ATRDPre10: initials, author index into SttbfRMark,
bookmark tag), PlcfandTxt (comment text ranges of the comment document, which
follows the header document in the text stream), the annotation bookmarks
(SttbfAtnBkmk tag → PlcfAtnBkf entry → PlcfAtnBkl CP) and, for Word 2002+
files, the ATRDPost10 `dttm` dates. Comment text goes through the same story
formatter as headers, so it keeps its character formatting.

- Ranged comments emit `w:commentRangeStart` / `w:commentRangeEnd` around the
  bookmarked text and a `CommentReference` run after the range end; point
  comments (tag -1 or an empty bookmark) emit only the reference run where the
  `\x05` annotation character stood. The anchors are placed while the runs are
  resolved (`commentMarks` in `resolveRunsForRange`), so they land correctly
  inside table cells and across runs.
- Missing comment tables with `ccpAtn > 0` are a warning; invalid tables raise
  `bad_atn`. Comments anchored outside the main text (headers, footnotes) are
  skipped with a warning for now.
