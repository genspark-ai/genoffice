# @genoffice/xls2xlsx

Legacy Excel 97-2003 import: parse `.xls` (OLE2/CFB + BIFF8) and rebuild it as `.xlsx`. Staged like `@genoffice/doc2docx` — stage 1 is the internal values foundation (not user-visible yet).

## Relationship to the Sheets app

The Sheets desktop app already converts `.xls` through the xlsx-engine sidecar: `apps/sheets/src/main/sheets-main.ts`'s `prepareWorkbookForOpen` hands the file to the resolved conversion bridge, which carries pictures, merged cells and column widths into the session. This package is the **JavaScript/headless** counterpart: it exists so the CLI, MCP and other processes without the native sidecar can read a legacy workbook, and so the format logic (BIFF8 values, then dates/formats, then styles) is testable in plain Node. Stage 1 is not wired into the Sheets open flow, and wiring it is a separate decision — the sidecar remains the app's conversion path.

## Stage 1 carries

- cell values: text (shared strings, CJK and rich strings flattened, all read through the 4-byte SST counts and `XLUnicodeRichExtendedString` layout), numbers (RK/MulRk compressed forms), booleans, error literals, blanks
- merged ranges (`Ref8` is zero-based), column widths, sheet order, hidden sheets, the CJK-safe sheet names
- formulas as their **cached values** — including string, boolean and error results from `FormulaValue` (the BIFF8 RPN translation is a later stage)
- the 1904 date system: `Date1904` is read even though the record is always present, and a 1904 workbook writes `<workbookPr date1904="1"/>` so date serials do not shift by four years
- a validated OLE2/CFB container read through the shared `@genoffice/compound-file` opener (signature, 512-byte header, size caps, owned stream bytes — the same opener `@genoffice/doc2docx` uses) and an explicit error model (`LegacyXlsError`, stable codes — encrypted workbooks refuse up front via their FilePass record)
- tests run against a genuine LibreOffice-authored BIFF8 fixture as well as the byte builder, so the layouts are not just the reader's own assumptions

## Not carried yet

number formats / date rendering (serials stay raw), cell and font styles, charts, images, defined names.
