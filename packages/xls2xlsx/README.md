# @genoffice/xls2xlsx

Legacy Excel 97-2003 import: parse `.xls` (OLE2/CFB + BIFF8) and rebuild it as `.xlsx`. Staged like `@genoffice/doc2docx` — stage 1 is the internal values foundation (not user-visible yet); the CLI and the Sheets open flow wire it up in the following stages.

## Stage 1 carries

- cell values: text (shared strings, CJK included), numbers (RK/MulRk compressed forms), booleans, error literals, blanks
- merged ranges, column widths, sheet order, hidden sheets, the CJK-safe sheet names
- formulas as their **cached values** (the BIFF8 RPN translation is a later stage)
- a validated OLE2/CFB container read (signature, size caps) and an explicit error model (`LegacyXlsError`, stable codes — encrypted workbooks refuse up front via their FilePass record)

## Not carried yet

number formats / date rendering (serials stay raw), cell and font styles, charts, images, defined names.
