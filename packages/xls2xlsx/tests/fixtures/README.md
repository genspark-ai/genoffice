# xls2xlsx test fixtures

`real-biff8.xls` is a genuine BIFF8 workbook, not built by the test byte
builder. It was produced by converting a small `.xlsx` (shared strings
including a rich two-run CJK string, formula cells with cached
number/string/boolean/error results, merged range `A6:C7`, and
`<workbookPr date1904="1"/>`) with LibreOffice:

```sh
soffice --headless -env:UserInstallation=file:///tmp/lo-profile \
  --convert-to xls --outdir <dir> source.xlsx
```

It pins the layouts the reader must accept from real producers: 4-byte SST
counts, `XLUnicodeRichExtendedString` with `cRun` before the characters,
`FormulaValue` special results, zero-based `Ref8` merge ranges and `Date1904`.
