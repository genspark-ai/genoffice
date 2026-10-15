# Test fixtures

## `optional_records.xls`

A real BIFF8 workbook written by Excel 2001, vendored from calamine's test
suite (`tests/optional_records.xls`, MIT licensed) where it pins the handling
of optional records (calamine issue #581). It carries three shared-formula
groups on its `Orig Sched` sheet, each stored the way Excel actually writes
one: a lone `tExp` token in every member's FORMULA record and the group's
tokens in a `SHRFMLA` record.

`src/tests.rs` converts it through `convert_to_xlsx` and checks the rebuilt
formulas (`N10/1000` for `M10` down the `M10:M36` group, `N11` for `AA11` down
the `AA11:AA74` and `AA75:AA92` groups). The expectations were read off a
LibreOffice conversion of the same file.
