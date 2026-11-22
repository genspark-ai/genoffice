# XLSX round-trip accuracy — styled workbook class

Attribute = a formatting/structure fact the fixture baked in. intact = present
with its original value on the original cell. The genoffice run saves with two
edits (Data!B2 = 5, and bold OFF on the probe cell that also carries double
underline + fill + double border + a custom format); the LibreOffice run is its
own open + re-save of the untouched fixture.

| attribute                              | genoffice | soffice                                           |
| -------------------------------------- | --------- | ------------------------------------------------- |
| cell text: title                       | ✓         | ✓                                                 |
| title style: bold 14 + fill + thin box | ✓         | degraded: A1 style parts                          |
| numfmt: date on A2                     | ✓         | ✓                                                 |
| numfmt: currency on B2                 | ✓         | ✓                                                 |
| numfmt: percent on A3                  | ✓         | ✓                                                 |
| formula: cross-sheet survives          | ✓         | ✓                                                 |
| formula: SUM range survives            | ✓         | ✓                                                 |
| rich text: 2 runs, first bold          | ✓         | degraded: runs survive, formatting does not       |
| probe: double underline                | ✓         | ✓                                                 |
| probe: fill                            | ✓         | ✓                                                 |
| probe: double border                   | ✓         | ✓                                                 |
| probe: custom format                   | ✓         | ✓                                                 |
| probe: bold state                      | ✓         | degraded: bold lost without an edit asking for it |
| font: superscript + color              | ✓         | lost: A4 font                                     |
| font: italic                           | ✓         | ✓                                                 |
| merged range A7:B7                     | ✓         | ✓                                                 |
| defined name TOTALS                    | ✓         | ✓                                                 |
| table part: name + columns             | ✓         | ✓                                                 |
| sheet order: Invoice, Data             | ✓         | ✓                                                 |
| frozen top row                         | ✓         | ✓                                                 |
| column width survives                  | ✓         | lost: col width                                   |
| table header text                      | ✓         | ✓                                                 |

genoffice intact 22/22; LibreOffice intact 17/22.
