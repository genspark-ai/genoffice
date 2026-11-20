# Sheets: spreadsheets

GenOffice Sheets is an Excel-compatible spreadsheet editor whose formula calculation and heavy data processing run inside a dedicated Rust engine process—isolated from the UI so even an extreme workbook can never bring the main application window down. It opens and saves genuine `.xlsx` workbooks natively, and opens `.csv` and `.tsv` files directly as editable tables.

## The interface

- **Ribbon**: Organized into eight main tabs along the top of the workspace (plus a contextual Chart Design tab when a chart is selected), each detailed below.
- **Formula bar**: Sits directly above the spreadsheet grid for viewing and editing the active cell's underlying formula or text, complete with function autocomplete and argument hints.
- **Sheet tabs** (bottom): Lists every worksheet in the active workbook so you can add new sheets, double-click to rename, drag to reorder, or right-click to delete, duplicate, color-code, or hide sheets.
- **Cell editing**: Double-click any cell or simply start typing while a cell is selected; pressing `Enter` commits your input and moves down one row, `Tab` commits and moves one column right, and `Escape` cancels the edit and restores the previous value, matching standard Excel habits.
- **Shortcuts**: Follows familiar desktop Excel key bindings across the board (`Ctrl+C / V / X`, `Ctrl+Z / Y`, `Ctrl+F`, `F2` to edit in place, `F4` to cycle absolute/relative references, and `⌘` equivalents on macOS).

## Ribbon tabs

- **Home**: Your everyday formatting and sheet controls—font family and size, cell fill colors, border styles, number formats (currency, percentage, thousands separator, increase/decrease decimal places), horizontal and vertical alignment, merge cells, row and column insert/delete and sizing, conditional formatting, Format as Table presets, cell style gallery, clipboard and format painter, plus quick sort and filter.
- **Insert**: Drop pictures, geometric shapes, text boxes, hyperlinks, cell comments, interactive checkboxes, page headers and footers, symbols, and math equations into your sheet; for data visualization, choose from **nine chart types** (column, bar, line, area, pie, scatter, radar, doughnut, and column+line combo) or click **Recommended Charts** to let Sheets rank and preview the best-fitting chart types for your current selection; you can also build a **PivotChart** directly from the active range, embed in-cell **Sparklines** (line, column, or win/loss), and add interactive **Slicer** and **Timeline** panels to filter a PivotTable.
  - Charts, PivotCharts, and sparklines are written as genuine OpenXML objects inside the saved `.xlsx` workbook and round-trip cleanly with desktop Excel.
  - **Slicer** and **Timeline** are live session controls: the filtering they apply to your PivotTable is saved right into the workbook so Excel displays the exact same filtered view, while the floating slicer widget itself lives in the session UI.
- **Page Layout**: Pick workbook theme colors and fonts, toggle whether gridlines and row/column headings appear when printing or exporting to PDF, and switch on page-break preview.
- **Formulas**: Insert functions through a guided wizard, use one-click AutoSum, define and manage named ranges (including batch-creating names from a selection's headers), draw precedent and dependent arrows to audit formulas, pin cells to the Watch Window, and control automatic or manual recalculation.
- **Data**: Sort and filter ranges (including multi-column custom sorts, criteria-range advanced filters, and one-click clear), split delimited text into columns, merge sheets from external workbooks, and refresh all PivotTables.
- **Review**: Step through cell comments (show/hide, jump to previous or next), translate selected cells with AI, and lock down structure using the **Protection** group—**Protect Sheet**, **Protect Workbook**, and **Allow Edit Ranges**.
  - Protection applied in GenOffice is written into the `.xlsx` file without a password so you cannot accidentally lock yourself out; once active, the ribbon button turns into **Unprotect …** and removes the lock in one click. Sheet protection applied by another program that _does_ carry a password hash cannot be stripped from here.
  - Clicking either protect button takes effect immediately without a modal dialog interrupting your flow—a status-bar note confirms that the protection state will be written into the workbook on your next save.
  - **Allow Edit Ranges** lets you designate specific input cells that remain unlocked and editable even while the rest of the worksheet is protected.
- **View**: Toggle gridlines, the **Formula Bar**, row/column headings, and the active-row/column cross-highlight; adjust canvas zoom; and switch between **Normal** and **Page Break Preview** views. Gridline and heading visibility are saved per worksheet inside the `.xlsx` file, whereas the active-cell cross-highlight is remembered as your personal local preference.
- **Chart Design**: Appears automatically whenever a chart is selected on the canvas, letting you switch chart types, swap color palettes and visual styles, or rebind the chart's source data range.

The **Data** tab, button by button (left to right in the screenshot below):

![The Data tab](img/sheets-data.png)

- **Pivot table**: Build an interactive PivotTable from the currently selected range and drag fields across rows, columns, values, and filters to aggregate your data.
- **Refresh**: Recalculate the active PivotTable immediately after you edit numbers in its underlying source range.
- **From text/CSV**: Import an external `.csv` or `.txt` file into a new worksheet, splitting columns cleanly by your chosen delimiter.
- **Merge workbooks**: Pull worksheets from one or more other `.xlsx` files on disk straight into your current workbook.
- **Refresh all**: Recalculate every PivotTable and external dataset across the entire workbook at once.
- **Sort** (dropdown): Sort the active column in ascending or descending order (keeping entire rows intact across neighboring columns), or open **Custom Sort** to stack multi-column sort rules.
- **Filter**: Attach `▼` filter dropdowns to the header row so you can check or uncheck specific values and filter rows by text or numeric conditions.
- **The three stacked buttons beside Filter**: **Clear** removes all active filter criteria and brings every hidden row back; **Reapply** runs the current filter rules again after data changes; and **Advanced** filters a table in place (or copies matches elsewhere) using a dedicated criteria range on the sheet.
- **Text to columns** (dropdown): Split a single column of pasted text into multiple columns either by a delimiter character (comma, tab, semicolon, space) or at fixed character widths.
- **Quick fill**: Type one example of how you want text extracted, combined, or reformatted in an adjacent column, then click Quick Fill (or press `⌃E` / `Ctrl+E`) to fill the rest of the column automatically by example.
- **Remove duplicates**: Scan the selected range and strip out duplicate rows based on whichever key columns you check.
- **Data validation** (dropdown): Enforce input rules on a range of cells—such as an in-cell dropdown picklist, a valid integer or decimal range, date bounds, or text-length limits—with optional input prompts and error alerts.
- **Consolidate**: Combine and summarize values from multiple separate ranges or worksheets into a single summary range matched by row and column labels.
- **What-if** (dropdown): Run **Goal Seek** to work backward from a desired formula result, automatically solving for the exact value an upstream input cell needs to hold to hit your target.
- **Group / Ungroup** (dropdown): Bundle contiguous rows or columns into collapsible outline groups with `+` / `-` gutter controls for folding detail rows in and out of view.
- **Subtotal**: Automatically insert category subtotal rows and a grand total row into a sorted dataset.

The **Formulas** tab, button by button:

![The Formulas tab](img/sheets-formulas.png)

- **Insert function** (`fx`): Open the searchable function picker with a step-by-step argument wizard that explains each parameter and previews the result as you type.
- **Excel 365 regex functions and PERCENTOF**: The formula bar supports Excel 365's regular-expression family—**`REGEXTEST`** (tests whether text matches a pattern), **`REGEXEXTRACT`** (extracts the first match, an array of all matches, or capture groups), and **`REGEXREPLACE`** (replaces matches with full support for `$0`, numbered `${n}`, and named `${name}` capture references, or replaces only the _n_th occurrence in context)—alongside **`PERCENTOF(data_subset, data_all)`** for computing a subset's share of a total; all four use standard Excel argument signatures, compile patterns with Unicode (`u`) support so `\p{...}` character classes work out of the box, and automatically refuse unsafe regex patterns that could cause catastrophic backtracking before they can freeze the sheet.
- **AutoSum** (dropdown): Insert a `SUM` formula in one click, or open the dropdown to insert `AVERAGE`, `COUNT`, `MAX`, or `MIN` over the adjacent range.
- **Recently used / Financial / Logical / Text / Date & time / Lookup & reference / Math & trig / More**: Browse the full function catalog organized by domain category and click any function to insert it with its argument prompts.
- **Name manager**: View, create, edit, and delete named ranges across the workbook from a single dialog.
- **Define name** (dropdown): Assign a readable name to the current selection, pick **Use in formula** to paste an existing name into a formula you are writing, or choose **Create from selection** to batch-generate named ranges automatically from the selection's top row or left column headers.
- **Trace precedents / Trace dependents**: Draw blue visual arrows directly across the grid showing which upstream cells feed into the active formula (**Trace precedents**) or which downstream formulas rely on the active cell (**Trace dependents**); click **Remove arrows** when you are done auditing.
- **Show formulas**: Toggle the entire sheet between displaying calculated results and displaying the raw `=...` formula strings inside every cell (with an indicator in the status bar so you always know when formula view is active).
- **Error checking**: Step through cells containing `#VALUE!`, `#NAME?`, `#REF!`, or `#DIV/0!` errors with an explanation of what went wrong in each formula.
- **Watch window**: Pin important summary cells from anywhere in the workbook into a floating watch list so you can monitor how their live values react while editing data on distant sheets.
- **Calculation options** (dropdown): Switch between **Automatic** recalculation (every dependent formula updates as soon as a cell changes) and **Manual** mode for massive workbooks; while in manual mode, use **Calculate now** (`F9`, whole workbook) or **Calculate sheet** (`⇧F9` / `Shift+F9`, active sheet only) whenever you are ready to recompute.

## Numbers and formatting

- **Number formats**: Format numbers as General, Number, Currency, Accounting, Percentage, Date, Time, Fraction, Scientific, or Plain Text without altering the underlying numeric precision stored in the cell.
- **Alignment, wrapping, borders, and fills**: Control horizontal and vertical text alignment, turn on text wrapping for long labels, merge header cells, and apply custom cell borders and background shading.
- **Row heights and column widths**: Drag row or column header boundaries to resize them manually, or double-click any boundary line to auto-fit the row height or column width to its longest entry.

## Working with data

**Sorting and filtering a table** (for example, sorting a table in descending order by one column):

1. Click **any single cell inside the column you want to sort by** (there is no need to highlight the entire column or table first—Sheets detects the contiguous data region automatically).
2. On the **Home** tab (or **Data** tab), click **Sort & Filter ▸ Descending**; entire rows reorder together as units so values across neighboring columns never get scrambled.
3. When you need multi-level rules (for instance, sorting by Department first and then by Revenue within each department, or sorting by cell color), follow the same menu path and choose **Custom Sort**.
4. To filter rows: select the header row and click **Sort & Filter ▸ Filter**—every column header gains a `▼` dropdown button where you can check or uncheck values to keep on screen, and clicking **Clear** restores every hidden row at once.

Three more everyday data essentials:

- **Sort and filter tools on the Data tab**: Combine quick filters with text-to-columns, duplicate removal, and criteria-range advanced filters covered in the Data tab breakdown above.
- **Frozen panes**: To keep your header rows or ID columns visible while scrolling through a large sheet, open the **View** tab, select the cell immediately below and to the right of the rows/columns you want locked, and click **Freeze Panes**.
- **Native `.csv` and `.tsv` editing**: Plain-text comma-separated (`.csv`) and tab-separated (`.tsv`) files open directly as live spreadsheet grids so you can sort, filter, and run formulas over them, and pressing Save writes your changes straight back out in the original delimited text format.

## Right-click menus

- **Right-clicking inside the grid**: Opens the spreadsheet editor's (Univer's) context menu with cut, copy, paste (including paste values or formats only), insert and delete rows or columns, hide/unhide, merge cells, freeze panes, and data actions.
- **Right-clicking the bottom status bar**: Lets you customize which live summary statistics appear in the status strip whenever multiple cells are selected (**Average**, **Count**, **Numerical Count**, **Min**, **Max**, **Sum**); your choices take effect immediately and persist across launches.
- **Right-clicking a worksheet tab at the bottom**: Opens the worksheet management menu for inserting, renaming, deleting, duplicating, color-coding, reordering, or hiding individual sheets.
- For the context menu on the top document tab strip (inline file rename, split beside, tear off into a new window), see [Tabs and window management](help://tabs-and-windows).

## AI assistant in Sheets

- **Natural-language range instructions**: Highlight any range of cells and tell the side AI panel what you need in plain language—whether that is cleaning up inconsistent date formats, generating a lookup or aggregation formula, filling sample data, or applying conditional formatting.
- **File and image attachments**: Click the `📎` button in the prompt box or drag external files straight onto the AI panel to attach reference spreadsheets, documents, or screenshots to your question; attached images display as inline thumbnails in the conversation.
- **Clickable cell citations**: When the AI's reply references a specific cell or range (such as `C12` or `Summary!B2:E20`), clicking that citation badge scrolls the spreadsheet grid straight to that location and highlights it.
- **One-click rollback**: Every AI edit flows through the spreadsheet's undoable editing pipeline, so you can roll back an AI turn from the panel or undo it with `⌘Z` / `Ctrl+Z` at any time.

## Saving and export

- **Full-fidelity `.xlsx` saves and PDF export**: Saving writes a standard `.xlsx` file with all formulas, number formats, charts, PivotTables, and named ranges intact; you can also use **Save As** to branch a copy or **Export to PDF** following your configured print pagination and page-setup rules.
- **Global AutoSave rules**: Just like the rest of the suite, AutoSave activates automatically once a workbook has been saved manually at least once.

## Process-isolated stability

- **Rust calculation sidecar isolation**: Because the Rust formula and data engine runs in its own separate operating-system process, even an unexpected engine fault on a massive or malformed file never crashes the GenOffice window—instead, the UI notifies you cleanly and attempts to recover the calculation session in place.
