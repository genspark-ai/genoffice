import { z } from 'zod'
import { columnIndex, columnLabel, formatAddress, parseRange, rangeCellCount } from './cell-address'
import { computeSortChanges } from './sort-range'

const cellAddressSchema = z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/)
const cellRangeSchema = z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$/)
const columnLabelSchema = z.string().regex(/^[A-Z]{1,3}$/)
const sheetNameSchema = z
  .string()
  .trim()
  .min(1)
  .max(31)
  .refine((name) => !/[:\\/?*[\]]/.test(name))
const hexColorSchema = z.string().regex(/^#[0-9A-Fa-f]{6}$/)
const cellScalarSchema = z.union([z.string(), z.number().finite(), z.boolean(), z.null()])

const setCellSchema = z.object({
  op: z.literal('set_cell'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
  value: cellScalarSchema,
  expectedValue: cellScalarSchema.optional(),
})

const setFormulaSchema = z.object({
  op: z.literal('set_formula'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
  formula: z.string().startsWith('=').max(8192),
  expectedFormula: z.string().optional(),
})

const clearCellSchema = z.object({
  op: z.literal('clear_cell'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
})

// String values starting with "=" are treated as formulas, matching what
// typing the same text into the cell editor would do.
// `start` (top-left target cell) is canonical; `range` is also accepted
// because every other range-shaped op uses that field name and models keep
// reaching for it — when given, its size must match the values grid, which
// doubles as a misaligned-write check (validated in expandToPrimitiveOps).
const setRangeSchema = z.object({
  op: z.literal('set_range'),
  sheetId: z.string().min(1),
  start: cellAddressSchema.optional(),
  range: cellRangeSchema.optional(),
  values: z.array(z.array(cellScalarSchema).min(1).max(100)).min(1).max(500),
})

const clearRangeSchema = z.object({
  op: z.literal('clear_range'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

const insertRowsSchema = z.object({
  op: z.literal('insert_rows'),
  sheetId: z.string().min(1),
  /** 1-based; new rows are inserted before this row */
  row: z.number().int().min(1).max(9999999),
  count: z.number().int().min(1).max(500),
})

const deleteRowsSchema = z.object({
  op: z.literal('delete_rows'),
  sheetId: z.string().min(1),
  /** 1-based first row to delete */
  row: z.number().int().min(1).max(9999999),
  count: z.number().int().min(1).max(500),
})

const insertColsSchema = z.object({
  op: z.literal('insert_cols'),
  sheetId: z.string().min(1),
  /** new columns are inserted before this column */
  column: columnLabelSchema,
  count: z.number().int().min(1).max(100),
})

const deleteColsSchema = z.object({
  op: z.literal('delete_cols'),
  sheetId: z.string().min(1),
  column: columnLabelSchema,
  count: z.number().int().min(1).max(100),
})

const addSheetSchema = z.object({
  op: z.literal('add_sheet'),
  name: sheetNameSchema,
})

const deleteSheetSchema = z.object({
  op: z.literal('delete_sheet'),
  sheetId: z.string().min(1),
})

// Edits an EXISTING chart (charts are listed in get_workbook_context):
// file charts by chart part path, session-added and demo charts by their
// visual id. At least one property required — checked at expansion because
// discriminated-union members cannot carry refinements.
const editChartSchema = z.object({
  op: z.literal('edit_chart'),
  chartPath: z
    .string()
    .regex(/^(xl\/charts\/[A-Za-z0-9._-]+\.xml|(added|demo)-chart-[A-Za-z0-9._-]+)$/),
  title: z.string().max(255).optional(),
  chartType: z.enum(['column', 'bar', 'line', 'area', 'pie', 'doughnut']).optional(),
  /** series index ("0"-based, as string) → #RRGGBB */
  seriesColors: z.record(z.string().regex(/^[0-9]{1,3}$/), hexColorSchema).optional(),
  /** 'none' hides the legend; a side moves it (creating it if absent) */
  legend: z.enum(['none', 'right', 'bottom', 'top', 'left']).optional(),
  /** plot-level data labels; 'none' removes them */
  dataLabels: z.enum(['none', 'value', 'percent', 'category-percent']).optional(),
  /** bar/line/area stacking; 'clustered' restores side-by-side */
  grouping: z.enum(['clustered', 'stacked', 'percentStacked']).optional(),
  /** axis-based charts only; null removes that axis title */
  axisTitles: z
    .object({
      category: z.string().max(255).nullable().optional(),
      value: z.string().max(255).nullable().optional(),
    })
    .optional(),
  /** rename a series and/or repoint its data at new single-row/column ranges */
  seriesData: z
    .array(
      z.object({
        index: z.number().int().min(0).max(255),
        name: z.string().max(255).optional(),
        valuesRange: cellRangeSchema.optional(),
        categoriesRange: cellRangeSchema.optional(),
        /** sheet the ranges live on; default: the chart's own sheet */
        sheetId: z.string().min(1).optional(),
      }),
    )
    .min(1)
    .max(24)
    .optional(),
})

// Creates a NEW chart from a data range (imported workbooks only). The range
// may include a header row and a leading category column — both are detected
// from the data, matching the ribbon's Insert Chart.
const addChartSchema = z.object({
  op: z.literal('add_chart'),
  sheetId: z.string().min(1),
  chartType: z.enum([
    'column',
    'bar',
    'line',
    'area',
    'pie',
    'doughnut',
    'scatter',
    'radar',
    'combo',
  ]),
  dataRange: cellRangeSchema,
  title: z.string().max(255).optional(),
  /** top-left cell of the chart frame; default: two columns right of the data */
  anchorCell: cellAddressSchema.optional(),
})

const addShapeSchema = z.object({
  op: z.literal('add_shape'),
  sheetId: z.string().min(1),
  shapeType: z.enum([
    'rect',
    'roundRect',
    'ellipse',
    'triangle',
    'diamond',
    'rightArrow',
    'leftArrow',
    'pentagon',
    'hexagon',
    'textbox',
  ]),
  anchorCell: cellAddressSchema,
  fillColor: hexColorSchema.optional(),
  text: z.string().max(1000).optional(),
})

// Edits a shape/text box ADDED THIS SESSION (ids listed by
// read_sheet_features); file-original drawings stay read-only. At least one
// property required — checked at expansion.
const editShapeSchema = z.object({
  op: z.literal('edit_shape'),
  visualId: z.string().min(1).max(128),
  text: z.string().max(1000).optional(),
  fillColor: hexColorSchema.optional(),
  /** moves the shape's top-left corner, keeping its size */
  anchorCell: cellAddressSchema.optional(),
})

// Inserts a local image file (PNG/JPEG/GIF, ≤20MB) the user pointed at.
const addImageSchema = z.object({
  op: z.literal('add_image'),
  sheetId: z.string().min(1),
  /** absolute path on this machine (~/ allowed) */
  path: z.string().min(1).max(1024),
  anchorCell: cellAddressSchema,
})

// Creates a real Excel table (ListObject) over a data range (imported
// workbooks only). The range's first row must hold the column headers —
// non-empty and unique; they become the table's column names in the file.
const addTableSchema = z.object({
  op: z.literal('add_table'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  /** Excel table name (defined-name rules: letter/underscore start, no spaces); default Table1, Table2, … */
  name: z
    .string()
    .regex(/^[\p{L}_\\][\p{L}\p{N}_.\\]{0,254}$/u)
    .optional(),
  /** built-in table style, e.g. "TableStyleMedium2" (the default) */
  style: z
    .string()
    .regex(/^TableStyle(?:Light|Medium|Dark)[1-9][0-9]?$/)
    .optional(),
  /** banded (striped) rows — Excel's default is on */
  bandedRows: z.boolean().optional(),
})

// Inserts blank rows into a session-created table. The table must have been
// created with add_table in the current session (fail-closed: file tables not
// supported yet — save and reopen first). Rows are inserted into the data area
// (not the header row). row is 1-based within the data area; omit to append.
const addTableRowSchema = z.object({
  op: z.literal('add_table_row'),
  sheetId: z.string().min(1),
  /** Table name (must match an add_table op from this session) */
  tableName: z.string().min(1),
  /** 1-based position within the data area to insert before; omit to append */
  row: z.number().int().min(1).optional(),
  count: z.number().int().min(1).max(1000).default(1),
})

// Inserts blank columns into a session-created table. Column is 1-based
// within the table area; omit to append. columnName must be non-blank and
// unique within the table.
const addTableColumnSchema = z.object({
  op: z.literal('add_table_column'),
  sheetId: z.string().min(1),
  tableName: z.string().min(1),
  /** 1-based column position within the table to insert before; omit to append */
  column: z.number().int().min(1).optional(),
  columnName: z.string().min(1).max(255),
  count: z.number().int().min(1).max(100).default(1),
})

// Deletes rows from the data area of a session-created table.
// row is 1-based within the data area; table must retain ≥1 data row.
const deleteTableRowSchema = z.object({
  op: z.literal('delete_table_row'),
  sheetId: z.string().min(1),
  tableName: z.string().min(1),
  /** 1-based first row to delete within the data area */
  row: z.number().int().min(1),
  count: z.number().int().min(1).max(1000).default(1),
})

// Deletes columns from a session-created table; table must retain ≥1 column.
const deleteTableColumnSchema = z.object({
  op: z.literal('delete_table_column'),
  sheetId: z.string().min(1),
  tableName: z.string().min(1),
  /** 1-based column to delete within the table */
  column: z.number().int().min(1),
  count: z.number().int().min(1).max(100).default(1),
})

// Creates a real PivotTable (imported workbooks only): the apply bakes the
// aggregated grid into the target cells, and the save writes native pivot
// parts with refreshOnLoad so Excel turns it into a live pivot on open.
// Constraints: 1–8 row dimension levels; 0–8 column dimension levels; with
// column field(s) exactly one value entry; up to 4 page (report filter) fields.
const addPivotSchema = z.object({
  op: z.literal('add_pivot'),
  sheetId: z.string().min(1),
  /** source data range, first row = field headers (non-blank, unique) */
  sourceRange: cellRangeSchema,
  /** top-left cell of the pivot output; must not overlap the source */
  targetCell: cellAddressSchema,
  /** sheet for the output; default: the source sheet */
  targetSheetId: z.string().min(1).optional(),
  /**
   * Row dimension field(s) — 1 to 8 headers from the source, outer levels first.
   * For a single level you may pass a bare string or a one-element array.
   * With multiple fields the pivot groups hierarchically and subtotal rows
   * are automatically inserted for every non-leaf level.
   */
  rowFields: z.union([
    z.string().min(1).max(255),
    z.array(z.string().min(1).max(255)).min(1).max(8),
  ]),
  /**
   * Column dimension field(s) to spread across columns — 1 to 8 headers,
   * Outer levels first. Single level may be a bare string; multiple levels take
   * an array, expanding members as a cartesian product in column-field order,
   * with a subtotal column appended after each non-leaf member.
   */
  columnField: z
    .union([z.string().min(1).max(255), z.array(z.string().min(1).max(255)).min(1).max(8)])
    .optional(),
  /**
   * Report filter fields (pageFields) — up to 4 source headers that are
   * placed above the pivot as filter drop-downs in the saved Excel file.
   * The baked grid omits the filter row; Excel/LibreOffice shows them on open.
   */
  pageFields: z.array(z.string().min(1).max(255)).max(4).optional(),
  values: z
    .array(
      z.object({
        field: z.string().min(1).max(255),
        agg: z.enum(['sum', 'count', 'average', 'max', 'min']),
        /** Optional Excel number format string, e.g. "#,##0.00" or "0%" */
        numFmt: z.string().min(1).max(255).optional(),
        /**
         * "Show values as" mode: percentOfTotal = percent of grand total /
         * percentOfRow = percent of row total / percentOfCol = percent of column
         * total; defaults to the plain aggregate value.
         */
        showDataAs: z.enum(['percentOfTotal', 'percentOfRow', 'percentOfCol']).optional(),
        /**
         * Calculated field: when formula is present, field is the new data field's
         * name (must not clash with source headers); the formula does basic
         * arithmetic over other source field names (e.g. "Revenue-Cost"; wrap names
         * containing spaces in single quotes); each group sums the referenced fields
         * first, then evaluates the formula, and agg must be sum.
         */
        formula: z.string().min(1).max(1_024).optional(),
      }),
    )
    .min(1)
    .max(8),
  /**
   * Dimension-field grouping: date groups by year/quarter/month (quarter/month
   * merge across years); range groups into fixed-step numeric intervals (bucket
   * boundaries rangeStart + k*rangeStep, default start 0). field must be a
   * dimension field from rowFields/columnField.
   */
  groupings: z
    .array(
      z.discriminatedUnion('kind', [
        z.object({
          kind: z.literal('date'),
          field: z.string().min(1).max(255),
          dateUnit: z.enum(['year', 'quarter', 'month']),
        }),
        z.object({
          kind: z.literal('range'),
          field: z.string().min(1).max(255),
          rangeStep: z.number().positive().finite(),
          rangeStart: z.number().finite().optional(),
        }),
      ]),
    )
    .max(8)
    .optional(),
  /**
   * Value/label filters: label filters row/column field members by label
   * (equals/contains/begins-with, case-insensitive); value filters the field's
   * members by one values entry's aggregate (top: largest count / greaterThan:
   * > from / between: from ≤ x ≤ to). field must be a dimension field from
   * rowFields/columnField, at most one filter per field.
   */
  filters: z
    .array(
      z.discriminatedUnion('kind', [
        z.object({
          kind: z.literal('label'),
          field: z.string().min(1).max(255),
          op: z.enum(['equal', 'contains', 'beginsWith']),
          value: z.string().max(255),
        }),
        z.object({
          kind: z.literal('value'),
          field: z.string().min(1).max(255),
          /** The data field it measures (values array index, defaults to 0) */
          valueIndex: z.number().int().nonnegative().max(7).default(0),
          op: z.enum(['top', 'greaterThan', 'between']),
          count: z.number().int().min(1).max(10_000).optional(),
          from: z.number().finite().optional(),
          to: z.number().finite().optional(),
        }),
      ]),
    )
    .max(8)
    .optional(),
  /** pivot table name; default Pivot1, Pivot2, … */
  name: z
    .string()
    .regex(/^[\p{L}_\\][\p{L}\p{N}_.\\ ]{0,254}$/u)
    .optional(),
})

const setRowsHiddenSchema = z.object({
  op: z.literal('set_rows_hidden'),
  sheetId: z.string().min(1),
  /** 1-based first row */
  row: z.number().int().min(1).max(9999999),
  count: z.number().int().min(1).max(10000).default(1),
  hidden: z.boolean(),
})

const setColsHiddenSchema = z.object({
  op: z.literal('set_cols_hidden'),
  sheetId: z.string().min(1),
  column: columnLabelSchema,
  count: z.number().int().min(1).max(1000).default(1),
  hidden: z.boolean(),
})

// target: full URL (https://…) or internal reference like "Sheet1!A1";
// null removes the link.
const setHyperlinkSchema = z.object({
  op: z.literal('set_hyperlink'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
  target: z.string().min(1).max(2048).nullable(),
})

const protectSheetSchema = z.object({
  op: z.literal('protect_sheet'),
  sheetId: z.string().min(1),
  protected: z.boolean(),
})

// Creates (or replaces) the sheet's auto-filter over the range. Filter
// criteria are then chosen by the user through the column dropdowns.
const setFilterSchema = z.object({
  op: z.literal('set_filter'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

const clearFilterSchema = z.object({
  op: z.literal('clear_filter'),
  sheetId: z.string().min(1),
})

// Checkbox-style column criteria on an existing auto-filter: rows whose cell
// text is in `values` stay visible. values: null clears the column's criteria.
const setFilterCriteriaSchema = z.object({
  op: z.literal('set_filter_criteria'),
  sheetId: z.string().min(1),
  column: columnLabelSchema,
  values: z.array(z.string().max(255)).min(1).max(1000).nullable(),
})

const cfFormatSchema = z
  .object({
    fillColor: hexColorSchema.optional(),
    fontColor: hexColorSchema.optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
  })
  .refine((format) => Object.values(format).some((value) => value !== undefined), {
    message: 'The rule format must set at least one property.',
  })

export const cfRuleSchema = z.discriminatedUnion('kind', [
  z.object({
    kind: z.literal('number'),
    operator: z.enum([
      'greaterThan',
      'greaterThanOrEqual',
      'lessThan',
      'lessThanOrEqual',
      'equal',
      'notEqual',
      'between',
      'notBetween',
    ]),
    value: z.number().finite(),
    /** required for between / notBetween */
    value2: z.number().finite().optional(),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('text'),
    operator: z.enum(['contains', 'notContains', 'beginsWith', 'endsWith']),
    text: z.string().min(1).max(255),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('blank'),
    /** true highlights empty cells, false highlights non-empty ones */
    blank: z.boolean(),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('duplicate'),
    /** true highlights unique values instead of duplicates */
    unique: z.boolean().optional(),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('top10'),
    rank: z.number().int().min(1).max(1000),
    percent: z.boolean().optional(),
    bottom: z.boolean().optional(),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('formula'),
    formula: z.string().startsWith('=').max(8192),
    format: cfFormatSchema,
  }),
  z.object({
    kind: z.literal('colorScale'),
    minColor: hexColorSchema,
    midColor: hexColorSchema.optional(),
    maxColor: hexColorSchema,
  }),
  z.object({
    kind: z.literal('dataBar'),
    color: hexColorSchema.optional(),
  }),
])

const addConditionalFormatSchema = z.object({
  op: z.literal('add_conditional_format'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  rule: cfRuleSchema,
})

const clearConditionalFormatsSchema = z.object({
  op: z.literal('clear_conditional_formats'),
  sheetId: z.string().min(1),
})

export const dvRuleSchema = z.discriminatedUnion('kind', [
  z.object({
    kind: z.literal('list'),
    values: z.array(z.string().min(1).max(255)).min(1).max(100),
  }),
  z.object({
    kind: z.literal('listRef'),
    /** range whose cell values form the dropdown options */
    range: cellRangeSchema,
  }),
  z.object({
    kind: z.literal('numberBetween'),
    min: z.number().finite(),
    max: z.number().finite(),
  }),
  z.object({
    kind: z.literal('dateBetween'),
    /** ISO dates, e.g. "2026-01-31" */
    start: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
    end: z.string().regex(/^\d{4}-\d{2}-\d{2}$/),
  }),
  z.object({ kind: z.literal('checkbox') }),
  z.object({
    kind: z.literal('formula'),
    formula: z.string().startsWith('=').max(8192),
  }),
])

// validation: null clears any rule on the range.
const setDataValidationSchema = z.object({
  op: z.literal('set_data_validation'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  validation: dvRuleSchema.nullable(),
})

// Print/page-layout settings; saved into the file, mirroring the Page Layout
// ribbon. At least one setting required (checked at expansion). scale and
// fitToWidth/fitToHeight are mutually exclusive.
const setPageSetupSchema = z.object({
  op: z.literal('set_page_setup'),
  sheetId: z.string().min(1),
  orientation: z.enum(['portrait', 'landscape']).optional(),
  /** OOXML paper-size code: 1 Letter, 8 A3, 9 A4, 11 A5 … */
  paperSize: z.number().int().min(1).max(118).optional(),
  /** print scale percent */
  scale: z.number().int().min(10).max(400).optional(),
  /** pages wide (0 = automatic) */
  fitToWidth: z.number().int().min(0).max(1000).optional(),
  /** pages tall (0 = automatic) */
  fitToHeight: z.number().int().min(0).max(1000).optional(),
  margins: z.enum(['normal', 'wide', 'narrow']).optional(),
  printGridlines: z.boolean().optional(),
  printHeadings: z.boolean().optional(),
  /** A1 range to print; null clears the print area */
  printArea: cellRangeSchema.nullable().optional(),
})

// Cell note (legacy comment): text null removes the note.
const setNoteSchema = z.object({
  op: z.literal('set_note'),
  sheetId: z.string().min(1),
  address: cellAddressSchema,
  text: z.string().min(1).max(32_000).nullable(),
})

// Recomputes every pivot table on the sheet from its current source data
// (layout unchanged; requires the fully-loaded mode).
const refreshPivotSchema = z.object({
  op: z.literal('refresh_pivot'),
  sheetId: z.string().min(1),
})

// Freeze panes: rows/columns counted from the top-left; 0/0 unfreezes.
const setFreezeSchema = z.object({
  op: z.literal('set_freeze'),
  sheetId: z.string().min(1),
  rows: z.number().int().min(0).max(100),
  columns: z.number().int().min(0).max(100),
})

const duplicateSheetSchema = z.object({
  op: z.literal('duplicate_sheet'),
  sheetId: z.string().min(1),
  name: sheetNameSchema.optional(),
})

const setSheetHiddenSchema = z.object({
  op: z.literal('set_sheet_hidden'),
  sheetId: z.string().min(1),
  hidden: z.boolean(),
})

const moveSheetSchema = z.object({
  op: z.literal('move_sheet'),
  sheetId: z.string().min(1),
  /** 1-based target tab position */
  position: z.number().int().min(1).max(1000),
})

// ref: an A1 reference like "Sheet1!$A$1:$B$10" or a constant/formula body
// (no leading =), as defined names are stored in the file.
const addDefinedNameSchema = z.object({
  op: z.literal('add_defined_name'),
  name: z.string().regex(/^[\p{L}_\\][\p{L}\p{N}_.\\]{0,254}$/u),
  ref: z.string().min(1).max(8192),
})

const deleteDefinedNameSchema = z.object({
  op: z.literal('delete_defined_name'),
  name: z.string().min(1).max(255),
})

// Per-cell edge semantics: every cell in the range gets the named edge(s);
// 'none' removes all borders. Clearing is 'none', so null is not accepted.
const borderPatchSchema = z.object({
  type: z.enum(['all', 'top', 'bottom', 'left', 'right', 'none']),
  color: hexColorSchema.optional(),
})

// Every field optional; null clears that property back to the default.
const formatPatchSchema = z
  .object({
    bold: z.boolean().nullable().optional(),
    italic: z.boolean().nullable().optional(),
    underline: z.boolean().nullable().optional(),
    strikethrough: z.boolean().nullable().optional(),
    fontFamily: z.string().min(1).max(128).nullable().optional(),
    fontSize: z.number().min(1).max(409).nullable().optional(),
    fontColor: hexColorSchema.nullable().optional(),
    fillColor: hexColorSchema.nullable().optional(),
    numberFormat: z.string().min(1).max(255).nullable().optional(),
    horizontalAlign: z.enum(['left', 'center', 'right']).nullable().optional(),
    verticalAlign: z.enum(['top', 'center', 'bottom']).nullable().optional(),
    wrapText: z.boolean().nullable().optional(),
    /** degrees, -90 (clockwise) to 90 (counterclockwise), or 'vertical' (stacked) */
    textRotation: z
      .union([z.number().int().min(-90).max(90), z.literal('vertical')])
      .nullable()
      .optional(),
    /** OOXML indent steps (0 or null clears); renders as left cell padding */
    indent: z.number().int().min(0).max(250).nullable().optional(),
    border: borderPatchSchema.optional(),
  })
  .refine((patch) => Object.values(patch).some((value) => value !== undefined), {
    message: 'format must set at least one property',
  })

const formatRangeSchema = z.object({
  op: z.literal('format_range'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  format: formatPatchSchema,
})

const sortRangeSchema = z.object({
  op: z.literal('sort_range'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  byColumn: columnLabelSchema,
  order: z.enum(['asc', 'desc']),
  hasHeader: z.boolean().optional(),
})

const mergeCellsSchema = z.object({
  op: z.literal('merge_cells'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

const unmergeCellsSchema = z.object({
  op: z.literal('unmerge_cells'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
})

const setRowHeightSchema = z.object({
  op: z.literal('set_row_height'),
  sheetId: z.string().min(1),
  /** 1-based first row */
  row: z.number().int().min(1).max(9999999),
  count: z.number().int().min(1).max(500).default(1),
  /** Excel points (2–409) */
  heightPoints: z.number().min(2).max(409),
})

const setColWidthSchema = z.object({
  op: z.literal('set_col_width'),
  sheetId: z.string().min(1),
  column: columnLabelSchema,
  count: z.number().int().min(1).max(100).default(1),
  /** pixels (10–2000) */
  widthPx: z.number().min(10).max(2000),
})

const renameSheetSchema = z.object({
  op: z.literal('rename_sheet'),
  sheetId: z.string().min(1),
  name: sheetNameSchema,
})

// Deletes a floating visual (chart/shape/image). Charts accept the chart
// part path or the visual id; shapes/images use the visual id (both listed
// in get_workbook_context / read_sheet_features).
const deleteVisualSchema = z.object({
  op: z.literal('delete_visual'),
  visualId: z.string().min(1).max(300),
})

// Removes a table created THIS session (converts it back to a plain range —
// values and formatting stay). File-native tables cannot be removed yet.
const deleteTableSchema = z.object({
  op: z.literal('delete_table'),
  sheetId: z.string().min(1),
  tableName: z.string().min(1).max(255),
})

// In-cell sparkline minicharts: one per dataRange row, hosted in the column
// right of the range (or starting at targetCell going down).
const addSparklineSchema = z.object({
  op: z.literal('add_sparkline'),
  sheetId: z.string().min(1),
  type: z.enum(['line', 'column', 'stacked']),
  dataRange: cellRangeSchema,
  targetCell: cellAddressSchema.optional(),
  color: hexColorSchema.optional(),
})

// Literal find & replace over TEXT values in a range (formulas and numbers
// are untouched). Expands to plain set_cell edits, so preview and undo
// behave like typing. Case-insensitive by default.
const findReplaceSchema = z.object({
  op: z.literal('find_replace'),
  sheetId: z.string().min(1),
  range: cellRangeSchema,
  find: z.string().min(1).max(255),
  replace: z.string().max(255),
  matchCase: z.boolean().optional(),
  /** match the whole cell text instead of substrings */
  wholeCell: z.boolean().optional(),
})

export const workbookOperationSchema = z.discriminatedUnion('op', [
  setCellSchema,
  setFormulaSchema,
  clearCellSchema,
  setRangeSchema,
  clearRangeSchema,
  formatRangeSchema,
  sortRangeSchema,
  mergeCellsSchema,
  unmergeCellsSchema,
  setRowHeightSchema,
  setColWidthSchema,
  editChartSchema,
  addChartSchema,
  addShapeSchema,
  editShapeSchema,
  addImageSchema,
  addTableSchema,
  addTableRowSchema,
  addTableColumnSchema,
  deleteTableRowSchema,
  deleteTableColumnSchema,
  addPivotSchema,
  setRowsHiddenSchema,
  setColsHiddenSchema,
  setHyperlinkSchema,
  protectSheetSchema,
  setFilterSchema,
  clearFilterSchema,
  setFilterCriteriaSchema,
  addConditionalFormatSchema,
  clearConditionalFormatsSchema,
  setDataValidationSchema,
  addDefinedNameSchema,
  deleteDefinedNameSchema,
  setPageSetupSchema,
  setFreezeSchema,
  setNoteSchema,
  refreshPivotSchema,
  insertRowsSchema,
  deleteRowsSchema,
  insertColsSchema,
  deleteColsSchema,
  addSheetSchema,
  deleteSheetSchema,
  duplicateSheetSchema,
  setSheetHiddenSchema,
  moveSheetSchema,
  renameSheetSchema,
  deleteVisualSchema,
  deleteTableSchema,
  findReplaceSchema,
  addSparklineSchema,
])

export type WorkbookOperation = z.infer<typeof workbookOperationSchema>
export type SetCellOperation = z.infer<typeof setCellSchema>
export type SetRangeOperation = z.infer<typeof setRangeSchema>
export type SetFormulaOperation = z.infer<typeof setFormulaSchema>
export type ClearCellOperation = z.infer<typeof clearCellSchema>
export type FormatRangeOperation = z.infer<typeof formatRangeSchema>
export type CellFormatPatch = z.infer<typeof formatPatchSchema>
export type BorderPatch = z.infer<typeof borderPatchSchema>
export type StructuralOperation =
  | z.infer<typeof insertRowsSchema>
  | z.infer<typeof deleteRowsSchema>
  | z.infer<typeof insertColsSchema>
  | z.infer<typeof deleteColsSchema>
  | z.infer<typeof addSheetSchema>
  | z.infer<typeof deleteSheetSchema>
  | z.infer<typeof duplicateSheetSchema>
  | z.infer<typeof setSheetHiddenSchema>
  | z.infer<typeof moveSheetSchema>
export type EditChartOperation = z.infer<typeof editChartSchema>
export type AddChartOperation = z.infer<typeof addChartSchema>
export type AddShapeOperation = z.infer<typeof addShapeSchema>
export type EditShapeOperation = z.infer<typeof editShapeSchema>
export type AddImageOperation = z.infer<typeof addImageSchema>
export type AddTableOperation = z.infer<typeof addTableSchema>
export type AddTableRowOperation = z.infer<typeof addTableRowSchema>
export type AddTableColumnOperation = z.infer<typeof addTableColumnSchema>
export type DeleteTableRowOperation = z.infer<typeof deleteTableRowSchema>
export type DeleteTableColumnOperation = z.infer<typeof deleteTableColumnSchema>
export type AddPivotOperation = z.infer<typeof addPivotSchema>
export type SetHyperlinkOperation = z.infer<typeof setHyperlinkSchema>
export type AddConditionalFormatOperation = z.infer<typeof addConditionalFormatSchema>
export type SetDataValidationOperation = z.infer<typeof setDataValidationSchema>
export type SetPageSetupOperation = z.infer<typeof setPageSetupSchema>
export type CfRule = z.infer<typeof cfRuleSchema>
export type DvRule = z.infer<typeof dvRuleSchema>
/** sheet-layout edits: no address shifts, so they may mix with content ops */
export type LayoutOperation =
  | z.infer<typeof mergeCellsSchema>
  | z.infer<typeof unmergeCellsSchema>
  | z.infer<typeof setRowHeightSchema>
  | z.infer<typeof setColWidthSchema>
  | EditChartOperation
  | AddChartOperation
  | AddShapeOperation
  | EditShapeOperation
  | AddImageOperation
  | AddTableOperation
  | AddTableRowOperation
  | AddTableColumnOperation
  | DeleteTableRowOperation
  | DeleteTableColumnOperation
  | AddPivotOperation
  | z.infer<typeof setRowsHiddenSchema>
  | z.infer<typeof setColsHiddenSchema>
  | SetHyperlinkOperation
  | z.infer<typeof protectSheetSchema>
  | z.infer<typeof setFilterSchema>
  | z.infer<typeof clearFilterSchema>
  | z.infer<typeof setFilterCriteriaSchema>
  | AddConditionalFormatOperation
  | z.infer<typeof clearConditionalFormatsSchema>
  | SetDataValidationOperation
  | z.infer<typeof addDefinedNameSchema>
  | z.infer<typeof deleteDefinedNameSchema>
  | SetPageSetupOperation
  | z.infer<typeof setFreezeSchema>
  | z.infer<typeof setNoteSchema>
  | z.infer<typeof refreshPivotSchema>
  | DeleteVisualOperation
  | DeleteTableOperation
  | AddSparklineOperation
export type DeleteVisualOperation = z.infer<typeof deleteVisualSchema>
export type DeleteTableOperation = z.infer<typeof deleteTableSchema>
export type AddSparklineOperation = z.infer<typeof addSparklineSchema>
export type FindReplaceOperation = z.infer<typeof findReplaceSchema>
export type CellContentOperation = SetCellOperation | SetFormulaOperation | ClearCellOperation
/** what range ops expand into; the only shapes executors have to handle */
export type PrimitiveOperation =
  | CellContentOperation
  | FormatRangeOperation
  | LayoutOperation
  | StructuralOperation
  | z.infer<typeof renameSheetSchema>

const STRUCTURAL_OPS = new Set([
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'add_sheet',
  'delete_sheet',
  'duplicate_sheet',
  'set_sheet_hidden',
  'move_sheet',
])
const LAYOUT_OPS = new Set([
  'merge_cells',
  'unmerge_cells',
  'set_row_height',
  'set_col_width',
  'edit_chart',
  'add_chart',
  'add_shape',
  'edit_shape',
  'add_image',
  'add_table',
  'add_pivot',
  'add_table_row',
  'add_table_column',
  'delete_table_row',
  'delete_table_column',
  'set_rows_hidden',
  'set_cols_hidden',
  'set_hyperlink',
  'protect_sheet',
  'set_filter',
  'clear_filter',
  'set_filter_criteria',
  'add_conditional_format',
  'clear_conditional_formats',
  'set_data_validation',
  'add_defined_name',
  'delete_defined_name',
  'set_page_setup',
  'set_freeze',
  'set_note',
  'refresh_pivot',
  'delete_visual',
  'delete_table',
  'add_sparkline',
])
// Content addresses cells, so it must not ride in the same batch as
// row/column shifts that would move those addresses. Layout ops keep their
// row/column coordinates too, so they count the same way.
const CELL_CONTENT_OPS = new Set([
  'set_cell',
  'set_formula',
  'clear_cell',
  'set_range',
  'clear_range',
  'format_range',
  'sort_range',
  'find_replace',
  ...LAYOUT_OPS,
])

export function isStructuralOp(
  op: WorkbookOperation | PrimitiveOperation,
): op is StructuralOperation {
  return STRUCTURAL_OPS.has(op.op)
}

export function isLayoutOp(op: WorkbookOperation | PrimitiveOperation): op is LayoutOperation {
  return LAYOUT_OPS.has(op.op)
}

export const workbookCommandBatchSchema = z
  .object({
    dslVersion: z.literal(1),
    transactionId: z.string().min(1).max(128),
    baseRevision: z.number().int().nonnegative(),
    summary: z.string().trim().min(1).max(500),
    operations: z.array(workbookOperationSchema).min(1).max(1000),
  })
  .refine(
    (batch) => {
      // Structural ops shift addresses, so mixing them with cell edits in one
      // batch would make previewed before/after states unverifiable. Propose
      // structure first, apply, then propose content.
      const hasStructural = batch.operations.some((op) => STRUCTURAL_OPS.has(op.op))
      const hasCellContent = batch.operations.some((op) => CELL_CONTENT_OPS.has(op.op))
      return !(hasStructural && hasCellContent)
    },
    {
      message:
        'Structural operations (rows/columns/sheets) and cell edits must be proposed in separate batches.',
    },
  )

export type WorkbookCommandBatch = z.infer<typeof workbookCommandBatchSchema>

export function parseWorkbookCommandBatch(input: unknown): WorkbookCommandBatch {
  return workbookCommandBatchSchema.parse(input)
}

export const MAX_EXPANDED_CELL_OPS = 2000

function replaceOccurrences(
  text: string,
  find: string,
  replace: string,
  matchCase: boolean,
): string {
  if (matchCase) return text.split(find).join(replace)
  const escaped = find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  // Callback form: a literal `$` in the replacement must stay literal.
  return text.replace(new RegExp(escaped, 'gi'), () => replace)
}

export type ExpandCellReader = (
  address: string,
  sheetId: string,
) => {
  value: string | number | boolean | null
  formula?: string | undefined
}

/// set_range guards, applied before any expansion. Jagged rows are rejected
/// because a shorter row silently leaves the old trailing cells in place —
/// the classic way a table rewrite shears its columns apart; requiring a
/// rectangle forces the writer to state the intended width (null clears a
/// cell). When the op targets a full `range`, its size must match the values
/// grid so misaligned writes fail before anything applies.
function setRangeOrigin(operation: SetRangeOperation): { startRow: number; startColumn: number } {
  const width = operation.values[0]?.length ?? 0
  const jaggedIndex = operation.values.findIndex((row) => row.length !== width)
  if (jaggedIndex !== -1) {
    throw new Error(
      `set_range values must be rectangular: row 1 has ${width} cell(s) but row ${jaggedIndex + 1} has ${operation.values[jaggedIndex]?.length}. ` +
        'Use null for cells that should be cleared, or split into separate set_range operations.',
    )
  }
  const { start, range } = operation
  if (!start && !range)
    throw new Error('set_range needs "start" — the top-left target cell, like "B2".')
  const bounds = parseRange(range ?? (start as string))
  if (start && range) {
    // Both fields together must agree; silently preferring one would let a
    // write land somewhere other than where the model believes it targeted.
    const startBounds = parseRange(start)
    if (
      startBounds.startRow !== bounds.startRow ||
      startBounds.startColumn !== bounds.startColumn
    ) {
      throw new Error(
        `set_range received both start ${start} and range ${range}, which disagree on the top-left cell — pass only one of them.`,
      )
    }
  }
  if (range) {
    const rows = bounds.endRow - bounds.startRow + 1
    const columns = bounds.endColumn - bounds.startColumn + 1
    if (rows !== operation.values.length || columns !== width) {
      throw new Error(
        `set_range range ${range} spans ${rows}×${columns} cells but values is ${operation.values.length} row(s) × ${width} cell(s) — ` +
          'make them match, or give "start" (the top-left cell) instead.',
      )
    }
  }
  return bounds
}

/// Expands set_range/clear_range into per-cell primitives so the preview,
/// CAS checks, and both apply paths keep working on single cells. sort_range
/// additionally needs `readCell` to compute the reordered values.
export function expandToPrimitiveOps(
  operations: readonly WorkbookOperation[],
  readCell?: ExpandCellReader,
): PrimitiveOperation[] {
  const expanded: PrimitiveOperation[] = []
  let cellOps = 0
  const countCell = (): void => {
    cellOps += 1
    if (cellOps > MAX_EXPANDED_CELL_OPS) {
      throw new Error(`The batch expands to more than ${MAX_EXPANDED_CELL_OPS} cell edits.`)
    }
  }
  for (const operation of operations) {
    if (operation.op === 'set_range') {
      const origin = setRangeOrigin(operation)
      operation.values.forEach((rowValues, rowOffset) => {
        rowValues.forEach((value, columnOffset) => {
          countCell()
          const address = formatAddress(
            origin.startRow + rowOffset,
            origin.startColumn + columnOffset,
          )
          if (typeof value === 'string' && value.startsWith('=')) {
            expanded.push({
              op: 'set_formula',
              sheetId: operation.sheetId,
              address,
              formula: value,
            })
          } else {
            expanded.push({ op: 'set_cell', sheetId: operation.sheetId, address, value })
          }
        })
      })
    } else if (operation.op === 'clear_range') {
      const bounds = parseRange(operation.range)
      if (rangeCellCount(bounds) > MAX_EXPANDED_CELL_OPS) {
        throw new Error(`The batch expands to more than ${MAX_EXPANDED_CELL_OPS} cell edits.`)
      }
      for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
        for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
          countCell()
          expanded.push({
            op: 'clear_cell',
            sheetId: operation.sheetId,
            address: formatAddress(row, column),
          })
        }
      }
    } else if (operation.op === 'format_range') {
      if (rangeCellCount(parseRange(operation.range)) > MAX_EXPANDED_CELL_OPS) {
        throw new Error(`format_range covers more than ${MAX_EXPANDED_CELL_OPS} cells.`)
      }
      expanded.push(operation)
    } else if (operation.op === 'find_replace') {
      if (!readCell)
        throw new Error('find_replace needs the current cell contents to plan against.')
      const bounds = parseRange(operation.range)
      if (rangeCellCount(bounds) > MAX_EXPANDED_CELL_OPS) {
        throw new Error(`find_replace covers more than ${MAX_EXPANDED_CELL_OPS} cells.`)
      }
      const matchCase = operation.matchCase ?? false
      const needle = matchCase ? operation.find : operation.find.toLowerCase()
      for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
        for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
          const address = formatAddress(row, column)
          const current = readCell(address, operation.sheetId)
          if (current.formula !== undefined || typeof current.value !== 'string') continue
          const haystack = matchCase ? current.value : current.value.toLowerCase()
          let next: string | null = null
          if (operation.wholeCell) {
            if (haystack === needle) next = operation.replace
          } else if (haystack.includes(needle)) {
            next = replaceOccurrences(current.value, operation.find, operation.replace, matchCase)
          }
          if (next === null || next === current.value) continue
          countCell()
          expanded.push({ op: 'set_cell', sheetId: operation.sheetId, address, value: next })
        }
      }
    } else if (operation.op === 'sort_range') {
      if (!readCell) throw new Error('sort_range needs the current cell contents to plan against.')
      if (rangeCellCount(parseRange(operation.range)) > MAX_EXPANDED_CELL_OPS) {
        throw new Error(`sort_range covers more than ${MAX_EXPANDED_CELL_OPS} cells.`)
      }
      const changes = computeSortChanges(
        {
          range: operation.range,
          byColumn: operation.byColumn,
          ascending: operation.order === 'asc',
          hasHeader: operation.hasHeader ?? false,
        },
        (address) => readCell(address, operation.sheetId),
      )
      for (const change of changes) {
        countCell()
        expanded.push({
          op: 'set_cell',
          sheetId: operation.sheetId,
          address: change.address,
          value: change.after,
          expectedValue: change.before,
        })
      }
    } else if (operation.op === 'add_pivot') {
      const columnFieldsArray =
        operation.columnField === undefined
          ? []
          : Array.isArray(operation.columnField)
            ? operation.columnField
            : [operation.columnField]
      if (columnFieldsArray.length > 0 && operation.values.length !== 1) {
        throw new Error('With a columnField the pivot supports exactly one values entry.')
      }
      if (new Set(columnFieldsArray).size !== columnFieldsArray.length) {
        throw new Error('columnField must not repeat a field.')
      }
      const rowFieldsArray = Array.isArray(operation.rowFields)
        ? operation.rowFields
        : [operation.rowFields]
      if (new Set(rowFieldsArray).size !== rowFieldsArray.length) {
        throw new Error('rowFields must not repeat a field.')
      }
      if (columnFieldsArray.some((field) => rowFieldsArray.includes(field))) {
        throw new Error('A field cannot be both a row and a column dimension.')
      }
      const allDimensionFields = [
        ...rowFieldsArray,
        ...columnFieldsArray,
        ...(operation.pageFields ?? []),
      ]
      if (operation.values.some((value) => allDimensionFields.includes(value.field))) {
        throw new Error('A values field cannot also be a row, column, or page filter field.')
      }
      // Calculated fields always aggregate with SUM (the formula operates on each
      // referenced field's in-group sum).
      if (operation.values.some((value) => value.formula !== undefined && value.agg !== 'sum')) {
        throw new Error('A calculated field must use agg "sum".')
      }
      // Grouping may only target row/column dimension fields, with at most one
      // rule per field.
      const axisFields = [...rowFieldsArray, ...columnFieldsArray]
      const groupedFields = (operation.groupings ?? []).map((grouping) => grouping.field)
      if (groupedFields.some((field) => !axisFields.includes(field))) {
        throw new Error('A grouping field must be a row or column dimension field.')
      }
      if (new Set(groupedFields).size !== groupedFields.length) {
        throw new Error('A field can only have one grouping rule.')
      }
      // Filters may only target row/column dimension fields, at most one per
      // field; value-filter params must be complete for the op, and the
      // referenced values entry must exist.
      const filteredFields = (operation.filters ?? []).map((filter) => filter.field)
      if (filteredFields.some((field) => !axisFields.includes(field))) {
        throw new Error('A filter field must be a row or column dimension field.')
      }
      if (new Set(filteredFields).size !== filteredFields.length) {
        throw new Error('A field can only have one filter.')
      }
      for (const filter of operation.filters ?? []) {
        if (filter.kind !== 'value') continue
        if (filter.valueIndex >= operation.values.length) {
          throw new Error(`Filter valueIndex ${filter.valueIndex} is out of range.`)
        }
        if (filter.op === 'top' && filter.count === undefined) {
          throw new Error('A top-N value filter needs "count".')
        }
        if (filter.op === 'greaterThan' && filter.from === undefined) {
          throw new Error('A greaterThan value filter needs "from".')
        }
        if (filter.op === 'between' && (filter.from === undefined || filter.to === undefined)) {
          throw new Error('A between value filter needs "from" and "to".')
        }
      }
      // Normalize rowFields → always array for downstream consumers
      expanded.push({
        ...operation,
        rowFields: rowFieldsArray,
      })
    } else if (operation.op === 'edit_shape') {
      if (
        operation.text === undefined &&
        operation.fillColor === undefined &&
        operation.anchorCell === undefined
      ) {
        throw new Error('edit_shape needs at least one of text / fillColor / anchorCell.')
      }
      expanded.push(operation)
    } else if (operation.op === 'edit_chart') {
      if (
        operation.title === undefined &&
        operation.chartType === undefined &&
        (!operation.seriesColors || Object.keys(operation.seriesColors).length === 0) &&
        operation.legend === undefined &&
        operation.axisTitles === undefined &&
        operation.dataLabels === undefined &&
        operation.grouping === undefined &&
        (!operation.seriesData || operation.seriesData.length === 0)
      ) {
        throw new Error(
          'edit_chart needs at least one of title / chartType / seriesColors / legend / dataLabels / grouping / axisTitles / seriesData.',
        )
      }
      for (const entry of operation.seriesData ?? []) {
        if (
          entry.name === undefined &&
          entry.valuesRange === undefined &&
          entry.categoriesRange === undefined
        ) {
          throw new Error(`seriesData[index=${entry.index}] needs a name or a data range.`)
        }
      }
      expanded.push(operation)
    } else if (operation.op === 'add_conditional_format') {
      const rule = operation.rule
      if (
        rule.kind === 'number' &&
        (rule.operator === 'between' || rule.operator === 'notBetween') &&
        rule.value2 === undefined
      ) {
        throw new Error(`add_conditional_format with operator "${rule.operator}" needs value2.`)
      }
      expanded.push(operation)
    } else if (operation.op === 'set_data_validation') {
      const rule = operation.validation
      if (rule?.kind === 'numberBetween' && rule.min > rule.max) {
        throw new Error('set_data_validation numberBetween needs min ≤ max.')
      }
      if (rule?.kind === 'dateBetween' && rule.start > rule.end) {
        throw new Error('set_data_validation dateBetween needs start ≤ end.')
      }
      expanded.push(operation)
    } else if (operation.op === 'set_page_setup') {
      const { op: _op, sheetId: _sheetId, ...settings } = operation
      if (Object.values(settings).every((value) => value === undefined)) {
        throw new Error('set_page_setup needs at least one setting.')
      }
      if (
        operation.scale !== undefined &&
        (operation.fitToWidth !== undefined || operation.fitToHeight !== undefined)
      ) {
        throw new Error('set_page_setup: scale and fitToWidth/fitToHeight are mutually exclusive.')
      }
      expanded.push(operation)
    } else {
      if (CELL_CONTENT_OPS.has(operation.op)) countCell()
      expanded.push(operation)
    }
  }
  return expanded
}

const FORMAT_FIELD_LABELS: Record<string, string> = {
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  strikethrough: 'strikethrough',
  fontFamily: 'font',
  fontSize: 'font size',
  fontColor: 'font color',
  fillColor: 'fill',
  numberFormat: 'number format',
  horizontalAlign: 'align',
  verticalAlign: 'vertical align',
  wrapText: 'wrap',
  textRotation: 'rotation',
  indent: 'indent',
}

export function formatOpLabel(op: FormatRangeOperation): string {
  const parts = Object.entries(op.format)
    .filter(([, value]) => value !== undefined)
    .map(([key, value]) => {
      const name = FORMAT_FIELD_LABELS[key] ?? key
      if (key === 'border') {
        const border = value as BorderPatch
        if (border.type === 'none') return 'clear borders'
        return `border ${border.type}${border.color ? ` ${border.color}` : ''}`
      }
      if (value === null) return `clear ${name}`
      if (value === true) return name
      if (value === false) return `no ${name}`
      return `${name} ${String(value)}`
    })
  return `${op.range}: ${parts.join(', ')}`
}

export function layoutOpLabel(op: LayoutOperation): string {
  switch (op.op) {
    case 'delete_visual':
      return `Delete visual ${op.visualId}`
    case 'delete_table':
      return `Delete table "${op.tableName}"`
    case 'add_sparkline':
      return `Insert ${op.type} sparklines for ${op.dataRange}`
    case 'merge_cells':
      return `Merge ${op.range}`
    case 'unmerge_cells':
      return `Unmerge ${op.range}`
    case 'set_row_height':
      return op.count === 1
        ? `Set row ${op.row} height to ${op.heightPoints}pt`
        : `Set rows ${op.row}–${op.row + op.count - 1} height to ${op.heightPoints}pt`
    case 'set_col_width':
      return op.count === 1
        ? `Set column ${op.column} width to ${op.widthPx}px`
        : `Set columns ${op.column}–${columnLabel(columnIndex(op.column) + op.count - 1)} width to ${op.widthPx}px`
    case 'edit_chart': {
      const parts: string[] = []
      if (op.title !== undefined) parts.push(`title "${op.title}"`)
      if (op.chartType !== undefined) parts.push(`type ${op.chartType}`)
      if (op.seriesColors && Object.keys(op.seriesColors).length > 0) {
        parts.push(
          `series colors ${Object.entries(op.seriesColors)
            .map(([i, c]) => `#${i}→${c}`)
            .join(', ')}`,
        )
      }
      if (op.legend !== undefined) {
        parts.push(op.legend === 'none' ? 'hide legend' : `legend ${op.legend}`)
      }
      if (op.axisTitles?.category !== undefined) {
        parts.push(
          op.axisTitles.category === null
            ? 'clear category-axis title'
            : `category axis "${op.axisTitles.category}"`,
        )
      }
      if (op.axisTitles?.value !== undefined) {
        parts.push(
          op.axisTitles.value === null
            ? 'clear value-axis title'
            : `value axis "${op.axisTitles.value}"`,
        )
      }
      for (const entry of op.seriesData ?? []) {
        const detail = [
          entry.name === undefined ? '' : `name "${entry.name}"`,
          entry.valuesRange === undefined ? '' : `values ${entry.valuesRange}`,
          entry.categoriesRange === undefined ? '' : `categories ${entry.categoriesRange}`,
        ]
          .filter(Boolean)
          .join(' ')
        parts.push(`series #${entry.index} ${detail}`)
      }
      return `Edit chart ${op.chartPath}: ${parts.join(', ')}`
    }
    case 'add_chart':
      return `Insert ${op.chartType} chart from ${op.dataRange}${op.title ? ` "${op.title}"` : ''}`
    case 'add_shape':
      return `Insert ${op.shapeType === 'textbox' ? 'text box' : `${op.shapeType} shape`} at ${op.anchorCell}`
    case 'edit_shape': {
      const parts = [
        op.text === undefined
          ? ''
          : `text "${op.text.length > 40 ? `${op.text.slice(0, 40)}…` : op.text}"`,
        op.fillColor === undefined ? '' : `fill ${op.fillColor}`,
        op.anchorCell === undefined ? '' : `move to ${op.anchorCell}`,
      ].filter(Boolean)
      return `Edit shape ${op.visualId}: ${parts.join(', ')}`
    }
    case 'add_image':
      return `Insert image ${op.path.split('/').pop() ?? op.path} at ${op.anchorCell}`
    case 'add_table':
      return `Create table${op.name ? ` ${op.name}` : ''} over ${op.range}`
    case 'add_table_row':
      return `Insert ${op.count ?? 1} row(s) into table ${op.tableName}${op.row !== undefined ? ` at position ${op.row}` : ' (append)'}`
    case 'add_table_column':
      return `Insert column "${op.columnName}" into table ${op.tableName}`
    case 'delete_table_row':
      return `Delete ${op.count ?? 1} row(s) from table ${op.tableName} at position ${op.row}`
    case 'delete_table_column':
      return `Delete ${op.count ?? 1} column(s) from table ${op.tableName} at position ${op.column}`
    case 'add_pivot': {
      const rowFieldsArray = Array.isArray(op.rowFields) ? op.rowFields : [op.rowFields]
      const columnFieldsArray =
        op.columnField === undefined
          ? []
          : Array.isArray(op.columnField)
            ? op.columnField
            : [op.columnField]
      return (
        `Create pivot${op.name ? ` ${op.name}` : ''} from ${op.sourceRange} ` +
        `(rows: ${rowFieldsArray.join(' > ')}${columnFieldsArray.length > 0 ? `, columns: ${columnFieldsArray.join(' > ')}` : ''}` +
        `${op.pageFields && op.pageFields.length > 0 ? `, filters: ${op.pageFields.join(', ')}` : ''}, ` +
        `values: ${op.values.map((value) => `${value.agg} ${value.field}`).join(', ')}) at ${op.targetCell}`
      )
    }
    case 'set_rows_hidden':
      return `${op.hidden ? 'Hide' : 'Unhide'} row${op.count === 1 ? ` ${op.row}` : `s ${op.row}–${op.row + op.count - 1}`}`
    case 'set_cols_hidden':
      return `${op.hidden ? 'Hide' : 'Unhide'} column${
        op.count === 1
          ? ` ${op.column}`
          : `s ${op.column}–${columnLabel(columnIndex(op.column) + op.count - 1)}`
      }`
    case 'set_hyperlink':
      return op.target === null
        ? `Remove link at ${op.address}`
        : `Link ${op.address} → ${op.target}`
    case 'protect_sheet':
      return op.protected ? 'Protect sheet' : 'Unprotect sheet'
    case 'set_filter':
      return `Add auto-filter on ${op.range}`
    case 'clear_filter':
      return 'Remove auto-filter'
    case 'set_filter_criteria':
      return op.values === null
        ? `Clear filter criteria on column ${op.column}`
        : `Filter column ${op.column} to ${op.values.length} value${op.values.length === 1 ? '' : 's'}`
    case 'add_conditional_format':
      return `Conditional format ${op.range}: ${cfRuleLabel(op.rule)}`
    case 'clear_conditional_formats':
      return 'Clear conditional formats'
    case 'set_data_validation':
      return op.validation === null
        ? `Clear data validation on ${op.range}`
        : `Data validation ${op.range}: ${dvRuleLabel(op.validation)}`
    case 'add_defined_name':
      return `Define name ${op.name} = ${op.ref}`
    case 'delete_defined_name':
      return `Delete defined name ${op.name}`
    case 'set_page_setup': {
      const parts: string[] = []
      if (op.orientation !== undefined) parts.push(op.orientation)
      if (op.paperSize !== undefined) parts.push(`paper ${op.paperSize}`)
      if (op.scale !== undefined) parts.push(`scale ${op.scale}%`)
      if (op.fitToWidth !== undefined) parts.push(`fit width ${op.fitToWidth || 'auto'}`)
      if (op.fitToHeight !== undefined) parts.push(`fit height ${op.fitToHeight || 'auto'}`)
      if (op.margins !== undefined) parts.push(`${op.margins} margins`)
      if (op.printGridlines !== undefined)
        parts.push(`${op.printGridlines ? 'print' : 'no'} gridlines`)
      if (op.printHeadings !== undefined)
        parts.push(`${op.printHeadings ? 'print' : 'no'} headings`)
      if (op.printArea !== undefined)
        parts.push(op.printArea === null ? 'clear print area' : `print area ${op.printArea}`)
      return `Page setup: ${parts.join(', ')}`
    }
    case 'set_freeze':
      return op.rows === 0 && op.columns === 0
        ? 'Unfreeze panes'
        : `Freeze ${[
            op.rows > 0 ? `first ${op.rows} row${op.rows === 1 ? '' : 's'}` : '',
            op.columns > 0 ? `first ${op.columns} column${op.columns === 1 ? '' : 's'}` : '',
          ]
            .filter(Boolean)
            .join(' and ')}`
    case 'set_note':
      return op.text === null
        ? `Remove note at ${op.address}`
        : `Note at ${op.address}: "${op.text.length > 60 ? `${op.text.slice(0, 60)}…` : op.text}"`
    case 'refresh_pivot':
      return 'Refresh pivot tables'
  }
}

function cfRuleLabel(rule: CfRule): string {
  switch (rule.kind) {
    case 'number':
      return `${rule.operator} ${rule.value}${rule.value2 === undefined ? '' : ` and ${rule.value2}`}`
    case 'text':
      return `${rule.operator} "${rule.text}"`
    case 'blank':
      return rule.blank ? 'is empty' : 'is not empty'
    case 'duplicate':
      return rule.unique ? 'unique values' : 'duplicate values'
    case 'top10':
      return `${rule.bottom ? 'bottom' : 'top'} ${rule.rank}${rule.percent ? '%' : ''}`
    case 'formula':
      return rule.formula
    case 'colorScale':
      return `color scale ${rule.minColor} → ${rule.maxColor}`
    case 'dataBar':
      return `data bar${rule.color ? ` ${rule.color}` : ''}`
  }
}

function dvRuleLabel(rule: DvRule): string {
  switch (rule.kind) {
    case 'list':
      return `list (${rule.values.slice(0, 5).join(', ')}${rule.values.length > 5 ? ', …' : ''})`
    case 'listRef':
      return `list from ${rule.range}`
    case 'numberBetween':
      return `number between ${rule.min} and ${rule.max}`
    case 'dateBetween':
      return `date between ${rule.start} and ${rule.end}`
    case 'checkbox':
      return 'checkbox'
    case 'formula':
      return `custom ${rule.formula}`
  }
}

export function structuralOpLabel(op: StructuralOperation): string {
  switch (op.op) {
    case 'insert_rows':
      return op.count === 1
        ? `Insert 1 row before row ${op.row}`
        : `Insert ${op.count} rows before row ${op.row}`
    case 'delete_rows':
      return op.count === 1
        ? `Delete row ${op.row}`
        : `Delete rows ${op.row}–${op.row + op.count - 1}`
    case 'insert_cols':
      return op.count === 1
        ? `Insert 1 column before column ${op.column}`
        : `Insert ${op.count} columns before column ${op.column}`
    case 'delete_cols':
      return op.count === 1
        ? `Delete column ${op.column}`
        : `Delete columns ${op.column}–${columnLabel(columnIndex(op.column) + op.count - 1)}`
    case 'add_sheet':
      return `Add sheet "${op.name}"`
    case 'delete_sheet':
      return `Delete sheet ${op.sheetId}`
    case 'duplicate_sheet':
      return `Duplicate sheet ${op.sheetId}${op.name ? ` as "${op.name}"` : ''}`
    case 'set_sheet_hidden':
      return `${op.hidden ? 'Hide' : 'Unhide'} sheet ${op.sheetId}`
    case 'move_sheet':
      return `Move sheet ${op.sheetId} to position ${op.position}`
  }
}
