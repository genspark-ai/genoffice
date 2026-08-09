import { z } from 'zod'

import { ADDABLE_SHAPE_TYPES } from './shape-types'
import type {
  AiChatRequest,
  AiChatResponse,
  AiSettings,
  AiStreamChunk,
  AiStreamRequest,
  GenSparkAccountStatus,
} from '@genoffice/ai-provider'

const MAX_RANGE_CELLS = 20_000
const cellScalarSchema = z.union([z.string(), z.number().finite(), z.boolean(), z.null()])
const cellAreaSchema = z
  .object({
    startRow: z.number().int().nonnegative(),
    startColumn: z.number().int().nonnegative(),
    endRow: z.number().int().nonnegative(),
    endColumn: z.number().int().nonnegative(),
  })
  .strict()
const worksheetMetadataSchema = z
  .object({
    id: z.string().min(1),
    name: z.string().min(1),
    rowCount: z.number().int().positive(),
    columnCount: z.number().int().positive(),
    columnWidths: z.array(
      z
        .object({
          startColumn: z.number().int().nonnegative(),
          endColumn: z.number().int().nonnegative(),
          width: z.number().nonnegative().optional(),
          hidden: z.boolean(),
          outlineLevel: z.number().int().min(1).max(7).optional(),
          collapsed: z.boolean().optional(),
        })
        .strict(),
    ),
    // 0 is legal in the file (sheetFormatPr on all-hidden sheets) and means
    // "no default"; the engine normalizes it to null, but a stray 0 must not
    // reject the whole workbook — the preload maps it back to null.
    defaultRowHeight: z.number().nonnegative().nullable(),
    defaultColumnWidth: z.number().nonnegative().nullable(),
    freeze: z
      .object({
        frozenColumns: z.number().int().nonnegative(),
        frozenRows: z.number().int().nonnegative(),
      })
      .strict()
      .nullable(),
    hidden: z.boolean(),
    tabColor: z.string().nullable(),
    showGridLines: z.boolean(),
    /// sheetView/@showFormulas — the sheet opens in formula view.
    showFormulas: z.boolean().optional(),
    tables: z.array(
      z
        .object({
          range: cellAreaSchema,
          headerRowCount: z.number().int().nonnegative(),
          showRowStripes: z.boolean(),
          showColumnStripes: z.boolean(),
          styleName: z.string().optional(),
          headerFill: z.string().optional(),
          headerFontColor: z.string().optional(),
          stripeFill: z.string().optional(),
        })
        .strict(),
    ),
    comments: z.array(
      z
        .object({
          row: z.number().int().nonnegative(),
          column: z.number().int().nonnegative(),
          author: z.string(),
          text: z.string(),
        })
        .strict(),
    ),
    /// PivotTable output areas; the renderer blocks edits inside them.
    pivotRanges: z.array(cellAreaSchema).max(1_000),
    /// Part paths for on-demand pivot definition reads (refresh support).
    /// Defaulted so a stale sidecar binary degrades to "no refresh" gracefully.
    pivotTables: z
      .array(
        z
          .object({
            path: z.string().min(1),
            cachePath: z.string().min(1).nullable(),
            outputRef: z.string().min(1),
          })
          .strict(),
      )
      .max(100)
      .default([]),
    /// x14 sparkline groups (in-cell line/column/win-loss minicharts).
    /// Defaulted for stale sidecar binaries, like pivotTables.
    sparklines: z
      .array(
        z
          .object({
            type: z.enum(['line', 'column', 'stacked']),
            color: z.string().optional(),
            negativeColor: z.string().optional(),
            cells: z
              .array(
                z
                  .object({
                    /// A1 address of the host cell.
                    cell: z.string().min(2),
                    /// `Sheet1!A2:C2` source range (sheet-qualified, `$` allowed).
                    sourceRef: z.string().min(1),
                  })
                  .strict(),
              )
              .max(500),
          })
          .strict(),
      )
      .max(100)
      .default([]),
  })
  .strict()
const richRunSchema = z
  .object({
    text: z.string(),
    bold: z.boolean(),
    italic: z.boolean(),
    underline: z.boolean(),
    strikethrough: z.boolean(),
    color: z.string().optional(),
    size: z.number().positive().optional(),
    family: z.string().optional(),
  })
  .strict()
const conditionalRuleSchema = z
  .object({
    ranges: z.array(cellAreaSchema),
    ruleType: z.string(),
    operator: z.string().optional(),
    formulas: z.array(z.string()),
    text: z.string().optional(),
    dxfIndex: z.number().int().nonnegative().optional(),
    priority: z.number().int(),
    rank: z.number().int().nonnegative().optional(),
    percent: z.boolean(),
    bottom: z.boolean(),
    cfvos: z.array(
      z
        .object({
          kind: z.string(),
          value: z.string().optional(),
          gte: z.boolean().optional(),
        })
        .strict(),
    ),
    colors: z.array(z.string()),
    iconSetName: z.string().optional(),
    iconReverse: z.boolean(),
    showValue: z.boolean(),
  })
  .strict()
const borderEdgeSchema = z
  .object({
    style: z.string().min(1),
    color: z.string().optional(),
  })
  .strict()
const cellStyleSchema = z
  .object({
    fontFamily: z.string().optional(),
    fontSize: z.number().positive().optional(),
    bold: z.boolean(),
    italic: z.boolean(),
    underline: z.boolean(),
    strikethrough: z.boolean(),
    wrapText: z.boolean(),
    fontColor: z.string().optional(),
    fillColor: z.string().optional(),
    horizontalAlignment: z.string().optional(),
    verticalAlignment: z.string().optional(),
    /// OOXML alignment indent steps read from the xf; absent when 0.
    indent: z.number().int().nonnegative().optional(),
    numberFormat: z.string().optional(),
    borderTop: borderEdgeSchema.optional(),
    borderBottom: borderEdgeSchema.optional(),
    borderLeft: borderEdgeSchema.optional(),
    borderRight: borderEdgeSchema.optional(),
    borderDiagonal: borderEdgeSchema.optional(),
    diagonalUp: z.boolean(),
    diagonalDown: z.boolean(),
  })
  .strict()
const drawingAnchorSchema = z
  .object({
    fromRow: z.number().int().nonnegative(),
    fromColumn: z.number().int().nonnegative(),
    fromRowOffset: z.number().int(),
    fromColumnOffset: z.number().int(),
    toRow: z.number().int().nonnegative(),
    toColumn: z.number().int().nonnegative(),
    toRowOffset: z.number().int(),
    toColumnOffset: z.number().int(),
  })
  .strict()
const visualObjectSchema = z
  .object({
    id: z.string().min(1),
    sheetId: z.string().min(1),
    kind: z.enum(['chart', 'image', 'shape']),
    anchor: drawingAnchorSchema,
    chart: z
      .object({
        chartTypes: z.array(z.string()),
        barDirection: z.string().optional(),
        title: z.string(),
        series: z.array(
          z
            .object({
              name: z.string(),
              categories: z.array(z.string()),
              values: z.array(z.number().finite()),
              numberFormat: z.string().optional(),
              /// numCache formatCode of the category (or scatter X) data.
              categoryFormat: z.string().optional(),
              color: z.string().optional(),
              trendline: z.string().optional(),
              /// `c:f` range references (groundwork for E5 data-range editing).
              valuesRef: z.string().optional(),
              categoriesRef: z.string().optional(),
              /// Explicit per-point fills (`c:dPt`), e.g. pie slice colors.
              pointColors: z
                .array(
                  z
                    .object({
                      index: z.number().int().nonnegative(),
                      color: z.string(),
                    })
                    .strict(),
                )
                .optional(),
              /// Pie `c:ser/c:explosion` (% of radius) and per-slice overrides.
              explosionPct: z.number().optional(),
              pointExplosions: z
                .array(
                  z
                    .object({
                      index: z.number().int().nonnegative(),
                      pct: z.number(),
                    })
                    .strict(),
                )
                .optional(),
            })
            .strict(),
        ),
        /// Parsed from the chart part by the sidecar; session edits overlay the
        /// same fields so the on-screen chart previews the pending save.
        legend: z.enum(['none', 'right', 'bottom', 'top', 'left']).optional(),
        axisTitles: z
          .object({
            category: z.string().nullable().optional(),
            value: z.string().nullable().optional(),
          })
          .strict()
          .optional(),
        dataLabels: z.enum(['none', 'value', 'percent', 'category-percent']).optional(),
        /// `c:dLblPos` and `c:numFmt` on the plot-level data labels.
        dataLabelPosition: z.enum(['center', 'inside-end', 'outside-end']).optional(),
        dataLabelFormat: z.string().optional(),
        /// bar/line/area `c:grouping`; absent for pie/scatter and older parts.
        grouping: z.enum(['clustered', 'stacked', 'percentStacked', 'standard']).optional(),
        /// Value-axis major gridlines; absent on pie/doughnut.
        gridlines: z.boolean().optional(),
        /// Explicit value-axis bounds (`c:scaling`); absent keys mean auto.
        valueAxis: z
          .object({
            min: z.number().finite().optional(),
            max: z.number().finite().optional(),
          })
          .strict()
          .optional(),
        /// `c:numFmt` on the category/date axis; wins over the series-level
        /// categoryFormat.
        categoryAxisFormat: z.string().optional(),
        gapWidthPct: z.number().optional(),
        holeSizePct: z.number().optional(),
      })
      .strict()
      .optional(),
    chartPath: z.string().optional(),
    mediaPath: z.string().optional(),
    mediaType: z.string().optional(),
    /// Renderer-only: preview bytes for images added this session (not yet in
    /// the package); the sidecar never emits this.
    mediaDataUrl: z.string().optional(),
    name: z.string().optional(),
    shapeType: z.string().optional(),
    fillColor: z.string().optional(),
    text: z.string().optional(),
    rotation: z.number().finite().optional(),
    /// Save-side edit locator: the drawing part this visual lives in and its
    /// anchor index within that part (document order, skipped anchors counted).
    drawingPath: z
      .string()
      .regex(/^xl\/drawings\/[A-Za-z0-9._/-]+\.xml$/)
      .optional(),
    drawingIndex: z.number().int().nonnegative().max(10_000).optional(),
  })
  .strict()

export const workbookFileSchema = z
  .object({
    sessionId: z.string().uuid(),
    name: z.string().min(1),
    /// Absolute on-disk path; feeds CELL("filename"). Absent for sessions
    /// that never touched disk.
    path: z.string().min(1).optional(),
    sha256: z.string().length(64),
    entryCount: z.number().int().nonnegative(),
    sheets: z.array(worksheetMetadataSchema).min(1),
    styles: z.array(cellStyleSchema),
    dxfStyles: z.array(cellStyleSchema),
    visuals: z.array(visualObjectSchema),
    definedNames: z.array(
      z
        .object({
          name: z.string().min(1),
          formula: z.string().min(1),
          /// localSheetId: position in workbook sheet order (sheet-scoped names).
          sheetIndex: z.number().int().nonnegative().optional(),
        })
        .strict(),
    ),
    readOnly: z.boolean(),
    /// Converted import (.xls/.csv): the first save opens a Save As dialog,
    /// so background flows (AutoSave) must not trigger mode 'save'.
    needsSaveAs: z.boolean().optional(),
  })
  .strict()

export const workbookRangeRequestSchema = z
  .object({
    sessionId: z.string().uuid(),
    sheetId: z.string().min(1),
    range: z
      .object({
        startRow: z.number().int().nonnegative(),
        endRow: z.number().int().nonnegative(),
        startColumn: z.number().int().nonnegative(),
        endColumn: z.number().int().nonnegative(),
      })
      .strict(),
  })
  .strict()
  .superRefine((request, context) => {
    const { range } = request
    if (range.startRow > range.endRow || range.startColumn > range.endColumn) {
      context.addIssue({ code: 'custom', message: 'Range boundaries are reversed.' })
      return
    }
    const rows = range.endRow - range.startRow + 1
    const columns = range.endColumn - range.startColumn + 1
    if (rows * columns > MAX_RANGE_CELLS) {
      context.addIssue({ code: 'custom', message: `Range exceeds ${MAX_RANGE_CELLS} cells.` })
    }
  })

const workbookCellRecordSchema = z
  .object({
    row: z.number().int().nonnegative(),
    column: z.number().int().nonnegative(),
    value: cellScalarSchema,
    formula: z.string().optional(),
    /// `<f t="array" ref>`: the legacy CSE range this master formula fills.
    arrayRef: z.string().max(64).optional(),
    styleIndex: z.number().int().nonnegative().optional(),
    rich: z.array(richRunSchema).optional(),
  })
  .strict()

export const workbookRangeResultSchema = z
  .object({
    cells: z.array(workbookCellRecordSchema).max(MAX_RANGE_CELLS),
    rows: z
      .array(
        z
          .object({
            row: z.number().int().nonnegative(),
            height: z.number().nonnegative().optional(),
            hidden: z.boolean(),
            outlineLevel: z.number().int().min(1).max(7).optional(),
            collapsed: z.boolean().optional(),
          })
          .strict(),
      )
      .max(MAX_RANGE_CELLS),
    merges: z.array(cellAreaSchema).max(MAX_RANGE_CELLS),
    hyperlinks: z
      .array(
        z
          .object({
            row: z.number().int().nonnegative(),
            column: z.number().int().nonnegative(),
            target: z.string(),
          })
          .strict(),
      )
      .max(MAX_RANGE_CELLS),
    conditionalRules: z.array(conditionalRuleSchema).max(MAX_RANGE_CELLS),
    autoFilter: cellAreaSchema.nullable(),
    dataValidations: z
      .array(
        z
          .object({
            ranges: z.array(cellAreaSchema),
            ruleType: z.string(),
            operator: z.string().optional(),
            formulas: z.array(z.string()),
            allowBlank: z.boolean(),
            /// Raw OOXML showDropDown flag — true SUPPRESSES the in-cell dropdown.
            suppressDropdown: z.boolean(),
            showInputMessage: z.boolean(),
            showErrorMessage: z.boolean(),
            errorStyle: z.string().optional(),
            errorTitle: z.string().optional(),
            error: z.string().optional(),
            promptTitle: z.string().optional(),
            prompt: z.string().optional(),
          })
          .strict(),
      )
      .max(MAX_RANGE_CELLS),
    /// Sheet-wide, delivered complete-only (like autoFilter). null = no
    /// <sheetProtection> element in the worksheet.
    sheetProtection: z
      .object({
        protected: z.boolean(),
        hasPassword: z.boolean(),
      })
      .strict()
      .nullable(),
    indexedThroughRow: z.number().int().nonnegative().nullable(),
    indexingComplete: z.boolean(),
  })
  .strict()

export const workbookFormulaCellsRequestSchema = z
  .object({
    sessionId: z.string().uuid(),
    sheetId: z.string().min(1),
  })
  .strict()

/// All formula cells of one sheet, for closure-mode analysis.
export const workbookFormulaCellsResultSchema = z
  .object({
    cells: z.array(workbookCellRecordSchema).max(100_000),
    indexingComplete: z.boolean(),
    truncated: z.boolean(),
  })
  .strict()

/// IronCalc recalculation fallback (closure mode unavailable): pending edits
/// go in as user input (formulas verbatim), evaluated values come back for the
/// requested ranges. The main process resolves sheet ids to file sheet names
/// — sheets added this session have no file part and fail the request.
export const workbookRecalcRequestSchema = z
  .object({
    sessionId: z.string().uuid(),
    edits: z
      .array(
        z
          .object({
            sheetId: z.string().min(1),
            row: z.number().int().nonnegative().max(1_048_575),
            column: z.number().int().nonnegative().max(16_383),
            input: z.string().max(32_767),
          })
          .strict(),
      )
      .max(10_000),
    reads: z
      .array(
        z
          .object({
            sheetId: z.string().min(1),
            range: z
              .object({
                startRow: z.number().int().nonnegative(),
                endRow: z.number().int().nonnegative(),
                startColumn: z.number().int().nonnegative(),
                endColumn: z.number().int().nonnegative(),
              })
              .strict(),
          })
          .strict(),
      )
      .max(200),
  })
  .strict()
  .superRefine((request, context) => {
    let total = 0
    for (const read of request.reads) {
      const { range } = read
      if (range.startRow > range.endRow || range.startColumn > range.endColumn) {
        context.addIssue({ code: 'custom', message: 'Range boundaries are reversed.' })
        return
      }
      total += (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1)
    }
    if (total > MAX_RANGE_CELLS) {
      context.addIssue({ code: 'custom', message: `Recalc read exceeds ${MAX_RANGE_CELLS} cells.` })
    }
  })

export const workbookRecalcResultSchema = z
  .object({
    cells: z
      .array(
        z
          .object({
            sheetId: z.string().min(1),
            row: z.number().int().nonnegative(),
            column: z.number().int().nonnegative(),
            formatted: z.string(),
            number: z.number().optional(),
            isFormula: z.boolean(),
          })
          .strict(),
      )
      .max(MAX_RANGE_CELLS),
  })
  .strict()

const MAX_SAVE_EDITS = 10_000
const hexColorSchema = z.string().regex(/^#[0-9A-Fa-f]{6}$/)

/// OOXML border line styles the editor can write.
export const editableBorderStyleSchema = z.enum([
  'thin',
  'medium',
  'thick',
  'dashed',
  'dotted',
  'double',
  'hair',
  'dashDot',
  'dashDotDot',
  'mediumDashed',
  'mediumDashDot',
  'mediumDashDotDot',
  'slantDashDot',
])

/// One border edge delta: an object sets the edge, null removes it.
const styleEditBorderSchema = z.union([
  z
    .object({
      style: editableBorderStyleSchema,
      color: hexColorSchema.optional(),
    })
    .strict(),
  z.null(),
])

/// Renderer-neutral style delta: only keys the user changed are present.
/// `false` means "remove this attribute from the cell's style".
export const workbookStyleEditSchema = z
  .object({
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    underlineStyle: z.enum(['single', 'double']).optional(),
    strikethrough: z.boolean().optional(),
    fontFamily: z.string().min(1).max(128).optional(),
    fontSize: z.number().positive().max(409).optional(),
    /// null removes the explicit font color (back to the theme default).
    fontColor: z.union([hexColorSchema, z.null()]).optional(),
    /// null clears the fill back to the default "none" pattern.
    fillColor: z.union([hexColorSchema, z.null()]).optional(),
    horizontalAlignment: z.enum(['left', 'center', 'right', 'justify', 'distributed']).optional(),
    verticalAlignment: z.enum(['top', 'center', 'bottom']).optional(),
    wrapText: z.boolean().optional(),
    /// OOXML textRotation: 0-90 counterclockwise, 91-180 clockwise (value-90),
    /// 255 stacked vertical; 0 clears the rotation.
    textRotation: z.union([z.number().int().min(0).max(180), z.literal(255)]).optional(),
    /// OOXML alignment indent steps; 0 clears. Renders on screen as left cell
    /// padding (INDENT_STEP_PX per step).
    indent: z.number().int().min(0).max(250).optional(),
    /// Cell protection flags (xf <protection>); meaningful once the sheet is
    /// protected. true = OOXML default for locked, false for hidden.
    protectionLocked: z.boolean().optional(),
    protectionHidden: z.boolean().optional(),
    numberFormat: z.string().min(1).max(255).optional(),
    borderTop: styleEditBorderSchema.optional(),
    borderBottom: styleEditBorderSchema.optional(),
    borderLeft: styleEditBorderSchema.optional(),
    borderRight: styleEditBorderSchema.optional(),
  })
  .strict()

export const workbookCellEditSchema = z
  .object({
    sheetId: z.string().min(1),
    row: z.number().int().nonnegative().max(1_048_575),
    column: z.number().int().nonnegative().max(16_383),
    /// false = style-only edit; the cell's stored content must stay untouched.
    writeValue: z.boolean(),
    value: cellScalarSchema,
    formula: z.string().min(1).max(8_192).optional(),
    style: workbookStyleEditSchema.optional(),
    /// Per-run styling for a rich-text value; `value` holds the joined text.
    rich: z.array(richRunSchema).max(1_000).optional(),
    /// Clear Formats / Clear All: reset the cell to the default style before
    /// applying any `style` delta.
    styleReset: z.boolean().optional(),
  })
  .strict()
  .refine((edit) => edit.writeValue || edit.style !== undefined || edit.styleReset === true, {
    message: 'A style-only edit needs a style delta or a reset.',
  })

/// One structural operation, in the coordinate space produced by all
/// preceding operations (they replay in order at save time).
export const workbookStructuralOpSchema = z.union([
  z
    .object({
      sheetId: z.string().min(1),
      kind: z.enum(['insert-rows', 'remove-rows', 'insert-cols', 'remove-cols']),
      index: z.number().int().nonnegative().max(1_048_575),
      count: z.number().int().positive().max(10_000),
    })
    .strict(),
  z
    .object({
      sheetId: z.string().min(1),
      kind: z.literal('move-rows'),
      index: z.number().int().nonnegative().max(1_048_575),
      count: z.number().int().positive().max(10_000),
      /// Pre-move insertion row; must lie outside the moved block.
      before: z.number().int().nonnegative().max(1_048_576),
    })
    .strict()
    .refine((op) => op.before < op.index || op.before > op.index + op.count, {
      message: 'A row move target must lie outside the moved block.',
    }),
  z
    .object({
      sheetId: z.string().min(1),
      kind: z.enum(['merge-cells', 'unmerge-cells']),
      range: cellAreaSchema,
    })
    .strict(),
  z
    .object({
      sheetId: z.string().min(1),
      kind: z.enum(['set-row-size', 'set-col-size']),
      start: z.number().int().nonnegative().max(1_048_575),
      end: z.number().int().nonnegative().max(1_048_575),
      /// File units: points for rows (Excel max 409.5), character width for
      /// columns (max 255). null resets to the sheet default.
      size: z.union([z.number().positive().max(500), z.null()]),
    })
    .strict()
    .refine((op) => op.end >= op.start && op.end - op.start < 100_000, {
      message: 'Invalid axis span.',
    }),
  z
    .object({
      sheetId: z.string().min(1),
      kind: z.enum(['set-rows-hidden', 'set-cols-hidden']),
      start: z.number().int().nonnegative().max(1_048_575),
      end: z.number().int().nonnegative().max(1_048_575),
      hidden: z.boolean(),
    })
    .strict()
    .refine((op) => op.end >= op.start && op.end - op.start < 100_000, {
      message: 'Invalid axis span.',
    }),
  z
    .object({
      sheetId: z.string().min(1),
      kind: z.enum(['set-rows-outline', 'set-cols-outline']),
      start: z.number().int().nonnegative().max(1_048_575),
      end: z.number().int().nonnegative().max(1_048_575),
      /// Absolute outline level; 0 removes the attribute. An omitted collapsed
      /// leaves the file's collapsed flag untouched.
      level: z.number().int().min(0).max(7),
      collapsed: z.boolean().optional(),
    })
    .strict()
    .refine((op) => op.end >= op.start && op.end - op.start < 100_000, {
      message: 'Invalid axis span.',
    }),
])

/// Declarative conditional-formatting snapshot for one sheet: the FULL rule
/// set in evaluation order (empty = remove all rules). Each `rule` carries
/// the Univer rule-config JSON; the gateway maps it strictly to OOXML and
/// fails the save on shapes it cannot represent.
export const workbookCfStateSchema = z
  .object({
    sheetId: z.string().min(1),
    rules: z
      .array(
        z
          .object({
            ranges: z.array(cellAreaSchema).min(1).max(100),
            stopIfTrue: z.boolean(),
            rule: z.record(z.string(), z.unknown()),
          })
          .strict(),
      )
      .max(500),
  })
  .strict()

/// Declarative data-validation snapshot for one sheet: the FULL rule set
/// (empty = remove all rules). Same recipe as conditional formatting — the
/// Univer rule JSON is the wire format, mapped strictly by the gateway.
export const workbookDvStateSchema = z
  .object({
    sheetId: z.string().min(1),
    rules: z
      .array(
        z
          .object({
            ranges: z.array(cellAreaSchema).min(1).max(100),
            rule: z.record(z.string(), z.unknown()),
          })
          .strict(),
      )
      .max(500),
  })
  .strict()

/// One printed header or footer: Excel's left/center/right sections, each
/// carrying field codes verbatim (&P page, &N pages, &D date, &T time,
/// &F file name, &A sheet name, && literal ampersand).
const headerFooterPartsSchema = z
  .object({
    left: z.string().max(255).optional(),
    center: z.string().max(255).optional(),
    right: z.string().max(255).optional(),
  })
  .strict()

/// Page Layout settings the user changed this session for one sheet; absent
/// fields stay verbatim in the file. Margin presets use the standard
/// normal/wide/narrow values.
export const workbookPageSetupStateSchema = z
  .object({
    sheetId: z.string().min(1),
    orientation: z.enum(['portrait', 'landscape']).optional(),
    /// OOXML paper-size code (1 = Letter, 9 = A4, …).
    paperSize: z.number().int().min(1).max(118).optional(),
    scale: z.number().int().min(10).max(400).optional(),
    fitToWidth: z.number().int().min(0).max(1_000).optional(),
    fitToHeight: z.number().int().min(0).max(1_000).optional(),
    fitToPage: z.boolean().optional(),
    margins: z.enum(['normal', 'wide', 'narrow']).optional(),
    printGridlines: z.boolean().optional(),
    printHeadings: z.boolean().optional(),
    showGridlines: z.boolean().optional(),
    showFormulas: z.boolean().optional(),
    printArea: z.union([z.string().min(1).max(255), z.null()]).optional(),
    printTitles: z.union([z.string().regex(/^\d{1,7}:\d{1,7}$/), z.null()]).optional(),
    /// Frozen pane counts; both present together, 0/0 removes the pane.
    frozenRows: z.number().int().min(0).max(1_048_575).optional(),
    frozenColumns: z.number().int().min(0).max(16_383).optional(),
    /// Printed header/footer sections, or null to clear that half.
    header: z.union([headerFooterPartsSchema, z.null()]).optional(),
    footer: z.union([headerFooterPartsSchema, z.null()]).optional(),
  })
  .strict()
  .refine((state) => Object.keys(state).length > 1, {
    message: 'A page-setup state needs at least one setting.',
  })

/// One hyperlink change at a cell, in final (post-operation) coordinates.
/// A '#Sheet!A1' target becomes an internal `location` anchor; anything else
/// is written as a TargetMode="External" relationship. null removes the link.
export const workbookHyperlinkEditSchema = z
  .object({
    sheetId: z.string().min(1),
    row: z.number().int().nonnegative().max(1_048_575),
    column: z.number().int().nonnegative().max(16_383),
    target: z.union([z.string().min(1).max(2083), z.null()]),
  })
  .strict()

export const workbookChartEditSchema = z
  .object({
    /// Constrained to the charts directory — the renderer chooses the path.
    chartPath: z.string().regex(/^xl\/charts\/[A-Za-z0-9._-]+\.xml$/),
    title: z.string().max(255).optional(),
    chartType: z.enum(['column', 'bar', 'line', 'area', 'pie', 'doughnut']).optional(),
    seriesColors: z.record(z.string().regex(/^[0-9]{1,3}$/), hexColorSchema).optional(),
    /// 'none' removes the legend; a side re-positions (creating it if needed).
    legend: z.enum(['none', 'right', 'bottom', 'top', 'left']).optional(),
    /// Plot-level data labels: values on bars/points, category+percent or
    /// percent on pie slices. 'none' removes them.
    dataLabels: z.enum(['none', 'value', 'percent', 'category-percent']).optional(),
    /// Placement and number format of the data labels (`c:dLblPos`/`c:numFmt`).
    dataLabelPosition: z.enum(['center', 'inside-end', 'outside-end']).optional(),
    dataLabelFormat: z.string().max(64).optional(),
    /// null removes that axis title. Axis-based charts only.
    axisTitles: z
      .object({
        category: z.string().max(255).nullable().optional(),
        value: z.string().max(255).nullable().optional(),
      })
      .strict()
      .optional(),
    /// Per-point fills (`c:dPt`), keyed series index → point index → color;
    /// how pie/doughnut slices get individual colors.
    pointColors: z
      .record(
        z.string().regex(/^[0-9]{1,3}$/),
        z.record(z.string().regex(/^[0-9]{1,3}$/), hexColorSchema),
      )
      .optional(),
    /// Bar/line/area stacking; 'clustered' means side-by-side (line/area
    /// write it as 'standard').
    grouping: z.enum(['clustered', 'stacked', 'percentStacked']).optional(),
    /// Value-axis major gridlines on/off (axis charts only).
    gridlines: z.boolean().optional(),
    /// Value-axis bounds; null resets that bound to auto.
    valueAxis: z
      .object({
        min: z.number().finite().nullable().optional(),
        max: z.number().finite().nullable().optional(),
      })
      .strict()
      .refine((axis) => axis.min !== undefined || axis.max !== undefined, {
        message: 'A value-axis edit needs min or max.',
      })
      .optional(),
    /// Bar family gap between categories, % of one bar width.
    gapWidthPct: z.number().int().min(0).max(500).optional(),
    /// Doughnut hole diameter, % of chart size.
    holeSizePct: z.number().int().min(10).max(90).optional(),
    /// Pie whole-ring explosion (series 0), % of radius.
    explosionPct: z.number().int().min(0).max(400).optional(),
    /// Pie per-slice explosion overrides (series 0), point index → %.
    pointExplosions: z
      .record(z.string().regex(/^[0-9]{1,3}$/), z.number().int().min(0).max(400))
      .optional(),
    /// Full series replacement (Select Data): existing series all drop and
    /// these are written in order. Wins over `series`/`seriesColors` edits.
    seriesSet: z
      .array(
        z
          .object({
            name: z.string().max(255),
            values: z.array(z.number().finite()).max(1_000),
            valuesRef: z.string().max(512).optional(),
            categories: z.array(z.string().max(255)).max(1_000).optional(),
            categoriesRef: z.string().max(512).optional(),
            color: hexColorSchema.optional(),
          })
          .strict(),
      )
      .min(1)
      .max(24)
      .optional(),
    /// Per-series rewrite of name and/or data (refs + caches travel together
    /// so the file and the on-screen render stay in sync).
    series: z
      .array(
        z
          .object({
            index: z.number().int().min(0).max(255),
            name: z.string().max(255).optional(),
            valuesRef: z.string().max(512).optional(),
            values: z.array(z.number().finite()).max(1_000).optional(),
            categoriesRef: z.string().max(512).optional(),
            categories: z.array(z.string().max(255)).max(1_000).optional(),
          })
          .strict()
          .refine(
            (entry) =>
              entry.name !== undefined ||
              entry.values !== undefined ||
              entry.categories !== undefined,
            { message: 'A series edit needs a name or data.' },
          ),
      )
      .max(24)
      .optional(),
  })
  .strict()
  .refine(
    (edit) =>
      edit.title !== undefined ||
      edit.chartType !== undefined ||
      (edit.seriesColors && Object.keys(edit.seriesColors).length > 0) ||
      (edit.pointColors && Object.keys(edit.pointColors).length > 0) ||
      edit.legend !== undefined ||
      edit.axisTitles !== undefined ||
      edit.dataLabels !== undefined ||
      edit.dataLabelPosition !== undefined ||
      edit.dataLabelFormat !== undefined ||
      edit.grouping !== undefined ||
      edit.gridlines !== undefined ||
      edit.valueAxis !== undefined ||
      edit.gapWidthPct !== undefined ||
      edit.holeSizePct !== undefined ||
      edit.explosionPct !== undefined ||
      (edit.pointExplosions && Object.keys(edit.pointExplosions).length > 0) ||
      (edit.seriesSet && edit.seriesSet.length > 0) ||
      (edit.series && edit.series.length > 0),
    { message: 'A chart edit needs at least one property.' },
  )

/// Edit to a visual that already lives in the file, located by the sidecar's
/// (drawingPath, anchor index) pair. `remove` deletes the anchor (charts
/// fail closed in the gateway); `anchor` rewrites its from/to markers.
export const workbookVisualEditSchema = z
  .object({
    drawingPath: z.string().regex(/^xl\/drawings\/[A-Za-z0-9._/-]+\.xml$/),
    drawingIndex: z.number().int().nonnegative().max(10_000),
    remove: z.literal(true).optional(),
    anchor: drawingAnchorSchema.optional(),
  })
  .strict()
  .refine((edit) => edit.remove === true || edit.anchor !== undefined, {
    message: 'A visual edit needs a removal or a new anchor.',
  })

/// OOXML customFilter comparison operators; absent means "equal". Wildcard
/// matching (contains / begins with) is encoded in the value string itself.
const customFilterOperatorSchema = z.enum([
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])

const filterColumnStateSchema = z
  .object({
    /// 0-based offset from the filter range's first column, per OOXML.
    colId: z.number().int().nonnegative().max(16_383),
    values: z.array(z.string().max(255)).max(10_000).optional(),
    blank: z.boolean().optional(),
    customs: z
      .object({
        and: z.boolean().optional(),
        filters: z
          .array(
            z
              .object({
                val: z.union([z.string().max(255), z.number().finite()]),
                operator: customFilterOperatorSchema.optional(),
              })
              .strict(),
          )
          .min(1)
          .max(2),
      })
      .strict()
      .optional(),
  })
  .strict()
  .refine(
    (column) =>
      column.values !== undefined || column.blank !== undefined || column.customs !== undefined,
    { message: 'A filter column needs values, a blank flag, or custom criteria.' },
  )

/// Declarative per-sheet filter snapshot taken at save time. `filter: null`
/// removes the sheet's autoFilter; `visibilityRange` rows not listed in
/// `hiddenRows` are unhidden (filtered-out rows and Excel's own filter
/// behavior treat hidden state declaratively within the range).
export const workbookFilterStateSchema = z
  .object({
    sheetId: z.string().min(1),
    filter: z
      .object({
        range: cellAreaSchema,
        columns: z.array(filterColumnStateSchema).max(1_000),
      })
      .strict()
      .nullable(),
    hiddenRows: z.array(z.number().int().nonnegative().max(1_048_575)).max(100_000),
    visibilityRange: cellAreaSchema,
  })
  .strict()

/// 1-31 characters, none of \ / ? * [ ] :, no leading/trailing apostrophe.
const sheetNameSchema = z
  .string()
  .min(1)
  .max(31)
  .regex(/^[^\\/?*[\]:]+$/)
  .refine((name) => !name.startsWith("'") && !name.endsWith("'"))

export const workbookSheetOpSchema = z.union([
  z
    .object({
      kind: z.literal('rename-sheet'),
      sheetId: z.string().min(1),
      newName: sheetNameSchema,
    })
    .strict(),
  z
    .object({
      kind: z.literal('add-sheet'),
      sheetId: z.string().min(1),
      name: sheetNameSchema,
    })
    .strict(),
  /// Like add-sheet, but the part is seeded by cloning sourceSheetId's part.
  z
    .object({
      kind: z.literal('duplicate-sheet'),
      sheetId: z.string().min(1),
      name: sheetNameSchema,
      sourceSheetId: z.string().min(1),
    })
    .strict(),
  z
    .object({
      kind: z.literal('remove-sheet'),
      sheetId: z.string().min(1),
    })
    .strict(),
  z
    .object({
      kind: z.literal('set-sheet-hidden'),
      sheetId: z.string().min(1),
      hidden: z.boolean(),
    })
    .strict(),
  /// Marker: only the tab order changed; the order itself rides sheetOrder.
  z.object({ kind: z.literal('reorder-sheets') }).strict(),
])

/// A visual created in the editor this session; the save writes charts as new
/// chart + drawing parts, shapes as anchors on the sheet's drawing part.
export const workbookVisualAddSchema = z
  .object({
    sheetId: z.string().min(1),
    anchor: drawingAnchorSchema,
    chart: z
      .object({
        chartType: z.enum([
          'column',
          'bar',
          'line',
          'area',
          'pie',
          'scatter',
          'radar',
          'doughnut',
          'combo',
        ]),
        title: z.string().max(255),
        series: z
          .array(
            z
              .object({
                name: z.string().max(255),
                categories: z.array(z.string().max(255)).max(1_000),
                values: z.array(z.number().finite()).max(1_000),
                valuesRef: z.string().max(512).optional(),
                categoriesRef: z.string().max(512).optional(),
                color: hexColorSchema.optional(),
                /// Per-point fills (`c:dPt`), keyed by 0-based point index.
                pointColors: z.record(z.string().regex(/^[0-9]{1,3}$/), hexColorSchema).optional(),
                explosionPct: z.number().int().min(0).max(400).optional(),
                pointExplosions: z
                  .record(z.string().regex(/^[0-9]{1,3}$/), z.number().int().min(0).max(400))
                  .optional(),
              })
              .strict(),
          )
          .min(1)
          .max(24),
        legend: z.enum(['none', 'right', 'bottom', 'top', 'left']).optional(),
        dataLabels: z.enum(['none', 'value', 'percent', 'category-percent']).optional(),
        dataLabelPosition: z.enum(['center', 'inside-end', 'outside-end']).optional(),
        dataLabelFormat: z.string().max(64).optional(),
        axisTitles: z
          .object({
            category: z.string().max(255).optional(),
            value: z.string().max(255).optional(),
          })
          .strict()
          .optional(),
        grouping: z.enum(['clustered', 'stacked', 'percentStacked', 'standard']).optional(),
        gridlines: z.boolean().optional(),
        valueAxis: z
          .object({
            min: z.number().finite().optional(),
            max: z.number().finite().optional(),
          })
          .strict()
          .optional(),
        gapWidthPct: z.number().int().min(0).max(500).optional(),
        holeSizePct: z.number().int().min(10).max(90).optional(),
      })
      .strict()
      .optional(),
    shape: z
      .object({
        shapeType: z.enum(ADDABLE_SHAPE_TYPES),
        fillColor: z
          .string()
          .regex(/^#[0-9a-fA-F]{6}$/)
          .optional(),
        text: z.string().max(1_000).optional(),
        isTextBox: z.boolean().optional(),
      })
      .strict()
      .optional(),
    image: z
      .object({
        mediaType: z.enum(['image/png', 'image/jpeg', 'image/gif']),
        /// ~20MB decoded
        base64: z.string().min(1).max(28_000_000),
      })
      .strict()
      .optional(),
  })
  .strict()
  .refine(
    (visual) =>
      [visual.chart, visual.shape, visual.image].filter((kind) => kind !== undefined).length === 1,
    { message: 'A visual addition carries exactly one of chart, shape, or image.' },
  )

/// Declarative note snapshot: the sheet's full comment set after the edit
/// (empty list removes all comments).
export const workbookNoteStateSchema = z
  .object({
    sheetId: z.string().min(1),
    notes: z
      .array(
        z
          .object({
            row: z.number().int().nonnegative().max(1_048_575),
            column: z.number().int().nonnegative().max(16_383),
            author: z.string().max(255),
            text: z.string().max(32_767),
          })
          .strict(),
      )
      .max(1_000),
  })
  .strict()

/// A table (ListObject) created in the editor this session; the save writes
/// a new xl/tables part and registers it on the worksheet. The header row is
/// always the area's first row (headerRowCount=1).
export const workbookTableAddSchema = z
  .object({
    sheetId: z.string().min(1),
    /// Table range in final (post-operation) coordinates, header row included.
    area: cellAreaSchema,
    /// Final unique table name (renderer-assigned when the AI op omits one).
    name: z.string().min(1).max(255),
    /// Header-row cell texts captured at apply time. The save reconciles them
    /// against journaled header-cell edits; blanks become Column1, Column2, …
    columnNames: z.array(z.string().max(255)).min(1).max(1_000),
    /// Built-in table style name; omitted = TableStyleMedium2.
    style: z
      .string()
      .regex(/^TableStyle(?:Light|Medium|Dark)[1-9][0-9]?$/)
      .optional(),
    bandedRows: z.boolean(),
  })
  .strict()

/// A PivotTable created in the editor this session. The aggregated grid is
/// already baked into cells (as ordinary cell edits); the save additionally
/// writes native pivot parts (cache with refreshOnLoad + table definition)
/// so Excel turns the baked grid into a live pivot on open.
export const workbookPivotAddSchema = z
  .object({
    /// Sheet that receives the pivot output.
    sheetId: z.string().min(1),
    sourceSheetId: z.string().min(1),
    /// Source data range, header row included (final coordinates).
    sourceArea: cellAreaSchema,
    /// Baked output area, headers and grand totals included.
    location: cellAreaSchema,
    name: z.string().min(1).max(255),
    /// All source field headers, in source-column order (cacheFields order).
    fieldNames: z.array(z.string().min(1).max(255)).min(1).max(200),
    /// Indices into fieldNames for the row dimension levels (outer → inner).
    rowFieldIndices: z.array(z.number().int().nonnegative()).min(1).max(8),
    columnFieldIndex: z.number().int().nonnegative().optional(),
    /// Indices into fieldNames for page (report-filter) fields.
    pageFieldIndices: z.array(z.number().int().nonnegative()).max(4).optional(),
    /// Distinct row/column item captions in baked (first-appearance) order.
    rowItems: z.array(z.string().max(255)).min(1).max(10_000),
    /// Deduplicated member lists per row level (required for multi-level rows;
    /// omissible for single level = [rowItems]).
    rowLevelItems: z
      .array(z.array(z.string().max(255)).max(10_000))
      .max(8)
      .optional(),
    /// Row-by-row layout of the output data rows (required for multi-level rows):
    /// data rows' members cover all levels; default (subtotal) rows only hold the
    /// fixed prefix levels.
    rowLines: z
      .array(
        z
          .object({
            t: z.enum(['data', 'default']),
            members: z.array(z.number().int().nonnegative()).min(1).max(8),
          })
          .strict(),
      )
      .max(20_000)
      .optional(),
    columnItems: z.array(z.string().max(255)).max(1_000).optional(),
    /// Multi-level columns: column dimension field indices (outer → inner). Takes
    /// precedence over columnFieldIndex when provided.
    columnFieldIndices: z.array(z.number().int().nonnegative()).min(1).max(8).optional(),
    /// Deduplicated member lists per column level (required with ≥2 column levels,
    /// symmetric with rowLevelItems).
    colLevelItems: z
      .array(z.array(z.string().max(255)).max(1_000))
      .max(8)
      .optional(),
    /// Column-by-column layout of the output data columns (excluding the trailing
    /// grand-total column; required with ≥2 column levels): data columns' members
    /// cover all levels; default (subtotal) columns only hold the fixed prefix
    /// levels.
    colLines: z
      .array(
        z
          .object({
            t: z.enum(['data', 'default']),
            members: z.array(z.number().int().nonnegative()).min(1).max(8),
          })
          .strict(),
      )
      .max(1_000)
      .optional(),
    /// Value/label filters: applied to row/column dimension fields, at most one
    /// per field.
    filters: z
      .array(
        z.discriminatedUnion('kind', [
          z
            .object({
              kind: z.literal('label'),
              field: z.number().int().nonnegative(),
              op: z.enum(['equal', 'contains', 'beginsWith']),
              value: z.string().max(255),
            })
            .strict(),
          z
            .object({
              kind: z.literal('value'),
              field: z.number().int().nonnegative(),
              /// The data field it measures (values array index).
              dataField: z.number().int().nonnegative(),
              op: z.enum(['top', 'greaterThan', 'between']),
              count: z.number().int().min(1).max(10_000).optional(),
              from: z.number().finite().optional(),
              to: z.number().finite().optional(),
            })
            .strict(),
        ]),
      )
      .max(16)
      .optional(),
    /// Filter-hidden member indices per row/column level (matching level items
    /// order).
    rowHiddenItems: z.array(z.array(z.number().int().nonnegative()).max(10_000)).max(8).optional(),
    colHiddenItems: z.array(z.array(z.number().int().nonnegative()).max(1_000)).max(8).optional(),
    /// Grouping rules for dimension fields (dates by year/quarter/month, numbers by
    /// fixed-step intervals).
    groupings: z
      .array(
        z.discriminatedUnion('kind', [
          z
            .object({
              kind: z.literal('date'),
              fieldIndex: z.number().int().nonnegative(),
              dateUnit: z.enum(['year', 'quarter', 'month']),
            })
            .strict(),
          z
            .object({
              kind: z.literal('range'),
              fieldIndex: z.number().int().nonnegative(),
              rangeStep: z.number().positive().finite(),
              rangeStart: z.number().finite().optional(),
            })
            .strict(),
        ]),
      )
      .max(8)
      .optional(),
    values: z
      .array(
        z
          .object({
            /// Source field index; -1 for calculated fields (formula present).
            fieldIndex: z.number().int().min(-1),
            agg: z.enum(['sum', 'count', 'average', 'max', 'min']),
            /// Optional Excel number format string, e.g. "#,##0.00"
            numFmt: z.string().min(1).max(255).optional(),
            /// "Show values as" mode (percentages); absent = plain aggregate value.
            showDataAs: z.enum(['percentOfTotal', 'percentOfRow', 'percentOfCol']).optional(),
            /// Calculated field: formula (referencing source field names) plus the new
            /// field name; agg is fixed to sum.
            formula: z.string().min(1).max(1_024).optional(),
            calcName: z.string().min(1).max(255).optional(),
          })
          .strict(),
      )
      .min(1)
      .max(8),
  })
  .strict()

export const workbookSaveRequestSchema = z
  .object({
    sessionId: z.string().uuid(),
    mode: z.enum(['save', 'save-as']),
    edits: z.array(workbookCellEditSchema).max(MAX_SAVE_EDITS),
    structuralOps: z.array(workbookStructuralOpSchema).max(1_000),
    chartEdits: z.array(workbookChartEditSchema).max(100),
    visualEdits: z.array(workbookVisualEditSchema).max(100),
    visualAdditions: z.array(workbookVisualAddSchema).max(50),
    tableAdditions: z.array(workbookTableAddSchema).max(50),
    pivotAdditions: z.array(workbookPivotAddSchema).max(20),
    sheetOps: z.array(workbookSheetOpSchema).max(100),
    /// Final tab order (Univer sheet ids); required with any sheet op.
    sheetOrder: z.array(z.string().min(1)).max(1_000),
    filterStates: z.array(workbookFilterStateSchema).max(1_000),
    hyperlinkEdits: z.array(workbookHyperlinkEditSchema).max(1_000),
    cfStates: z.array(workbookCfStateSchema).max(1_000),
    dvStates: z.array(workbookDvStateSchema).max(1_000),
    pageSetupStates: z.array(workbookPageSetupStateSchema).max(1_000),
    noteStates: z.array(workbookNoteStateSchema).max(1_000),
    /// Recalculated formula-cell values written back into <v> so the saved file's
    /// inputs and outputs agree for readers without a formula engine.
    formulaValues: z
      .array(
        z
          .object({
            sheetId: z.string().min(1),
            row: z.number().int().min(0),
            column: z.number().int().min(0),
            value: z.union([z.string().max(10_000), z.number(), z.boolean(), z.null()]),
          })
          .strict(),
      )
      .max(MAX_SAVE_EDITS),
    /// pivotCacheDefinition paths to flag refreshOnLoad (recomputed pivots).
    pivotCacheRefreshPaths: z.array(z.string().regex(/^xl\/[A-Za-z0-9._/-]+\.xml$/)).max(100),
    /// Pivot output-area expansion after layout growth: on save, the
    /// pivotTableDefinition's location ref is widened to newOutputRef (conflict
    /// checks fail closed in the gateway). When relayout (an A3 layout edit) is
    /// present, the whole definition/cacheDefinition is regenerated from the new
    /// layout (original name/cacheId/records rel kept; records are emptied and
    /// rebuilt via refreshOnLoad); its name is a placeholder — the saver uses the
    /// original name from the file.
    pivotRefreshUpdates: z
      .array(
        z
          .object({
            cachePath: z.string().regex(/^xl\/[A-Za-z0-9._/-]+\.xml$/),
            sheetId: z.string().min(1),
            newOutputRef: z.string().min(3).max(64),
            relayout: workbookPivotAddSchema.optional(),
          })
          .strict(),
      )
      .max(100),
    /// Desired worksheet-protection state per sheet (no password support).
    sheetProtections: z
      .array(
        z
          .object({
            sheetId: z.string().min(1),
            protected: z.boolean(),
          })
          .strict(),
      )
      .max(1_000),
    /// x14 sparkline groups created this session (defaulted so older callers
    /// keep working).
    sparklineAdditions: z
      .array(
        z
          .object({
            sheetId: z.string().min(1),
            type: z.enum(['line', 'column', 'stacked']),
            color: hexColorSchema.optional(),
            cells: z
              .array(
                z
                  .object({
                    cell: z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/),
                    /// Sheet-qualified source range, e.g. `Sheet1!A2:F2`.
                    sourceRef: z.string().min(1).max(512),
                  })
                  .strict(),
              )
              .min(1)
              .max(500),
          })
          .strict(),
      )
      .max(50)
      .default([]),
    /// Declarative defined-names snapshot (null = untouched). `preserveNames`
    /// lists file names the editor could not model — they stay verbatim, as do
    /// _xlnm.* built-ins and hidden names.
    definedNamesState: z
      .object({
        names: z
          .array(
            z
              .object({
                name: z.string().min(1).max(255),
                formula: z.string().min(1).max(8_192),
                sheetIndex: z.number().int().nonnegative().optional(),
              })
              .strict(),
          )
          .max(2_000),
        preserveNames: z.array(z.string().min(1).max(255)).max(2_000),
      })
      .strict()
      .nullable(),
  })
  .strict()
  .refine(
    (request) =>
      // Explicit Save As is a valid request even with nothing to apply: it
      // writes the unchanged workbook to a new path.
      request.mode === 'save-as' ||
      request.edits.length > 0 ||
      request.structuralOps.length > 0 ||
      request.chartEdits.length > 0 ||
      request.visualEdits.length > 0 ||
      request.sheetOps.length > 0 ||
      request.filterStates.length > 0 ||
      request.hyperlinkEdits.length > 0 ||
      request.cfStates.length > 0 ||
      request.dvStates.length > 0 ||
      request.pageSetupStates.length > 0 ||
      request.noteStates.length > 0 ||
      request.pivotCacheRefreshPaths.length > 0 ||
      request.pivotRefreshUpdates.length > 0 ||
      request.sheetProtections.length > 0 ||
      request.definedNamesState !== null ||
      request.visualAdditions.length > 0 ||
      request.tableAdditions.length > 0 ||
      request.pivotAdditions.length > 0 ||
      request.sparklineAdditions.length > 0,
    { message: 'A save needs at least one edit.' },
  )
  .refine((request) => request.sheetOps.length === 0 || request.sheetOrder.length > 0, {
    message: 'Sheet operations need the final sheet order.',
  })

export const workbookSaveResultSchema = z.union([
  z.object({ canceled: z.literal(true) }).strict(),
  z
    .object({
      canceled: z.literal(false),
      file: workbookFileSchema,
      touchedEntries: z.array(z.string()).max(MAX_SAVE_EDITS),
    })
    .strict(),
])

export const workbookMediaRequestSchema = z
  .object({
    sessionId: z.string().uuid(),
    visualId: z.string().min(1),
  })
  .strict()

/// Reads a user-named local image for AI-proposed insertion; the main
/// process verifies extension, magic bytes, and size before returning.
export const localImageRequestSchema = z
  .object({
    path: z.string().min(1).max(1024),
  })
  .strict()

export const localImageResultSchema = z
  .object({
    mediaType: z.enum(['image/png', 'image/jpeg', 'image/gif']),
    /// ~20MB decoded, mirroring the picker-based insert limit.
    base64: z.string().min(1).max(28_000_000),
  })
  .strict()

/// Insert → Screenshot: enumerate capturable windows/screens for the picker
/// grid, then grab the chosen source at full resolution. 'denied' means the
/// OS blocks capture (macOS Screen Recording permission) — the picker shows
/// guidance instead of an empty grid.
export const screenSourcesResultSchema = z
  .object({
    status: z.enum(['ok', 'denied']),
    sources: z.array(
      z
        .object({
          id: z.string().min(1),
          name: z.string(),
          kind: z.enum(['screen', 'window']),
          /// Small preview as a data URL; empty when the OS returned none.
          thumbnail: z.string(),
        })
        .strict(),
    ),
  })
  .strict()

export const screenCaptureRequestSchema = z
  .object({
    id: z.string().min(1).max(256),
  })
  .strict()

export const screenCaptureResultSchema = z
  .object({
    mediaType: z.literal('image/png'),
    base64: z.string().min(1).max(28_000_000),
    width: z.number().int().positive(),
    height: z.number().int().positive(),
  })
  .strict()

export const workbookPivotRequestSchema = z
  .object({
    sessionId: z.string().uuid(),
    /// Part paths straight from the sheet metadata's pivotTables entries.
    path: z.string().regex(/^xl\/[A-Za-z0-9._/-]+\.xml$/),
    cachePath: z.string().regex(/^xl\/[A-Za-z0-9._/-]+\.xml$/),
  })
  .strict()

const pivotLayoutLineSchema = z
  .object({
    t: z.string().max(32),
    members: z.array(z.number().int().nullable()).max(64),
    depth: z.number().int().nonnegative().max(64),
    dataField: z.number().int().nonnegative().max(255),
  })
  .strict()

/// Mirror of the gateway parser's PivotDefinition (validated across IPC).
export const workbookPivotDefinitionSchema = z
  .object({
    outputRef: z.string().min(1).max(64),
    firstDataRow: z.number().int().nonnegative().max(1_048_575),
    firstDataCol: z.number().int().nonnegative().max(16_383),
    fields: z
      .array(
        z
          .object({
            name: z.string().max(255),
            sharedItems: z
              .array(z.union([z.string().max(32_767), z.number(), z.boolean(), z.null()]))
              .max(100_000),
            /// Grouped field (date/numeric ranges): when present, sharedItems are group
            /// labels.
            grouping: z
              .discriminatedUnion('kind', [
                z
                  .object({
                    kind: z.literal('date'),
                    dateUnit: z.enum(['year', 'quarter', 'month']),
                  })
                  .strict(),
                z
                  .object({
                    kind: z.literal('range'),
                    rangeStep: z.number().positive().finite(),
                    rangeStart: z.number().finite().optional(),
                  })
                  .strict(),
              ])
              .optional(),
            /// Calculated field (cacheField@formula): not a source data column, takes no
            /// cache records.
            formula: z.string().max(2_048).optional(),
          })
          .strict(),
      )
      .max(1_000),
    fieldItems: z
      .array(
        z
          .array(
            z
              .object({
                x: z.number().int().nullable(),
                hidden: z.boolean(),
              })
              .strict(),
          )
          .max(100_000),
      )
      .max(1_000),
    rowFields: z.array(z.number().int()).max(64),
    colFields: z.array(z.number().int()).max(64),
    rowLines: z.array(pivotLayoutLineSchema).max(100_000),
    colLines: z.array(pivotLayoutLineSchema).max(10_000),
    dataFields: z
      .array(
        z
          .object({
            name: z.string().max(255),
            field: z.number().int().nonnegative().max(1_000),
            subtotal: z.string().max(32),
            /// "Show values as" mode (percentages); the parser only carries it over when
            /// supported.
            showDataAs: z.enum(['percentOfTotal', 'percentOfRow', 'percentOfCol']).optional(),
            /// Calculated-field formula (carried when fld points at a cacheField with a
            /// formula).
            formula: z.string().max(2_048).optional(),
          })
          .strict(),
      )
      .max(256),
    pageFields: z
      .array(
        z
          .object({
            field: z.number().int().nonnegative().max(1_000),
            item: z.number().int().nullable(),
          })
          .strict(),
      )
      .max(256),
    /// Value/label filters (pivotFilters); filtered-out members are hidden entries
    /// in fieldItems.
    filters: z
      .array(
        z.discriminatedUnion('kind', [
          z
            .object({
              kind: z.literal('label'),
              field: z.number().int().nonnegative().max(1_000),
              op: z.enum(['equal', 'contains', 'beginsWith']),
              value: z.string().max(255),
            })
            .strict(),
          z
            .object({
              kind: z.literal('value'),
              field: z.number().int().nonnegative().max(1_000),
              dataField: z.number().int().nonnegative().max(255),
              op: z.enum(['top', 'greaterThan', 'between']),
              count: z.number().int().min(1).optional(),
              from: z.number().optional(),
              to: z.number().optional(),
            })
            .strict(),
        ]),
      )
      .max(64),
    sourceSheet: z.string().max(255),
    sourceRef: z.string().max(64),
    unsupported: z.array(z.string().max(255)).max(100),
  })
  .strict()

export const workbookMediaResultSchema = z
  .object({
    mediaType: z.string().regex(/^image\//),
    base64: z.string().min(1),
  })
  .strict()

export type WorkbookFile = z.infer<typeof workbookFileSchema>
export type WorkbookStyleEdit = z.infer<typeof workbookStyleEditSchema>
export type WorkbookCellEdit = z.infer<typeof workbookCellEditSchema>
export type WorkbookStructuralOp = z.infer<typeof workbookStructuralOpSchema>
export type WorkbookChartEdit = z.infer<typeof workbookChartEditSchema>
export type WorkbookVisualEdit = z.infer<typeof workbookVisualEditSchema>
export type WorkbookHyperlinkEdit = z.infer<typeof workbookHyperlinkEditSchema>
export type WorkbookCfState = z.infer<typeof workbookCfStateSchema>
export type WorkbookDvState = z.infer<typeof workbookDvStateSchema>
export type WorkbookPageSetupState = z.infer<typeof workbookPageSetupStateSchema>
export type WorkbookNoteState = z.infer<typeof workbookNoteStateSchema>
export type WorkbookSheetOp = z.infer<typeof workbookSheetOpSchema>
export type WorkbookFilterState = z.infer<typeof workbookFilterStateSchema>
export type WorkbookSaveRequest = z.infer<typeof workbookSaveRequestSchema>
export type WorkbookSaveResult = z.infer<typeof workbookSaveResultSchema>
export type WorkbookRangeRequest = z.infer<typeof workbookRangeRequestSchema>
export type WorkbookFormulaCellsRequest = z.infer<typeof workbookFormulaCellsRequestSchema>
export type WorkbookFormulaCellsResult = z.infer<typeof workbookFormulaCellsResultSchema>
export type WorkbookRangeResult = z.infer<typeof workbookRangeResultSchema>
export type WorkbookRecalcRequest = z.infer<typeof workbookRecalcRequestSchema>
export type WorkbookRecalcResult = z.infer<typeof workbookRecalcResultSchema>
export type WorkbookMediaRequest = z.infer<typeof workbookMediaRequestSchema>
export type WorkbookMediaResult = z.infer<typeof workbookMediaResultSchema>
export type WorkbookPivotRequest = z.infer<typeof workbookPivotRequestSchema>
export type WorkbookPivotDefinition = z.infer<typeof workbookPivotDefinitionSchema>
export type LocalImageRequest = z.infer<typeof localImageRequestSchema>
export type LocalImageResult = z.infer<typeof localImageResultSchema>
export type ScreenSourcesResult = z.infer<typeof screenSourcesResultSchema>
export type ScreenCaptureRequest = z.infer<typeof screenCaptureRequestSchema>
export type ScreenCaptureResult = z.infer<typeof screenCaptureResultSchema>
export type WorkbookVisualObject = z.infer<typeof visualObjectSchema>
export type WorkbookVisualAdd = z.infer<typeof workbookVisualAddSchema>
export type WorkbookTableAdd = z.infer<typeof workbookTableAddSchema>
export type WorkbookPivotAdd = z.infer<typeof workbookPivotAddSchema>
export type WorkbookCellStyle = z.infer<typeof cellStyleSchema>
export type WorkbookRichRun = z.infer<typeof richRunSchema>
export type WorkbookConditionalRule = z.infer<typeof conditionalRuleSchema>

// ---- AI settings + chat/stream: canonical types live in @genoffice/ai-provider,
// shared with apps/docs. Validated here like every other renderer→main request in
// this file; the validated shape is cast to AiSettings at the main-process call
// site, which always has exactly the 5 known provider keys once merged through
// resolveAiSettings/defaultAiSettings. ----

const aiProviderConfigSchema = z
  .object({
    apiKey: z.string(),
    model: z.string(),
    baseUrl: z.string().optional(),
  })
  .strict()

export const aiSettingsInputSchema = z
  .object({
    provider: z.string().min(1),
    providers: z.record(z.string(), aiProviderConfigSchema),
  })
  .strict()

const agentToolResultSchema = z
  .object({
    id: z.string(),
    name: z.string(),
    output: z.string(),
    isError: z.boolean().optional(),
  })
  .strict()

const agentToolCallSchema = z
  .object({
    id: z.string(),
    name: z.string(),
    input: z.record(z.string(), z.unknown()),
  })
  .strict()

/// Inline vision input on a user turn (image attachments, base64 without data: prefix).
const agentImageSchema = z
  .object({
    base64: z.string().min(1),
    mime: z.string().min(1).max(64),
  })
  .strict()

const agentMessageSchema = z.union([
  z
    .object({
      role: z.literal('user'),
      text: z.string(),
      images: z.array(agentImageSchema).max(20).optional(),
    })
    .strict(),
  z
    .object({
      role: z.literal('assistant'),
      text: z.string(),
      toolCalls: z.array(agentToolCallSchema).optional(),
    })
    .strict(),
  z.object({ role: z.literal('tool'), results: z.array(agentToolResultSchema) }).strict(),
])

const agentToolDefSchema = z
  .object({
    name: z.string(),
    description: z.string(),
    inputSchema: z.record(z.string(), z.unknown()),
  })
  .strict()

const MAX_AI_MESSAGES = 500
const MAX_AI_TOOLS = 50

export const aiChatRequestSchema = z
  .object({
    settings: aiSettingsInputSchema,
    system: z.string(),
    user: z.string(),
  })
  .strict()

export const aiStreamRequestSchema = z
  .object({
    requestId: z.string().min(1),
    settings: aiSettingsInputSchema,
    system: z.string(),
    messages: z.array(agentMessageSchema).max(MAX_AI_MESSAGES),
    tools: z.array(agentToolDefSchema).max(MAX_AI_TOOLS).optional(),
    maxTokens: z.number().int().positive().optional(),
  })
  .strict()

export type AiSettingsInput = z.infer<typeof aiSettingsInputSchema>
export type AiChatRequestInput = z.infer<typeof aiChatRequestSchema>
export type AiStreamRequestInput = z.infer<typeof aiStreamRequestSchema>

/// A rendered print job: the renderer lays the sheet out as HTML, the main
/// process turns it into a PDF via a hidden window.
export const workbookExportPdfRequestSchema = z
  .object({
    fileName: z.string().min(1).max(255),
    html: z.string().min(1).max(20_000_000),
    landscape: z.boolean(),
    pageSize: z.union([
      z.enum(['A3', 'A4', 'A5', 'Legal', 'Letter', 'Tabloid']),
      /// Custom size in inches.
      z
        .object({
          width: z.number().positive().max(100),
          height: z.number().positive().max(100),
        })
        .strict(),
    ]),
    /// Inches.
    margins: z
      .object({
        top: z.number().min(0).max(3),
        bottom: z.number().min(0).max(3),
        left: z.number().min(0).max(3),
        right: z.number().min(0).max(3),
      })
      .strict(),
    scale: z.number().min(0.1).max(2),
  })
  .strict()

export const workbookExportPdfResultSchema = z.union([
  z.object({ canceled: z.literal(true) }).strict(),
  z.object({ canceled: z.literal(false), path: z.string().min(1) }).strict(),
])

export type WorkbookExportPdfRequest = z.infer<typeof workbookExportPdfRequestSchema>
export type WorkbookExportPdfResult = z.infer<typeof workbookExportPdfResultSchema>

// ---- Chat attachments (local files fed to the agent via tools; same structure
// as apps/docs and apps/slides) ----

/** Image attachment extensions: no text extraction; read as base64 on send and
 * passed to the model as a multimodal image with the user message */
export const ATTACHMENT_IMAGE_EXTS = new Set(['png', 'jpg', 'jpeg', 'gif', 'webp'])

export interface AttachmentMeta {
  /** Absolute local path; the file never leaves the machine */
  path: string
  name: string
  /** Lowercase extension, no dot */
  ext: string
  sizeBytes: number
}

export interface AttachmentAddResult {
  accepted: AttachmentMeta[]
  /** Per-file rejection reason (too large/unsupported type/unreadable) */
  rejected: string[]
}

export interface AttachmentReadResult {
  ok: boolean
  error?: string
  name?: string
  /** Total character count of the extracted text */
  totalChars?: number
  /** The requested chunk */
  text?: string
  offset?: number
}

/** Raw bytes of an image attachment (multimodal input) */
export interface AttachmentImageResult {
  ok: boolean
  /** raw base64 (no data: prefix) */
  base64?: string
  mime?: string
  error?: string
}

export type UiTheme = 'light' | 'dark' | 'system'

export interface DesktopApi {
  /** current UI language (persisted by the shell in app-settings.json) */
  getLanguage(): Promise<'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar'>
  /** language switched from the shell home page */
  onLanguageChanged(
    handler: (
      lang: 'zh' | 'en' | 'ja' | 'ko' | 'fr' | 'de' | 'es' | 'th' | 'id' | 'ru' | 'ar',
    ) => void,
  ): () => void
  /** current UI theme preference (persisted by the shell in app-settings.json) */
  getTheme(): Promise<UiTheme>
  /** theme switched from the shell home page */
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  selectWorkbook(): Promise<WorkbookFile | null>
  readWorkbookRange(request: WorkbookRangeRequest): Promise<WorkbookRangeResult>
  readWorkbookFormulas(request: WorkbookFormulaCellsRequest): Promise<WorkbookFormulaCellsResult>
  recalcWorkbook(request: WorkbookRecalcRequest): Promise<WorkbookRecalcResult>
  readWorkbookMedia(request: WorkbookMediaRequest): Promise<WorkbookMediaResult>
  readPivotDefinition(request: WorkbookPivotRequest): Promise<WorkbookPivotDefinition>
  readLocalImage(request: LocalImageRequest): Promise<LocalImageResult>
  captureScreenSources(): Promise<ScreenSourcesResult>
  /// null when the source vanished between listing and capture.
  captureScreenSource(request: ScreenCaptureRequest): Promise<ScreenCaptureResult | null>
  saveWorkbookEdits(request: WorkbookSaveRequest): Promise<WorkbookSaveResult>
  /// Crash-recovery copy of the pending edits, written under userData.
  /// Best-effort: never prompts, never touches the opened file.
  writeWorkbookRecovery(request: WorkbookSaveRequest): Promise<{ ok: boolean }>
  /// Rename a still-untitled workbook after AI-generated content (no-op unless
  /// the file still carries the shell's auto-created untitled name).
  autoRenameWorkbook(
    sessionId: string,
    baseName: string,
  ): Promise<{ renamed: boolean; name?: string }>
  exportPdf(request: WorkbookExportPdfRequest): Promise<WorkbookExportPdfResult>
  closeWorkbook(sessionId: string): Promise<void>
  openExternal(url: string): Promise<void>
  /// Application-menu File commands (Open/Save/Save As); returns unsubscribe.
  onMenuAction(callback: (action: MenuAction) => void): () => void
  /// The open workbook was renamed on disk (renamed in the shell Home list);
  /// emits the new file name.
  onWorkbookRenamed(callback: (newName: string) => void): () => void
  /// Mirrors the pending-edit badge to the main process for the close guard.
  notifyPendingEdits(count: number): void
  /// Main asks the renderer to save before closing; reply via reportCloseSaveResult.
  onCloseSaveRequest(callback: () => void): () => void
  reportCloseSaveResult(ok: boolean): void
  /// Returns true once when this tab was opened via "New Spreadsheet" from the
  /// shell home.
  consumeNewBlankWorkbook(): Promise<boolean>
  /// Is a shell-queued workbook path still waiting to be opened? (The shell's
  /// 'open' nudge loop can time out on slow cold starts; the renderer pulls.)
  hasQueuedWorkbook(): Promise<boolean>
  getAiSettings(): Promise<AiSettings>
  setAiSettings(settings: AiSettings): Promise<void>
  aiChat(request: AiChatRequest): Promise<AiChatResponse>
  /// start a streaming AI call; deltas arrive via onAiStream with the same requestId
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  /// Genspark account status (gsk login state); withEmail also returns the email
  /// (needs a network request, slower)
  aiGskStatus(withEmail?: boolean): Promise<GenSparkAccountStatus>
  /// Opens the browser to sign in to Genspark (fire-and-forget; aiGskStatus
  /// becomes signed-in on completion)
  aiGskLogin(): Promise<void>
  /// Web search (main-process Serper/DuckDuckGo, shared with docs/slides)
  webSearch(query: string, maxResults?: number): Promise<WebSearchResult>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
  /// Chat attachments: multi-select file dialog (returns null on cancel)
  pickAttachments(): Promise<AttachmentAddResult | null>
  /// Validates dropped paths and returns attachment metadata
  addAttachmentPaths(paths: string[]): Promise<AttachmentAddResult>
  /// Persists a clipboard-pasted image (no local path) to a temp file and adds it
  /// as an attachment
  addPastedImage(data: ArrayBuffer, ext: string): Promise<AttachmentAddResult>
  /// Reads one chunk of an attachment's extracted text
  readAttachment(path: string, offset: number, maxChars: number): Promise<AttachmentReadResult>
  /// Reads an image attachment as base64 for multimodal input (≤5MB)
  readAttachmentImage(path: string): Promise<AttachmentImageResult>
  /// Absolute path of a File dropped onto the window (Electron webUtils)
  getPathForFile(file: File): string
}

export type MenuAction = 'open' | 'save' | 'save-as' | 'export-pdf' | 'undo' | 'redo'

export interface WebSearchResult {
  results: Array<{ title: string; url: string; snippet: string }>
  answer?: string
  method: string
}
