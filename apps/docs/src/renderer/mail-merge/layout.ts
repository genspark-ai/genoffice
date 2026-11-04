import type { CellBorders, SectionSettings } from '@genoffice/docx-engine'
import type { PmNode } from '../editor/convert'

/**
 * The paper a label sheet or an envelope prints on, and the label grid on it.
 *
 * Labels ride one continuous table — a cell per record, rows flowing onto as
 * many pages as the records take — the shape Word's label output has, so the
 * table's own pagination does the page breaks. The table carries the stock's
 * empty gutter cells between the label columns (Word's 5160 table is
 * 3780/180/3780/180/3780), so the labels sit where the stock has them even
 * though the table spans the whole text column. Envelopes keep the per-record
 * documents the letter mode already builds, with the section's paper swapped
 * for the envelope. genoffice#2097.
 */

/** the paper the merge output prints on, and the label grid on it */
export interface MergePageLayout {
  /** the paper, twips (a label stock's sheet, or the envelope itself) */
  pageWidth: number
  pageHeight: number
  /** the page margins, twips */
  marginTop: number
  marginRight: number
  marginBottom: number
  marginLeft: number
  /** the label grid on the paper; absent = one full-page cell per record */
  grid?: {
    columns: number
    rows: number
    /** the gaps between labels, twips */
    gapX: number
    gapY: number
  }
}

export const TWIPS_PER_INCH = 1440

export const twipsOfInches = (inches: number): number => Math.round(inches * TWIPS_PER_INCH)

export const twipsOfMm = (mm: number): number => Math.round((mm * TWIPS_PER_INCH) / 25.4)

/** the editor's table columns are px: 1px = 15 twips */
const TWIPS_PER_PX = 15

/**
 * A label sheet has no printed grid: the gridlines are view-only. The save path
 * honours per-cell borders, so every cell writes all four sides as "none" (the
 * table-level default single lines are overridden by the cells).
 */
const BORDERLESS_CELLS: CellBorders = {
  top: { style: 'none' },
  left: { style: 'none' },
  bottom: { style: 'none' },
  right: { style: 'none' },
}

/** the most columns/rows a custom grid offers: 500 one-pixel columns are not a label stock */
export const MAX_CUSTOM_GRID = 20

/** clamp a custom grid count into [1, MAX_CUSTOM_GRID]; a non-finite typed value falls back */
export function clampGridCount(value: number, fallback = 1): number {
  const count = Math.trunc(value)
  return Number.isFinite(count) ? Math.min(MAX_CUSTOM_GRID, Math.max(1, count)) : fallback
}

/**
 * The Avery 5160 sheet: 30 labels of 2.625" × 1" on US Letter — 0.1875" side
 * margins, 0.125" horizontal gaps, 0.5" top and bottom, no vertical gaps.
 * The geometry sums to the paper: 2 × 0.1875" + 3 × 2.625" + 2 × 0.125" = 8.5",
 * 2 × 0.5" + 10 × 1" = 11".
 */
export const AVERY_5160: MergePageLayout = {
  pageWidth: twipsOfInches(8.5),
  pageHeight: twipsOfInches(11),
  marginTop: twipsOfInches(0.5),
  marginRight: twipsOfInches(0.1875),
  marginBottom: twipsOfInches(0.5),
  marginLeft: twipsOfInches(0.1875),
  grid: {
    columns: 3,
    rows: 10,
    gapX: twipsOfInches(0.125),
    gapY: 0,
  },
}

/** the envelope sizes the merge offers, as the paper the record prints on */
export const ENVELOPE_PRESETS = [
  { id: 'size10', name: 'Size 10', width: twipsOfInches(9.5), height: twipsOfInches(4.125) },
  { id: 'dl', name: 'DL', width: twipsOfMm(220), height: twipsOfMm(110) },
  { id: 'c5', name: 'C5', width: twipsOfMm(229), height: twipsOfMm(162) },
  { id: 'c6', name: 'C6', width: twipsOfMm(162), height: twipsOfMm(114) },
] as const

export type EnvelopePresetId = (typeof ENVELOPE_PRESETS)[number]['id']

/**
 * The envelope paper for a preset, with a plain 0.5" margin all round.
 *
 * The document itself is printed as it is: its text starts at the top-left
 * margin, where Word's Envelope wizard puts the return address — the wizard's
 * centred delivery-address frame is not reproduced here, and the panel says so.
 */
export function envelopeLayout(id: EnvelopePresetId): MergePageLayout {
  const preset = ENVELOPE_PRESETS.find((entry) => entry.id === id) ?? ENVELOPE_PRESETS[0]!
  return {
    pageWidth: preset.width,
    pageHeight: preset.height,
    marginTop: TWIPS_PER_INCH / 2,
    marginRight: TWIPS_PER_INCH / 2,
    marginBottom: TWIPS_PER_INCH / 2,
    marginLeft: TWIPS_PER_INCH / 2,
  }
}

/**
 * The custom label grid: an even grid on Letter or A4 with plain half-inch
 * margins, the cells splitting the usable area. Any sheet of near-even labels
 * fills it; the stock presets cover the exact layouts. The column/row count is
 * clamped so a typed value cannot ask for hundreds of one-pixel cells.
 */
export function customGridLayout(
  paper: 'letter' | 'a4',
  columns: number,
  rows: number,
): MergePageLayout {
  const pageWidth = paper === 'a4' ? twipsOfMm(210) : twipsOfInches(8.5)
  const pageHeight = paper === 'a4' ? twipsOfMm(297) : twipsOfInches(11)
  return {
    pageWidth,
    pageHeight,
    marginTop: TWIPS_PER_INCH / 2,
    marginRight: TWIPS_PER_INCH / 2,
    marginBottom: TWIPS_PER_INCH / 2,
    marginLeft: TWIPS_PER_INCH / 2,
    grid: { columns: clampGridCount(columns), rows: clampGridCount(rows), gapX: 0, gapY: 0 },
  }
}

/** the section the output prints with: the layout's paper and margins */
export function layoutSection(
  layout: MergePageLayout,
  base: SectionSettings | null,
): SectionSettings {
  return {
    ...(base ?? ({ pageBorder: false, orientation: 'portrait' } as SectionSettings)),
    pageWidth: layout.pageWidth,
    pageHeight: layout.pageHeight,
    orientation: layout.pageWidth > layout.pageHeight ? 'landscape' : 'portrait',
    marginTop: layout.marginTop,
    marginRight: layout.marginRight,
    marginBottom: layout.marginBottom,
    marginLeft: layout.marginLeft,
  }
}

/**
 * The label cell's own size, twips: the grid's label columns plus the gutters
 * split the usable area (the table's spacer cells emit the gutters).
 */
export function labelCellSize(layout: MergePageLayout): { width: number; height: number } {
  const grid = layout.grid!
  const usableWidth =
    layout.pageWidth - layout.marginLeft - layout.marginRight - grid.gapX * (grid.columns - 1)
  const usableHeight =
    layout.pageHeight - layout.marginTop - layout.marginBottom - grid.gapY * (grid.rows - 1)
  return {
    width: Math.floor(usableWidth / grid.columns),
    height: Math.floor(usableHeight / grid.rows),
  }
}

export function emptyParagraph(): PmNode {
  return { type: 'docParagraph', content: [] }
}

/**
 * The paragraph-shaped blocks one label cell holds, in the node types a table
 * cell takes. Headings go in as paragraphs with their formatting; the other
 * block kinds (a table, a picture) have no place in a label cell, so a
 * template built for labels keeps to paragraphs.
 */
export function cellContentOf(doc: PmNode): PmNode[] {
  const out: PmNode[] = []
  for (const block of doc.content ?? []) {
    if (block.type === 'docParagraph' || block.type === 'docListItem') out.push(block)
    else if (block.type === 'docHeading') out.push({ ...block, type: 'docParagraph' })
  }
  return out.length > 0 ? out : [emptyParagraph()]
}

/** how many template blocks a label cell cannot carry (tables, pictures, protected blocks) */
export function unsupportedLabelBlocks(doc: PmNode): number {
  return (doc.content ?? []).filter(
    (block) =>
      block.type !== 'docParagraph' && block.type !== 'docListItem' && block.type !== 'docHeading',
  ).length
}

/** the table's column count: the label columns plus the gutter columns between them */
function tableColumnCount(grid: NonNullable<MergePageLayout['grid']>): number {
  return grid.gapX > 0 ? grid.columns * 2 - 1 : grid.columns
}

/** a label cell width in the table's px unit */
const cellWidthPx = (layout: MergePageLayout): number =>
  Math.max(1, Math.round(labelCellSize(layout).width / TWIPS_PER_PX))

/** a gutter cell width in the table's px unit */
const gutterWidthPx = (twips: number): number => Math.max(1, Math.round(twips / TWIPS_PER_PX))

/**
 * The label sheet as a ProseMirror document: one continuous table whose cells
 * carry the records in order, row-major, padded with empty cells. The stock's
 * gutters become empty spacer cells/rows so the labels sit on the paper where
 * the stock has them; the table is borderless and a trailing empty paragraph
 * follows it (same reason as generateTableXml: Word wants a paragraph after a
 * table before the section properties).
 *
 * `cells` is one merged cell content per record (see cellContentOf); the
 * grid's column width rides the first row's cells in px (1px = 15 twips).
 */
export function labelSheetDoc(cells: readonly PmNode[][], layout: MergePageLayout): PmNode {
  const grid = layout.grid!
  const size = labelCellSize(layout)
  const labelWidth = cellWidthPx(layout)
  const gutterWidth = grid.gapX > 0 ? gutterWidthPx(grid.gapX) : 0
  const columns = tableColumnCount(grid)

  const labelCell = (cellContent: PmNode[] | undefined): PmNode => ({
    type: 'docTableCell',
    attrs: { colspan: 1, rowspan: 1, colwidth: [labelWidth], borders: BORDERLESS_CELLS },
    content: cellContent ?? [emptyParagraph()],
  })
  const gutterCell = (): PmNode => ({
    type: 'docTableCell',
    attrs: { colspan: 1, rowspan: 1, colwidth: [gutterWidth], borders: BORDERLESS_CELLS },
    content: [emptyParagraph()],
  })
  // a vertical gutter spans the whole grid, so it is one borderless cell
  const gutterRow = (): PmNode => ({
    type: 'docTableRow',
    attrs: { heightTwips: grid.gapY, heightRule: 'exact' },
    content: [
      {
        type: 'docTableCell',
        attrs: { colspan: columns, rowspan: 1, borders: BORDERLESS_CELLS },
        content: [emptyParagraph()],
      },
    ],
  })

  const rows: PmNode[] = []
  for (let first = 0; first < cells.length; first += grid.columns) {
    const content: PmNode[] = []
    for (let column = 0; column < grid.columns; column++) {
      content.push(labelCell(cells[first + column]))
      if (gutterWidth > 0 && column < grid.columns - 1) content.push(gutterCell())
    }
    rows.push({
      type: 'docTableRow',
      attrs: { heightTwips: size.height, heightRule: 'exact' },
      content,
    })
    if (grid.gapY > 0 && first + grid.columns < cells.length) rows.push(gutterRow())
  }
  return {
    type: 'doc',
    content: [
      {
        type: 'docTable',
        attrs: { docxIndex: null },
        content: rows,
      },
      emptyParagraph(),
    ],
  }
}
