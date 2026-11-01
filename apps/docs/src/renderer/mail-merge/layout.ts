import type { SectionSettings } from '@genoffice/docx-engine'
import type { PmNode } from '../editor/convert'

/**
 * The paper a label sheet or an envelope prints on, and the label grid on it.
 *
 * Labels ride one continuous table — a cell per record, rows flowing onto as
 * many pages as the records take — the shape Word's label output has, so the
 * table's own pagination does the page breaks. Envelopes keep the per-record
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
  { id: 'size10', width: twipsOfInches(9.5), height: twipsOfInches(4.125) },
  { id: 'dl', width: twipsOfMm(220), height: twipsOfMm(110) },
  { id: 'c5', width: twipsOfMm(229), height: twipsOfMm(162) },
  { id: 'c6', width: twipsOfMm(162), height: twipsOfMm(114) },
] as const

export type EnvelopePresetId = (typeof ENVELOPE_PRESETS)[number]['id']

/** the envelope paper for a preset, with a plain 0.5" margin all round */
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
 * fills it; the stock presets cover the exact layouts.
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
    grid: { columns, rows, gapX: 0, gapY: 0 },
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

/** the label cell's own size, twips: the grid splits the usable area */
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

/**
 * The label sheet as a ProseMirror document: one continuous table whose cells
 * carry the records in order, row-major, padded with empty cells.
 *
 * `cells` is one merged cell content per record (see cellContentOf); the
 * grid's column width rides the first row's cells in px (1px = 15 twips).
 */
export function labelSheetDoc(cells: readonly PmNode[][], layout: MergePageLayout): PmNode {
  const grid = layout.grid!
  const size = labelCellSize(layout)
  const colwidth = [Math.max(1, Math.round(size.width / 15))]
  const rows: PmNode[] = []
  for (let first = 0; first < cells.length; first += grid.columns) {
    const content: PmNode[] = []
    for (let column = 0; column < grid.columns; column++) {
      const record = cells[first + column]
      content.push({
        type: 'docTableCell',
        attrs: { colspan: 1, rowspan: 1, colwidth },
        content: record ?? [emptyParagraph()],
      })
    }
    rows.push({
      type: 'docTableRow',
      attrs: { heightTwips: size.height, heightRule: 'exact' },
      content,
    })
  }
  return {
    type: 'doc',
    content: [
      {
        type: 'docTable',
        attrs: { docxIndex: null },
        content: rows,
      },
    ],
  }
}
