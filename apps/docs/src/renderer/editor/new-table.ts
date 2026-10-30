import type {
  CellBorder,
  TableBorders,
  TableCell,
  TableModel,
  TableParagraph,
} from '@genoffice/docx-engine'

// Word's Insert Table look (Table Grid style): single 0.5pt borders on every
// edge and cell paragraphs at single spacing with no space after, so a one-line
// cell hugs its text instead of inheriting Normal's 8pt/1.08 gap. Cells start
// vertically centered (a product choice; Word starts at top).
const LINE: CellBorder = { style: 'single', szEighths: 4, color: 'auto' }

export const WORD_TABLE_BORDERS: TableBorders = {
  top: LINE,
  bottom: LINE,
  left: LINE,
  right: LINE,
  insideH: LINE,
  insideV: LINE,
}

export function wordCellParagraph(
  text: string,
  run?: { bold?: boolean; color?: string },
): TableParagraph {
  return {
    spaceAfter: 0,
    lineSpacing: 1,
    lineRule: 'auto',
    runs: text === '' ? [] : [{ text, ...run }],
  }
}

export function wordTableCell(paras: string[], style?: Partial<TableCell>): TableCell {
  const texts = paras.length > 0 ? paras : ['']
  return {
    vAlign: 'center',
    ...style,
    paras: texts,
    richParas: texts.map((text) => wordCellParagraph(text, { bold: style?.bold })),
  }
}

/** a freshly inserted table laid out like Word's: full-width autofit grid with equal columns */
export function wordTableModel(rows: TableCell[][], cols: number): TableModel {
  return {
    rows,
    colWidthsPct: Array.from({ length: cols }, () => 100 / cols),
    widthPct: 100,
    autoFit: 'window',
    borders: WORD_TABLE_BORDERS,
  }
}
