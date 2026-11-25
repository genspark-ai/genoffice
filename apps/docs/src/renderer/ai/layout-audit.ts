/**
 * Layout check without a picture of the page: says in the tool result what the
 * restricted-HTML parser dropped or the page cannot fit, so the model can
 * correct it instead of reporting a layout the reader does not see.
 */
import type { AiSectionState } from './page-setup'

/** A4 with 2.54 cm margins: the text width assumed for a document that does not exist yet */
export const DEFAULT_TEXT_WIDTH_TWIPS = 9026

const SIDE_BY_SIDE_CSS =
  /display\s*:\s*(?:inline-)?(?:flex|grid)|display\s*:\s*inline-block|column-count\s*:|columns\s*:|float\s*:\s*(?:left|right)/i

const CELL_BLOCK_TAGS = 'ul, ol, h1, h2, h3, h4, h5, h6, pre, blockquote'

const LATIN_EM = 0.55
const BODY_PT = 11
const CELL_PADDING_PT = 11
const NARROW_COLUMN_PT = 40

const CJK = /[\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/

export function textColumnWidthTwips(section: AiSectionState): number {
  const { width, margins, columns } = section
  const text = width - margins.left - margins.right - margins.gutter
  const count = Math.max(1, columns.count)
  return Math.max(0, (text - (count - 1) * columns.spacing) / count)
}

export function sectionOfBlock(
  sections: readonly AiSectionState[],
  blockIndex: number,
): AiSectionState | undefined {
  return (
    sections.find((s) => blockIndex >= s.firstBlock && blockIndex <= s.lastBlock) ?? sections[0]
  )
}

function parse(raw: string): Document | null {
  const fence = /```(?:html)?\s*([\s\S]*?)```/.exec(raw.trim())
  const text = fence ? fence[1] : raw
  if (!/<[a-z][\s\S]*>/i.test(text)) return null
  return new DOMParser().parseFromString(text, 'text/html')
}

function cm(pt: number): string {
  return `${(pt / 28.35).toFixed(1)} cm`
}

function widestToken(text: string): { token: string; widthPt: number } {
  let best = { token: '', widthPt: 0 }
  for (const word of text.split(/[\s\u00ad\u200b-]+/)) {
    // a CJK character is its own break opportunity
    for (const token of word.split(CJK)) {
      const widthPt = token.length * LATIN_EM * BODY_PT
      if (widthPt > best.widthPt) best = { token, widthPt }
    }
  }
  return best
}

/** a printed PDF keeps nested tables and cell markup, and has no page setup to change afterwards */
export type LayoutTarget = 'docx' | 'pdf'

function sideBySideFindings(doc: Document, target: LayoutTarget): string[] {
  const styled = Array.from(doc.querySelectorAll('[style]')).some((el) =>
    SIDE_BY_SIDE_CSS.test(el.getAttribute('style') ?? ''),
  )
  const sheet = Array.from(doc.querySelectorAll('style')).some((el) =>
    SIDE_BY_SIDE_CSS.test(el.textContent ?? ''),
  )
  if (!styled && !sheet) return []
  const lost =
    'The HTML places content side by side with CSS (flex / grid / columns / float). CSS is not read, so that content now runs in ONE column, top to bottom.'
  return [
    target === 'pdf'
      ? `${lost} To put items next to each other use a table with one row.`
      : `${lost} For text flowing in columns use set_page_setup columns (insert_section_break continuous limits it to a part of the document); to put items next to each other use a table with one row, whose cells hold plain text lines only.`,
  ]
}

function tableStructureFindings(doc: Document, target: LayoutTarget): string[] {
  const out: string[] = []
  const tables = Array.from(doc.querySelectorAll('table')).filter(
    (table) => !table.parentElement?.closest('table'),
  )
  tables.forEach((table, i) => {
    const label = tables.length > 1 ? `Table ${i + 1}` : 'The table'
    const merged = Array.from(table.querySelectorAll('td, th')).filter(
      (cell) =>
        Number(cell.getAttribute('colspan') ?? 1) > 1 ||
        Number(cell.getAttribute('rowspan') ?? 1) > 1,
    ).length
    if (merged > 0) {
      out.push(
        `${label} uses colspan/rowspan on ${merged} cell(s). Merged cells are not supported: each of them became a single cell and the cells after it moved out of their columns. Rewrite the table with exactly one cell per column in every row.`,
      )
    }
    if (target === 'pdf') return
    if (table.querySelector('table')) {
      out.push(
        `${label} has a table inside a cell. Nested tables are not supported: the inner table was reduced to plain text in that cell.`,
      )
    }
    const structured = Array.from(table.querySelectorAll('td, th')).filter((cell) =>
      cell.querySelector(CELL_BLOCK_TAGS),
    ).length
    if (structured > 0) {
      out.push(
        `${label} has lists, headings or code blocks inside ${structured} cell(s). Cells hold plain text only: that markup was reduced to plain lines without bullets, numbers or heading styles.`,
      )
    }
  })
  return out
}

function tableFitFindings(doc: Document, textWidthTwips: number): string[] {
  const out: string[] = []
  const textPt = textWidthTwips / 20
  if (textPt <= 0) return out
  const tables = Array.from(doc.querySelectorAll('table')).filter(
    (table) => !table.parentElement?.closest('table'),
  )
  tables.forEach((table, i) => {
    const label = tables.length > 1 ? `Table ${i + 1}` : 'The table'
    const rows = Array.from(table.querySelectorAll('tr'))
      .filter((tr) => tr.closest('table') === table)
      .map((tr) => Array.from(tr.children).filter((c) => /^t[hd]$/i.test(c.tagName)))
    const cols = Math.max(0, ...rows.map((r) => r.length))
    if (cols === 0) return
    const columnPt = textPt / cols
    const roomPt = columnPt - CELL_PADDING_PT
    if (columnPt < NARROW_COLUMN_PT) {
      out.push(
        `${label} has ${cols} columns, which leaves about ${cm(columnPt)} per column on this page (text width ${cm(textPt)}): the text will break after every few characters into very tall rows. Use fewer columns, split it into two tables, or switch the section to landscape with set_page_setup.`,
      )
      return
    }
    let worst: { column: number; token: string; widthPt: number } | null = null
    for (const row of rows) {
      row.forEach((cell, column) => {
        const { token, widthPt } = widestToken(cell.textContent ?? '')
        if (widthPt > roomPt && (!worst || widthPt > worst.widthPt))
          worst = { column, token, widthPt }
      })
    }
    if (worst) {
      const { column, token, widthPt } = worst as { column: number; token: string; widthPt: number }
      const shown = token.length > 24 ? `${token.slice(0, 24)}...` : token
      out.push(
        `${label}: column ${column + 1} holds an unbreakable string of ${token.length} characters ("${shown}", about ${cm(widthPt)}) in a column about ${cm(columnPt)} wide. It will be broken mid-word or pushed past the cell edge. Shorten it, add spaces or hyphens where it may break, or use fewer columns.`,
      )
    }
  })
  return out
}

/** `textWidthTwips`: the text column the content lands in; without it only structure is checked */
export function auditHtmlLayout(
  raw: string,
  textWidthTwips?: number,
  target: LayoutTarget = 'docx',
): string[] {
  const doc = parse(raw)
  if (!doc) return []
  return [
    ...sideBySideFindings(doc, target),
    ...tableStructureFindings(doc, target),
    ...(textWidthTwips && target === 'docx' ? tableFitFindings(doc, textWidthTwips) : []),
  ]
}

export function layoutAuditNote(
  raw: string,
  textWidthTwips?: number,
  target: LayoutTarget = 'docx',
): string {
  const findings = auditHtmlLayout(raw, textWidthTwips, target)
  if (findings.length === 0) return ''
  return `\nLayout check: the page differs from what the HTML asked for. Fix these before you answer, or tell the user plainly:\n${findings.map((f) => `- ${f}`).join('\n')}`
}
