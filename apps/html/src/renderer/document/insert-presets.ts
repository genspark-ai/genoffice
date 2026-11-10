import type { HtmlOp } from './ops'
import type { ElementEntry, ParseMap } from './parse-map'

export type InsertKind =
  'heading' | 'paragraph' | 'list' | 'button' | 'image' | 'table' | 'section' | 'divider'

export const INSERT_KINDS: InsertKind[] = [
  'heading',
  'paragraph',
  'list',
  'button',
  'image',
  'table',
  'section',
  'divider',
]

/** kinds whose first edit is typing: the new element opens in text-edit mode */
export const TEXT_INSERT_KINDS = new Set<InsertKind>(['heading', 'paragraph', 'button'])

const STRUCTURAL = new Set(['html', 'head', 'body'])
/** head content that can sit at the top level of a document without an explicit <head> */
const HEAD_ONLY = new Set(['title', 'meta', 'link', 'style', 'script', 'base', 'noscript'])
/** parts that only make sense inside their list / table: never inserted between, never moved out */
export const TABLE_LIST_PARTS = new Set([
  'li',
  'dt',
  'dd',
  'tr',
  'td',
  'th',
  'thead',
  'tbody',
  'tfoot',
  'caption',
  'colgroup',
  'col',
  'option',
  'optgroup',
])

export interface InsertLabels {
  heading: string
  paragraph: string
  listItem: string
  button: string
  sectionTitle: string
  sectionBody: string
  imageAlt: string
  tableHeaders: string[]
  tableCell: string
}

const CELL_STYLE = 'border: 1px solid #d0d5dd; padding: 8px; text-align: left;'

/** ribbon choices that shape the starter markup */
export interface InsertOptions {
  /** remote image instead of a picked file */
  url?: string
  /** table rows including the header row (picker grid); default 3 */
  rows?: number
  /** table columns (picker grid); default 3 */
  cols?: number
}

/** markup for a starter element; text placeholders come from the UI language */
export function insertPresetHtml(
  kind: InsertKind,
  labels: InsertLabels,
  opts: { imageSrc?: string; tableBodyRows?: number } = {},
): string {
  switch (kind) {
    case 'heading':
      return `<h2>${esc(labels.heading)}</h2>`
    case 'paragraph':
      return `<p>${esc(labels.paragraph)}</p>`
    case 'list':
      return `<ul>\n  <li>${esc(labels.listItem)}</li>\n  <li>${esc(labels.listItem)}</li>\n  <li>${esc(labels.listItem)}</li>\n</ul>`
    case 'button':
      return `<button type="button">${esc(labels.button)}</button>`
    case 'image':
      return `<img src="${esc(opts.imageSrc ?? '')}" alt="${esc(labels.imageAlt)}" style="max-width: 100%;">`
    case 'table': {
      const th = labels.tableHeaders.map((h) => `<th style="${CELL_STYLE}">${esc(h)}</th>`).join('')
      const td = labels.tableHeaders
        .map(() => `<td style="${CELL_STYLE}">${esc(labels.tableCell)}</td>`)
        .join('')
      const body = Array.from(
        { length: Math.max(1, opts.tableBodyRows ?? 2) },
        () => `    <tr>${td}</tr>`,
      )
      return [
        '<table style="width: 100%; border-collapse: collapse;">',
        `  <thead>\n    <tr>${th}</tr>\n  </thead>`,
        `  <tbody>\n${body.join('\n')}\n  </tbody>`,
        '</table>',
      ].join('\n')
    }
    case 'section':
      return `<section style="padding: 32px 24px;">\n  <h2>${esc(labels.sectionTitle)}</h2>\n  <p>${esc(labels.sectionBody)}</p>\n</section>`
    case 'divider':
      return '<hr>'
  }
}

const insideHead = (map: ParseMap, e: ElementEntry): boolean => {
  for (let cur: ElementEntry | undefined = e; cur;) {
    if (cur.tag === 'head') return true
    cur = cur.parentSid === null ? undefined : map.bySid.get(cur.parentSid)
  }
  return false
}

/**
 * Where a new element goes: right after the selected element; else at the end
 * of <main> / <body>; else — documents with an implied body — after the last
 * top-level content element, or after <head>. `null` only when the document has
 * no element to anchor to.
 */
export function insertOp(
  map: ParseMap,
  selected: ElementEntry | undefined,
  html: string,
): HtmlOp | null {
  if (selected && !STRUCTURAL.has(selected.tag)) {
    // a block between two <li>s or <td>s is parsed out of its list / table: land after the whole thing
    let anchor = selected
    while (TABLE_LIST_PARTS.has(anchor.tag) && anchor.parentSid !== null) {
      const parent = map.bySid.get(anchor.parentSid)
      if (!parent || STRUCTURAL.has(parent.tag)) break
      anchor = parent
    }
    return { op: 'insert_html', sid: anchor.sid, position: 'after', html: `\n${html}` }
  }
  const container =
    map.elements.find((e) => e.tag === 'main') ?? map.elements.find((e) => e.tag === 'body')
  if (container)
    return { op: 'insert_html', sid: container.sid, position: 'append', html: `${html}\n` }
  const content = map.elements.filter(
    (e) => !STRUCTURAL.has(e.tag) && !HEAD_ONLY.has(e.tag) && !insideHead(map, e),
  )
  if (content.length > 0) {
    const depth = Math.min(...content.map((e) => e.depth))
    const last = content.filter((e) => e.depth === depth).at(-1)!
    return { op: 'insert_html', sid: last.sid, position: 'after', html: `\n${html}` }
  }
  const head = map.elements.find((e) => e.tag === 'head')
  if (head) return { op: 'insert_html', sid: head.sid, position: 'after', html: `\n${html}` }
  return null
}

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;')
}
