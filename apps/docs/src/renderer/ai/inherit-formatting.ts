/**
 * Formatting inheritance for AI rewrites (replace_blocks).
 *
 * The restricted HTML the model writes carries no font/size/indent/spacing,
 * column widths, borders or cell shading — and the model never saw them, the
 * selection context strips them too. A rewrite therefore has to take its
 * formatting from the blocks it replaces: paragraph attrs and the dominant run
 * style for text blocks (issue #175), and the whole table/row/cell property
 * set for tables, whose cells only ever reach the model as plain text.
 */
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { inlineToRuns, pmTableToModel, type PmMark, type PmNode } from '../editor/convert'

const hasDelMark = (node: ProseMirrorNode) => node.marks.some((m) => m.type.name === 'del')

// ---- text blocks ----

/**
 * Paragraph attrs a rewritten block keeps from the block it replaces.
 * docxIndex rides along so the save path treats the rewrite as an in-place
 * edit of the original paragraph and reuses its raw pPr bytes (numbering,
 * section breaks, keepNext…), exactly like retyping the text by hand.
 * Revision state, bookmarks and comment anchors stay out: the AI's own
 * ins/del handling owns the former, and the latter are text-range anchors.
 */
const INHERITED_PARA_ATTRS = [
  'docxIndex',
  'styleId',
  'align',
  'lineSpacing',
  'lineRule',
  'lineRawTwips',
  'snapToGrid',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'spaceBeforeAuto',
  'spaceAfterAuto',
  'contextualSpacing',
  'pageBreakBefore',
  'bidi',
  'autoSpace',
  'shadingFill',
  'borders',
  'borderLines',
  'tabStops',
  'emptyRunSize',
  'emptyRunFont',
] as const

/**
 * Run attrs the rewritten text inherits from the replaced block's dominant
 * run: the properties the save path models directly. Display-only attrs
 * (charSpacing/caps/em…) are kept faithful by rawRPr and rawRPr itself is
 * tied to the original run bytes (revision records included), so neither travels.
 */
const INHERITED_TEXT_STYLE_ATTRS = [
  'color',
  'sizeHalfPoints',
  'font',
  'eaSlotEmpty',
  'fontAscii',
  'csFont',
  'highlight',
  'shading',
  'styleId',
  'cs',
] as const

/** docTextStyle attrs of the run carrying the most live characters in the block */
function dominantTextStyle(node: ProseMirrorNode): Record<string, unknown> | null {
  const buckets = new Map<string, { attrs: Record<string, unknown>; chars: number }>()
  node.forEach((child) => {
    if (!child.isText || !child.text) return
    if (hasDelMark(child)) return
    const mark = child.marks.find((m) => m.type.name === 'docTextStyle')
    const picked: Record<string, unknown> = {}
    for (const key of INHERITED_TEXT_STYLE_ATTRS) {
      const value = mark?.attrs[key]
      if (value != null) picked[key] = value
    }
    const id = JSON.stringify(picked)
    const bucket = buckets.get(id) ?? { attrs: picked, chars: 0 }
    bucket.chars += child.text.length
    buckets.set(id, bucket)
  })
  let best: { attrs: Record<string, unknown>; chars: number } | null = null
  for (const bucket of buckets.values()) if (!best || bucket.chars > best.chars) best = bucket
  return best && Object.keys(best.attrs).length > 0 ? best.attrs : null
}

/** a replaced block can lend its formatting to a new block of the same role */
export function sameBlockRole(old: ProseMirrorNode, next: PmNode): boolean {
  if (old.type.name !== next.type) return false
  if (next.type === 'docHeading') return Number(old.attrs.level) === Number(next.attrs?.level)
  return true
}

/**
 * The new block's own formatting (from <pre>/<blockquote>/list parsing) wins
 * over inherited values. `first` = this is the first new block formatted from
 * `old`: the page break before the old paragraph belongs to that one block
 * only, or an expansion into several paragraphs would start a page per paragraph.
 */
export function inheritFrom(
  old: ProseMirrorNode,
  next: PmNode,
  opts: { anchor: boolean; first: boolean },
): PmNode {
  const attrs: Record<string, unknown> = {}
  for (const key of INHERITED_PARA_ATTRS) {
    if (key === 'docxIndex' && !opts.anchor) continue
    if (key === 'pageBreakBefore' && !opts.first) continue
    // false is a real value here (snapToGrid / spacing auto / contextualSpacing
    // off override the style chain); only unset attrs are skipped
    const value = old.attrs[key]
    if (value != null) attrs[key] = value
  }
  const own = next.attrs ?? {}
  for (const [key, value] of Object.entries(own)) {
    // blockNode() always emits these as null / true; they are not model choices
    if ((key === 'docxIndex' || key === 'styleId') && value === null) continue
    if (key === 'aiChanged') continue
    attrs[key] = value
  }
  if (next.type === 'docListItem' && old.attrs.kind === own.kind && old.attrs.numId) {
    // same list kind: stay in the original list so numbering continues
    attrs.numId = old.attrs.numId
  }
  attrs.aiChanged = true

  const inherited = dominantTextStyle(old)
  const content = inherited
    ? next.content?.map((child) => {
        if (child.type !== 'text') return child
        const marks = child.marks ?? []
        const ownStyle = marks.find((m) => m.type === 'docTextStyle')?.attrs ?? {}
        const merged: Record<string, unknown> = { ...inherited }
        for (const [key, value] of Object.entries(ownStyle)) if (value != null) merged[key] = value
        return {
          ...child,
          marks: [
            ...marks.filter((m) => m.type !== 'docTextStyle'),
            { type: 'docTextStyle', attrs: merged },
          ],
        }
      })
    : next.content
  return { ...next, attrs, ...(content ? { content } : {}) }
}

// ---- tables ----

const CELL_PARA_TYPES = new Set(['docParagraph', 'docListItem'])

/** cell attrs that describe how a cell looks, independent of its place in the grid */
const CELL_STYLE_ATTRS = [
  'fill',
  'color',
  'bold',
  'align',
  'vAlign',
  'borders',
  'cellMar',
  'textDirection',
] as const

/**
 * Row attrs that survive a row count change. Raw trPr bytes and row revision
 * records describe the original row and are dropped with its position.
 */
const ROW_STYLE_ATTRS = ['heightTwips', 'heightRule', 'repeatHeader', 'repeatHeaderEdited'] as const

/**
 * Inline marks a rewritten cell text takes over from the cell's dominant run,
 * complete with the run's docTextStyle. Cells reach the model as plain text
 * (see tableToHtml), so unlike body paragraphs the bold/italic/underline of
 * the old text has to be inherited too — and the whole mark set, rawRPr
 * included, exactly as typing into the cell would continue it: a uniformly
 * formatted cell then saves through the byte-exact cell-text patch instead of
 * a table regeneration. Links, comments and revision marks stay behind.
 */
const CELL_INHERITED_MARKS = new Set(['bold', 'italic', 'underline', 'docTextStyle'])

/**
 * Paragraph text as the model saw it (tableToHtml renders pmTableToModel's
 * run texts, the table parser collapses whitespace).
 */
function inlineText(content: PmNode[] | undefined): string {
  return inlineToRuns(content ?? [])
    .map((run) => run.text)
    .join('')
    .replace(/\s+/g, ' ')
    .trim()
}

function cellParagraphs(cell: PmNode): PmNode[] {
  return (cell.content ?? []).filter((n) => CELL_PARA_TYPES.has(n.type))
}

function cellText(cell: PmNode): string {
  return cellParagraphs(cell)
    .map((p) => inlineText(p.content))
    .join('\n')
}

function rowText(row: PmNode): string {
  return (row.content ?? []).map(cellText).join('\u0001')
}

/** inheritable marks of the run carrying the most live characters */
function dominantCellMarks(para: ProseMirrorNode): PmMark[] {
  const buckets = new Map<string, { marks: PmMark[]; chars: number }>()
  para.forEach((child) => {
    if (!child.isText || !child.text || hasDelMark(child)) return
    const marks: PmMark[] = child.marks
      .filter((m) => CELL_INHERITED_MARKS.has(m.type.name))
      .map((m) => ({
        type: m.type.name,
        ...(m.type.name === 'docTextStyle' ? { attrs: m.attrs } : {}),
      }))
    const id = JSON.stringify(marks)
    const bucket = buckets.get(id) ?? { marks, chars: 0 }
    bucket.chars += child.text.length
    buckets.set(id, bucket)
  })
  let best: { marks: PmMark[]; chars: number } | null = null
  for (const bucket of buckets.values()) if (!best || bucket.chars > best.chars) best = bucket
  return best?.marks ?? []
}

/**
 * A rewritten cell paragraph: the model's text with the paragraph attrs (list
 * membership included) and the dominant run marks of the old paragraph. Marks
 * the table parser put on the text (synthetic header bold — the model cannot
 * write inline formatting in cells) are dropped, so a header cell that was not
 * bold does not turn bold.
 */
function inheritCellParagraph(old: ProseMirrorNode, next: PmNode, first: boolean): PmNode {
  const marks = dominantCellMarks(old)
  const list = old.type.name === 'docListItem'
  const plain: PmNode = {
    type: list ? 'docListItem' : 'docParagraph',
    attrs: list ? { kind: old.attrs.kind, numId: old.attrs.numId, ilvl: old.attrs.ilvl } : {},
    content: (next.content ?? []).map((child) =>
      child.type === 'text'
        ? { type: 'text', text: child.text, ...(marks.length > 0 ? { marks } : {}) }
        : child,
    ),
  }
  // paragraph attrs only: the text already carries the complete run marks
  const styled = inheritFrom(old, plain, { anchor: false, first })
  return { ...styled, content: plain.content }
}

/**
 * New cell content: unchanged text keeps the old cell's content verbatim
 * (per-run formatting, hyperlinks, images, nested tables); rewritten text
 * inherits from the old paragraphs (paragraph i from old paragraph i, extra
 * paragraphs from the last one). Nested tables and anchored boxes stay put.
 */
function inheritCellContent(old: ProseMirrorNode, oldJson: PmNode, next: PmNode): PmNode[] {
  if (cellText(oldJson) === cellText(next)) return oldJson.content ?? []
  const templates: ProseMirrorNode[] = []
  old.forEach((child) => {
    if (CELL_PARA_TYPES.has(child.type.name)) templates.push(child)
  })
  const paras = cellParagraphs(next)
  const content: PmNode[] =
    templates.length === 0
      ? paras
      : paras.map((p, i) => {
          const j = Math.min(i, templates.length - 1)
          return inheritCellParagraph(templates[j], p, i < templates.length)
        })
  // keep the cell's non-text members where they were (clamped to the new length)
  let paraIndex = 0
  const extras: Array<{ at: number; node: PmNode }> = []
  for (const node of oldJson.content ?? []) {
    if (CELL_PARA_TYPES.has(node.type)) paraIndex++
    else extras.push({ at: Math.min(paraIndex, content.length), node })
  }
  for (let i = extras.length - 1; i >= 0; i--) content.splice(extras[i].at, 0, extras[i].node)
  return content
}

/**
 * `full`: the new cell sits where the old one was, in an identical grid —
 * every attr (spans, widths, raw tcPr bytes, revision records) carries over.
 * `style`: the grid changed; only the look carries over, geometry comes from
 * the parsed cell.
 */
function inheritCell(
  old: ProseMirrorNode,
  oldJson: PmNode,
  next: PmNode,
  mode: 'full' | 'style',
): PmNode {
  const attrs: Record<string, unknown> =
    mode === 'full' ? { ...oldJson.attrs } : { ...(next.attrs ?? {}) }
  if (mode === 'style') {
    for (const key of CELL_STYLE_ATTRS) {
      const value = oldJson.attrs?.[key]
      if (value != null) attrs[key] = value
    }
  }
  return { type: oldJson.type, attrs, content: inheritCellContent(old, oldJson, next) }
}

/**
 * Which old row each new row is formatted from, when the row count changed.
 * Rows with identical text pair up in order (anchors); the header row pairs
 * with the header row unless either matched elsewhere. Between anchors, rows
 * line up by position — from the end when the model added rows there, so a
 * "Total" row edited along with the insertion keeps its own formatting and the
 * inserted rows take the row above them, like Word's insert-row-below.
 */
function pairRows(oldRows: PmNode[], newRows: PmNode[]): number[] {
  const oldTexts = oldRows.map(rowText)
  const template = newRows.map(() => -1)
  let cursor = 0
  newRows.forEach((row, r) => {
    const text = rowText(row)
    const j = oldTexts.findIndex((t, k) => k >= cursor && t === text)
    if (j !== -1) {
      template[r] = j
      cursor = j + 1
    }
  })
  if (template[0] === -1 && !template.includes(0)) template[0] = 0

  const assign = (news: number[], olds: number[], above: number, below: number) => {
    if (olds.length === 0) {
      for (const r of news) template[r] = above !== -1 ? above : below
    } else if (news.length <= olds.length) {
      news.forEach((r, i) => (template[r] = olds[i]))
    } else {
      const shift = news.length - olds.length
      news.forEach((r, i) => {
        template[r] = i < shift ? (above !== -1 ? above : olds[0]) : olds[i - shift]
      })
    }
  }
  let prevNew = -1
  let prevOld = -1
  for (let r = 0; r <= newRows.length; r++) {
    const anchorOld = r === newRows.length ? oldRows.length : template[r]
    if (anchorOld === -1) continue
    const news: number[] = []
    for (let k = prevNew + 1; k < r; k++) news.push(k)
    const olds: number[] = []
    for (let k = prevOld + 1; k < anchorOld; k++) olds.push(k)
    if (news.length > 0) assign(news, olds, prevOld, anchorOld < oldRows.length ? anchorOld : -1)
    prevNew = r
    prevOld = anchorOld
  }
  return template
}

type CellRef = { node: ProseMirrorNode; json: PmNode }

/**
 * The PM cell behind each grid entry of each row, aligned 1:1 with
 * pmTableToModel's rows: a cell spanning several rows owns the entry of every
 * row it covers (the merged-away rows have no cell of their own there).
 */
function entryOwners(old: ProseMirrorNode, oldJson: PmNode): CellRef[][] {
  type Active = { remaining: number; span: number; owner: CellRef }
  let active = new Map<number, Active>()
  const rows: CellRef[][] = []
  ;(oldJson.content ?? []).forEach((rowJson, r) => {
    const rowNode = old.child(r)
    const entries: Array<{ column: number; owner: CellRef }> = []
    const occupied = new Set<number>()
    for (const [column, span] of active) {
      entries.push({ column, owner: span.owner })
      for (let i = 0; i < span.span; i++) occupied.add(column + i)
    }
    const added = new Map<number, Active>()
    let cursor = 0
    ;(rowJson.content ?? []).forEach((json, i) => {
      while (occupied.has(cursor)) cursor++
      const colspan = Math.max(1, Number(json.attrs?.colspan) || 1)
      const rowspan = Math.max(1, Number(json.attrs?.rowspan) || 1)
      const owner = { node: rowNode.child(i), json }
      entries.push({ column: cursor, owner })
      if (rowspan > 1) added.set(cursor, { remaining: rowspan - 1, span: colspan, owner })
      for (let k = 0; k < colspan; k++) occupied.add(cursor + k)
      cursor += colspan
    })
    rows.push(entries.sort((a, b) => a.column - b.column).map((e) => e.owner))
    const next = new Map<number, Active>()
    for (const [column, span] of active) {
      if (span.remaining > 1) next.set(column, { ...span, remaining: span.remaining - 1 })
    }
    for (const [column, span] of added) next.set(column, span)
    active = next
  })
  return rows
}

/**
 * Give a rewritten table the formatting of the table it replaces.
 *
 * The model saw the table as tableToHtml renders pmTableToModel: one <td> per
 * grid entry (a spanning cell once, merged-away cells as empty cells). When
 * the parsed table has the same grid, every table/row/cell attr carries over
 * and each cell keeps its content if the text did not change — the rewrite
 * then saves through the same surgical cell-text patch as typing into the cell.
 *
 * When the model added or removed rows (or columns), the table props still
 * carry over; rows pair up as pairRows describes and cells take their look
 * from the cell in the same column of the paired row (plus its spans and
 * widths while the column grid is unchanged; vertical merges dissolve, see
 * below). Column widths follow the parsed table only when the column count
 * changed.
 *
 * `anchor`: keep docxIndex so the save path treats this as an edit of the
 * original table (false for tracked rewrites, where the old table stays).
 */
export function inheritTableFormatting(
  old: ProseMirrorNode,
  next: PmNode,
  opts: { anchor: boolean },
): PmNode {
  const oldJson = old.toJSON() as PmNode
  const oldRows = oldJson.content ?? []
  const newRows = next.content ?? []
  if (oldRows.length === 0 || newRows.length === 0) return next
  const grid = pmTableToModel(oldJson).rows
  // the model saw one <td> per grid entry — a spanning cell once, a merged-away
  // slot as an empty cell — so an equal <td> count means the column grid is
  // unchanged even when the old table has horizontal merges
  const oldCols = Math.max(...grid.map((row) => row.length))
  const newCols = Math.max(...newRows.map((row) => row.content?.length ?? 0))
  const sameGrid = newCols === oldCols
  // Vertical merges only survive an unchanged row count: an inserted row cannot
  // be placed relative to a merge the model never saw (it would end up under a
  // rowspan or leave the merged-away row with a hole), so a row count change
  // dissolves them into plain cells — the look stays, the merge does not. Only
  // the rows a merge runs through are affected; a spanning title row elsewhere
  // in the table keeps its column span.
  const rowMerged = (r: number): boolean => grid[r].some((cell) => cell.vMerge != null)
  let owners: CellRef[][] | undefined
  const ownerOf = (r: number, c: number): CellRef => (owners ??= entryOwners(old, oldJson))[r][c]

  // exact grid: the parsed row lists one cell per physical grid cell (padding
  // cells past a ragged row are only accepted when they are empty)
  const exactRow = (r: number, row: PmNode): boolean => {
    const cells = row.content ?? []
    const physical = grid[r]
    if (!physical || cells.length < physical.length) return false
    if (!cells.slice(physical.length).every((c) => cellText(c) === '')) return false
    const live = physical.filter((c) => c.vMerge !== 'continue').length
    return live === (oldRows[r].content?.length ?? 0)
  }
  const exact = newRows.length === oldRows.length && newRows.every((row, r) => exactRow(r, row))

  /**
   * Row r of the old table, cell by cell against the parsed row. `dissolve`
   * gives each merged-away slot a plain cell of its own, styled like the cell
   * that spanned it, and drops every rowspan.
   */
  const mapGridRow = (r: number, row: PmNode, dissolve: boolean): PmNode => {
    const oldRow = old.child(r)
    const oldRowJson = oldRows[r]
    const cells: PmNode[] = []
    let pmIndex = 0
    grid[r].forEach((physical, c) => {
      const newCell = row.content![c]
      if (physical.vMerge === 'continue') {
        if (!dissolve) return // merged away: no cell of its own
        const owner = ownerOf(r, c)
        const cell = inheritCell(owner.node, owner.json, newCell, 'style')
        cell.attrs!.colspan = owner.json.attrs?.colspan ?? 1
        cell.attrs!.colwidth = owner.json.attrs?.colwidth ?? null
        cells.push(cell)
        return
      }
      const cell = inheritCell(oldRow.child(pmIndex), oldRowJson.content![pmIndex], newCell, 'full')
      if (dissolve) cell.attrs!.rowspan = 1
      cells.push(cell)
      pmIndex++
    })
    return { ...oldRowJson, content: cells }
  }

  const rows: PmNode[] = exact
    ? newRows.map((row, r) => mapGridRow(r, row, false))
    : (() => {
        const template = pairRows(oldRows, newRows)
        return newRows.map((row, r) => {
          const j = template[r]
          if (sameGrid && exactRow(j, row)) return mapGridRow(j, row, rowMerged(j))
          const oldRow = old.child(j)
          const oldRowJson = oldRows[j]
          const oldCells = oldRowJson.content ?? []
          // in an unchanged grid the template row also lends its column geometry
          // (spans and widths), cell by cell — unless a vertical merge runs
          // through it, where its cells no longer line up with the columns
          const geometry =
            sameGrid && !rowMerged(j) && oldCells.length === (row.content?.length ?? 0)
          const cells = (row.content ?? []).map((newCell, c) => {
            const k = Math.min(c, oldCells.length - 1)
            const cell = inheritCell(oldRow.child(k), oldCells[k], newCell, 'style')
            if (geometry) {
              cell.attrs!.colspan = oldCells[k].attrs?.colspan ?? 1
              cell.attrs!.colwidth = oldCells[k].attrs?.colwidth ?? null
            }
            return cell
          })
          const attrs: Record<string, unknown> = { ...(row.attrs ?? {}) }
          for (const key of ROW_STYLE_ATTRS) {
            const value = oldRowJson.attrs?.[key]
            if (value != null) attrs[key] = value
          }
          return { type: 'docTableRow', attrs, content: cells }
        })
      })()

  const attrs: Record<string, unknown> = {
    ...oldJson.attrs,
    docxIndex: opts.anchor ? (oldJson.attrs?.docxIndex ?? null) : null,
    blockRevision: null,
  }
  if (!exact && !sameGrid) {
    // the old column grid no longer applies; the parser's equal split does
    attrs.colWidthsPct = next.attrs?.colWidthsPct ?? null
    attrs.widthPx = next.attrs?.widthPx ?? null
  }
  return { type: 'docTable', attrs, content: rows }
}
