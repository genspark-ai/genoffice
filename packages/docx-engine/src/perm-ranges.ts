import { attrsOf, childrenOf, deepXmlParser, findChild, nameOf, type XNode } from './xml-utils'

/**
 * Editing permission ranges (w:permStart / w:permEnd).
 *
 * A document protected as read-only (w:documentProtection w:edit="readOnly")
 * can still leave regions open to everyone: Word marks them with a range
 * start/end pair carrying w:edGrp="everyone", optionally limited to table
 * columns with w:colFirst / w:colLast. Fill-in templates (grant proposals,
 * forms) are built this way: a locked frame with one open cell per answer.
 *
 * The ranges are resolved here in one document-order walk over the body, at
 * the granularity the editor works with: a top-level paragraph, or a whole
 * cell of a top-level table. A paragraph or cell is editable when an open
 * range covers any of its content (text, tabs, breaks, pictures, equations);
 * one with no content, when a range is open at both its start and its end.
 * Markers alone open nothing: Word ends a range that selected whole
 * paragraphs with a w:permEnd at the start of the next one, and that
 * paragraph stays locked. This matches Word for ranges over whole paragraphs
 * or cells and is slightly more permissive for a range that covers only part
 * of a paragraph.
 *
 * Only w:edGrp="everyone" ranges open anything: ranges granted to a named
 * editor (w:ed) or another group cannot be matched to the person editing here.
 */

/** body-level element as scanned from document.xml */
interface ScannedElement {
  readonly name: string
  readonly start: number
  readonly end: number
}

export interface PermRangeMap {
  /** start offsets (in document.xml) of top-level paragraphs inside an editable range */
  readonly paragraphs: ReadonlySet<number>
  /**
   * start offset of each top-level table → one flag per w:tc, in the order
   * childrenThroughSdt reaches them (rows, then cells, through w:sdt wrappers)
   */
  readonly tables: ReadonlyMap<number, readonly boolean[]>
}

interface OpenRange {
  readonly colFirst?: number
  readonly colLast?: number
  /** table the range started in: column limits only apply to its cells */
  readonly table?: XNode
}

/** collectors record whether an open range covered the element being walked */
interface Collector {
  hit: boolean
  /** whether an open range covers this element where the walk is now */
  covered: () => boolean
  /** the element had content at all */
  content: boolean
  /** a range was open where the element starts */
  openAtStart: boolean
}

/** what a range must cover for a paragraph or cell to open: its visible content */
const CONTENT = new Set([
  'w:t',
  'w:tab',
  'w:br',
  'w:cr',
  'w:sym',
  'w:noBreakHyphen',
  'w:softHyphen',
  'w:drawing',
  'w:pict',
  'w:object',
  'm:oMath',
  'm:oMathPara',
])

/** recursion guard: table XML may legitimately nest thousands of levels */
const MAX_DEPTH = 256

export function hasPermRanges(documentXml: string): boolean {
  return documentXml.includes('<w:permStart')
}

const intAttr = (raw: string | undefined): number | undefined => {
  if (raw === undefined || !/^\d+$/.test(raw)) return undefined
  return Number(raw)
}

const gridSpanOf = (tc: XNode): number =>
  intAttr(attrsOf(findChild(findChild(tc, 'w:tcPr') ?? {}, 'w:gridSpan') ?? {})['w:val']) ?? 1

const gridBeforeOf = (tr: XNode): number =>
  intAttr(attrsOf(findChild(findChild(tr, 'w:trPr') ?? {}, 'w:gridBefore') ?? {})['w:val']) ?? 0

export function permRangeMap(
  documentXml: string,
  elements: readonly ScannedElement[],
): PermRangeMap {
  const open = new Map<string, OpenRange>()
  const collectors: Collector[] = []
  const paragraphs = new Set<number>()
  const tables = new Map<number, boolean[]>()

  /** content met: every element the open ranges cover there opens */
  const content = (): void => {
    for (const c of collectors) {
      c.content = true
      if (!c.hit && c.covered()) c.hit = true
    }
  }
  const collect = (covered: () => boolean): Collector => ({
    hit: false,
    covered,
    content: false,
    openAtStart: covered(),
  })
  /** an element with no content opens when a range runs through it */
  const settle = (c: Collector): boolean => c.hit || (!c.content && c.openAtStart && c.covered())

  const marker = (node: XNode, name: string, table: XNode | undefined): void => {
    const attrs = attrsOf(node)
    const id = attrs['w:id']
    if (id === undefined) return
    if (name === 'w:permStart') {
      if (attrs['w:edGrp'] !== 'everyone') return
      const colFirst = intAttr(attrs['w:colFirst'])
      const colLast = intAttr(attrs['w:colLast'])
      open.set(id, {
        ...(colFirst !== undefined ? { colFirst, colLast: colLast ?? colFirst } : {}),
        ...(table ? { table } : {}),
      })
    } else {
      open.delete(id)
    }
  }

  /** an open range covers a cell spanning grid columns [first, last] of `table` */
  const coversCell = (table: XNode, first: number, last: number): boolean => {
    for (const range of open.values()) {
      if (range.colFirst === undefined || range.table !== table) return true
      if (first <= range.colLast! && last >= range.colFirst) return true
    }
    return false
  }

  const walk = (nodes: readonly XNode[], depth: number, table: XNode | undefined): void => {
    if (depth > MAX_DEPTH) return
    for (const node of nodes) {
      const name = nameOf(node)
      if (!name) continue
      if (name === 'w:permStart' || name === 'w:permEnd') marker(node, name, table)
      else if (name === 'w:tbl') walkTable(node, depth + 1, undefined)
      else if (CONTENT.has(name)) content()
      else walk(childrenOf(node), depth + 1, table)
    }
  }

  /** walks rows and cells in childrenThroughSdt order; `flags` collects one entry per w:tc */
  const walkTable = (tbl: XNode, depth: number, flags: boolean[] | undefined): void => {
    if (depth > MAX_DEPTH) return
    const visitCell = (tc: XNode, col: number): void => {
      const last = col + gridSpanOf(tc) - 1
      const collector = collect(() => coversCell(tbl, col, last))
      collectors.push(collector)
      walk(childrenOf(tc), depth + 1, tbl)
      collectors.pop()
      flags?.push(settle(collector))
    }
    const visitRowChildren = (nodes: readonly XNode[], state: { col: number }): void => {
      for (const node of nodes) {
        const name = nameOf(node)
        if (name === 'w:tc') {
          visitCell(node, state.col)
          state.col += gridSpanOf(node)
        } else if (name === 'w:sdt') {
          visitRowChildren(childrenOf(findChild(node, 'w:sdtContent') ?? {}), state)
        } else if (name === 'w:permStart' || name === 'w:permEnd') {
          marker(node, name, tbl)
        } else if (name) {
          walk(childrenOf(node), depth + 1, tbl)
        }
      }
    }
    const visitTableChildren = (nodes: readonly XNode[]): void => {
      for (const node of nodes) {
        const name = nameOf(node)
        if (name === 'w:tr') {
          visitRowChildren(childrenOf(node), { col: gridBeforeOf(node) })
        } else if (name === 'w:sdt') {
          visitTableChildren(childrenOf(findChild(node, 'w:sdtContent') ?? {}))
        } else if (name === 'w:permStart' || name === 'w:permEnd') {
          marker(node, name, tbl)
        } else if (name) {
          walk(childrenOf(node), depth + 1, tbl)
        }
      }
    }
    visitTableChildren(childrenOf(tbl))
  }

  for (const el of elements) {
    const xml = documentXml.slice(el.start, el.end)
    let nodes: XNode[]
    try {
      nodes = deepXmlParser.parse(xml) as XNode[]
    } catch {
      continue
    }
    if (el.name === 'w:p') {
      const collector = collect(() => open.size > 0)
      collectors.push(collector)
      walk(nodes, 0, undefined)
      collectors.pop()
      if (settle(collector)) paragraphs.add(el.start)
    } else if (el.name === 'w:tbl') {
      const tbl = nodes.find((node) => nameOf(node) === 'w:tbl')
      if (!tbl) continue
      const flags: boolean[] = []
      walkTable(tbl, 0, flags)
      if (flags.some(Boolean)) tables.set(el.start, flags)
    } else {
      walk(nodes, 0, undefined)
    }
  }
  return { paragraphs, tables }
}

const PERM_MARKER = /<w:perm(?:Start|End)\b[^>]*?\/>/g

/** the w:permStart / w:permEnd elements that are direct children of one w:p, split by kind */
export function paragraphPermMarkers(pXml: string): { starts: string[]; ends: string[] } {
  const starts: string[] = []
  const ends: string[] = []
  // nested paragraphs (textboxes) carry their own markers: only scan this
  // paragraph's own level, outside any w:txbxContent
  const own = pXml.replace(/<w:txbxContent\b[\s\S]*?<\/w:txbxContent>/g, '')
  for (const m of own.matchAll(PERM_MARKER)) {
    if (m[0].startsWith('<w:permStart')) starts.push(m[0])
    else ends.push(m[0])
  }
  return { starts, ends }
}

const markerId = (markerXml: string): string | undefined => /\bw:id="([^"]*)"/.exec(markerXml)?.[1]

/**
 * Put back permission markers a rebuilt paragraph lost. `original` is the
 * paragraph as it was read, `rebuilt` the regenerated w:p. Starts go right
 * after the paragraph properties, ends right before the closing tag: a
 * rebuilt paragraph is wholly inside or outside a range, which is the
 * granularity editing works at. Markers the rebuilt XML already carries (by
 * w:id) are left alone.
 */
export function restoreParagraphPermMarkers(
  original: string,
  rebuilt: string,
  which: { starts: boolean; ends: boolean } = { starts: true, ends: true },
): string {
  const own = paragraphPermMarkers(original)
  const starts = which.starts ? own.starts : []
  const ends = which.ends ? own.ends : []
  if (starts.length === 0 && ends.length === 0) return rebuilt
  const present = new Set(
    [...rebuilt.matchAll(PERM_MARKER)].map(
      (m) => `${m[0].startsWith('<w:permStart') ? 's' : 'e'}${markerId(m[0])}`,
    ),
  )
  const missingStarts = starts.filter((s) => !present.has(`s${markerId(s)}`))
  const missingEnds = ends.filter((e) => !present.has(`e${markerId(e)}`))
  if (missingStarts.length === 0 && missingEnds.length === 0) return rebuilt
  return injectIntoParagraph(rebuilt, missingStarts.join(''), missingEnds.join(''))
}

/** insert `head` after the w:pPr (or the open tag) and `tail` before </w:p> */
export function injectIntoParagraph(pXml: string, head: string, tail: string): string {
  let out = pXml
  if (tail) {
    const close = out.lastIndexOf('</w:p>')
    if (close !== -1) out = out.slice(0, close) + tail + out.slice(close)
    else if (out.endsWith('/>')) out = `${out.slice(0, -2)}>${tail}</w:p>`
  }
  if (head) {
    const pPr = /^<w:p\b[^>]*>(<w:pPr\b[^>]*\/>|<w:pPr\b[\s\S]*?<\/w:pPr>)?/.exec(out)
    if (pPr) out = out.slice(0, pPr[0].length) + head + out.slice(pPr[0].length)
  }
  return out
}
