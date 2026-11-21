/**
 * Insert > File with a range: the part of another document a bookmark
 * covers, as a package of its own that the ordinary parse reads.
 */
import { decodeEntities } from './parse-xml-text'
import { scanBody, type BodyElement } from './scan'
import { loadDocxZip } from './zip-load'

const MAIN_PART = 'word/document.xml'
const TAG = /<(\/?)([A-Za-z_][\w.-]*(?::[\w.-]+)?)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g
const START = /<w:bookmarkStart\b[^>]*>/g
/** children of a paragraph that are not content: the range may leave a paragraph holding only these */
const MARKERS = new Set([
  'w:pPr',
  'w:bookmarkStart',
  'w:bookmarkEnd',
  'w:proofErr',
  'w:permStart',
  'w:permEnd',
])

const attr = (tag: string, name: string): string | null =>
  new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1] ?? null

/** the bookmarks of the body a user can name, in document order (Word's own `_` ones left out) */
function bookmarksIn(xml: string): Array<{ name: string; id: string; at: number }> {
  const out: Array<{ name: string; id: string; at: number }> = []
  for (const m of xml.matchAll(START)) {
    const name = decodeEntities(attr(m[0], 'w:name') ?? '')
    const id = attr(m[0], 'w:id')
    if (name && id !== null && !name.startsWith('_') && !out.some((b) => b.name === name))
      out.push({ name, id, at: m.index })
  }
  return out
}

/** names of the bookmarks Insert > File can take as a range */
export async function docxBookmarkNames(bytes: Uint8Array): Promise<string[]> {
  const xml = await (await loadDocxZip(bytes)).file(MAIN_PART)?.async('string')
  if (!xml) return []
  const { bodyContentStart, bodyContentEnd } = scanBody(xml)
  return bookmarksIn(xml)
    .filter((b) => b.at >= bodyContentStart && b.at < bodyContentEnd)
    .map((b) => b.name)
}

interface Child {
  name: string
  start: number
  end: number
}

/** the direct children of the element at [start, end) */
function childrenOf(xml: string, start: number, end: number): Child[] {
  const out: Child[] = []
  let depth = 0
  let open: { name: string; start: number } | null = null
  TAG.lastIndex = start
  for (let m = TAG.exec(xml); m && m.index < end; m = TAG.exec(xml)) {
    const [text, closing, name, , selfClosing] = m
    if (closing) {
      depth--
      if (depth === 1 && open) {
        out.push({ ...open, end: m.index + text.length })
        open = null
      }
    } else if (selfClosing) {
      if (depth === 1) out.push({ name, start: m.index, end: m.index + text.length })
    } else {
      if (depth === 1) open = { name, start: m.index }
      depth++
    }
  }
  return out
}

/**
 * A paragraph cut down to its children inside [from, to); null when a bound
 * is not between two of its children (a bookmark inside a hyperlink, a field,
 * a content control), so the paragraph is taken whole.
 */
function trimmedParagraph(
  xml: string,
  el: BodyElement,
  from: number | null,
  to: number | null,
): { xml: string; content: boolean } | null {
  const children = childrenOf(xml, el.start, el.end)
  const between = (at: number): boolean =>
    children.some((c) => c.start === at) || children.every((c) => c.end <= at || c.start > at)
  if ((from !== null && !between(from)) || (to !== null && !between(to))) return null
  const kept = children.filter(
    (c) =>
      c.name === 'w:pPr' || ((from === null || c.start > from) && (to === null || c.end <= to)),
  )
  const openEnd = xml.indexOf('>', el.start) + 1
  return {
    xml: `${xml.slice(el.start, openEnd)}${kept.map((c) => xml.slice(c.start, c.end)).join('')}</w:p>`,
    content: kept.some((c) => !MARKERS.has(c.name)),
  }
}

export interface BookmarkRange {
  /** the package with only the bookmark's content in its body */
  bytes: Uint8Array
  /** the range lies inside one paragraph and does not take its paragraph mark */
  inline: boolean
}

/**
 * The content between a bookmark's start and end as a document. A range that
 * starts or ends inside a paragraph takes only that part of it; one that
 * starts or ends inside a table takes the whole table. Null when the document
 * has no such bookmark.
 */
export async function docxBookmarkRange(
  bytes: Uint8Array,
  name: string,
): Promise<BookmarkRange | null> {
  const zip = await loadDocxZip(bytes)
  const part = zip.file(MAIN_PART)
  const xml = await part?.async('string')
  if (!part || !xml) return null
  const mark = bookmarksIn(xml).find((b) => b.name === name)
  if (!mark) return null
  const endRe = new RegExp(`<w:bookmarkEnd\\b[^>]*\\bw:id="${mark.id}"[^>]*>`, 'g')
  endRe.lastIndex = mark.at
  const end = endRe.exec(xml)
  if (!end) return null

  const scan = scanBody(xml)
  const elements = scan.elements.filter(
    (el) => !['w:sectPr', 'w:bookmarkStart', 'w:bookmarkEnd'].includes(el.name),
  )
  const pieces: string[] = []
  let inline = false
  for (const el of elements) {
    if (el.end <= mark.at) continue
    // a block that opens at or after the end marker is past the range
    if (el.start >= end.index) break
    const from = el.start < mark.at ? mark.at : null
    const to = el.end > end.index ? end.index : null
    const whole = xml.slice(el.start, el.end)
    const trimmed =
      el.name === 'w:p' && (from !== null || to !== null)
        ? trimmedParagraph(xml, el, from, to)
        : null
    if (!trimmed) pieces.push(whole)
    else if (trimmed.content) {
      pieces.push(trimmed.xml)
      inline = to !== null && pieces.length === 1
    } else if (to === null) {
      // the range opens at the end of this paragraph: it takes the paragraph mark
      pieces.push(trimmed.xml)
    }
  }
  if (pieces.length === 0) return null

  const sectPr = scan.elements.filter((el) => el.name === 'w:sectPr').pop()
  const body = pieces.join('') + (sectPr ? xml.slice(sectPr.start, sectPr.end) : '')
  zip.file(MAIN_PART, xml.slice(0, scan.bodyContentStart) + body + xml.slice(scan.bodyContentEnd))
  return {
    bytes: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
    inline,
  }
}
