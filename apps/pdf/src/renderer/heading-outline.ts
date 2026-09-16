import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { OutlineNode } from './OutlinePanel'

/**
 * Fallback navigation for PDFs without embedded bookmarks: text lines set
 * noticeably larger than the body size become headings, size tiers become
 * nesting levels.
 */

interface RawTextItem {
  str?: string
  transform?: number[]
  hasEOL?: boolean
}

export interface TextLine {
  pageIndex: number
  text: string
  /** Dominant font size of the line, rounded to half a point */
  size: number
  /** Baseline in PDF user space */
  y: number
}

const MAX_PAGES = 400
const MIN_HEADINGS = 2
const MAX_HEADINGS = 1500
const MAX_LEVELS = 4
const MAX_TITLE_CHARS = 100
/** Above this share of lines the "body" size guess is unreliable (tables, slides) */
const MAX_HEADING_RATIO = 0.4

const half = (n: number): number => Math.round(n * 2) / 2

/** Group a page's text items into lines by baseline; each line keeps its dominant size */
export function collectLines(items: RawTextItem[], pageIndex: number): TextLine[] {
  const lines: TextLine[] = []
  let text = ''
  let y = 0
  let sizes = new Map<number, number>()
  let breakNext = true
  const flush = (): void => {
    const trimmed = text.replace(/\s+/g, ' ').trim()
    if (trimmed) {
      let size = 0
      let best = -1
      for (const [s, chars] of sizes) {
        if (chars > best || (chars === best && s > size)) {
          best = chars
          size = s
        }
      }
      lines.push({ pageIndex, text: trimmed, size, y })
    }
    text = ''
    sizes = new Map()
  }
  for (const it of items) {
    if (typeof it.str !== 'string' || !it.transform) continue
    const t = it.transform
    const size = half(Math.hypot(t[2] ?? 0, t[3] ?? 0))
    if (size <= 0) continue
    // Rotated runs (tilted baseline) are decoration, not headings
    if (Math.abs(t[1] ?? 0) > size * 1e-3) continue
    const itemY = t[5] ?? 0
    if (!breakNext && text && Math.abs(itemY - y) > 0.5 * Math.max(size, 1)) flush()
    if (!text) y = itemY
    text += it.str
    if (it.str.trim()) sizes.set(size, (sizes.get(size) ?? 0) + it.str.trim().length)
    breakNext = it.hasEOL === true
    if (breakNext) flush()
  }
  flush()
  return lines
}

/** Sentences, stat callouts ("900 units"), and letter-spaced decoration are not headings */
const looksLikeTitle = (text: string): boolean => {
  if (text.length > MAX_TITLE_CHARS) return false
  const letters = (text.match(/\p{L}/gu) ?? []).length
  const digits = (text.match(/\p{N}/gu) ?? []).length
  if (letters < 2 || digits > letters) return false
  if (/[.。．!！?？,，;；:：✓]$/.test(text)) return false
  const tokens = text.split(' ')
  const singles = tokens.filter((token) => [...token].length === 1).length
  return tokens.length < 4 || singles < tokens.length * 0.8
}

interface Heading {
  pageIndex: number
  text: string
  size: number
  y: number
}

/** Heading tree from collected lines; null when the document has no usable heading structure */
export function outlineFromLines(lines: TextLine[]): OutlineNode[] | null {
  if (lines.length === 0) return null
  const weight = new Map<number, number>()
  for (const line of lines) weight.set(line.size, (weight.get(line.size) ?? 0) + line.text.length)
  let body = 0
  let bodyWeight = -1
  for (const [size, chars] of weight) {
    if (chars > bodyWeight) {
      bodyWeight = chars
      body = size
    }
  }
  const isHeading = (line: TextLine): boolean =>
    line.size >= body + 1 && line.size >= body * 1.12 && looksLikeTitle(line.text)

  // Adjacent same-size lines are one wrapped heading
  const headings: Heading[] = []
  let headingLines = 0
  let lastIndex = -2
  lines.forEach((line, i) => {
    if (!isHeading(line)) return
    headingLines++
    const last = headings[headings.length - 1]
    if (
      last &&
      lastIndex === i - 1 &&
      last.pageIndex === line.pageIndex &&
      last.size === line.size &&
      last.y - line.y > 0 &&
      last.y - line.y <= 1.8 * line.size &&
      last.text.length + line.text.length <= MAX_TITLE_CHARS
    ) {
      last.text = `${last.text} ${line.text}`
    } else {
      headings.push({ ...line })
    }
    lastIndex = i
  })

  // Running headers/footers repeat across pages; drop them
  const pagesByText = new Map<string, Set<number>>()
  for (const h of headings) {
    const key = h.text.toLowerCase()
    const set = pagesByText.get(key) ?? new Set<number>()
    set.add(h.pageIndex)
    pagesByText.set(key, set)
  }
  const kept = headings.filter((h) => (pagesByText.get(h.text.toLowerCase())?.size ?? 0) < 3)

  if (kept.length < MIN_HEADINGS || kept.length > MAX_HEADINGS) return null
  if (headingLines > lines.length * MAX_HEADING_RATIO) return null

  // Size tiers → levels: sizes within a point of the tier share it
  const tiers: number[] = []
  for (const size of [...new Set(kept.map((h) => h.size))].sort((a, b) => b - a)) {
    const tier = tiers[tiers.length - 1]
    if (tier === undefined || tier - size > 1) tiers.push(size)
  }
  const levelOf = (size: number): number =>
    Math.min(
      MAX_LEVELS - 1,
      Math.max(
        0,
        tiers.findIndex((tier) => tier - size <= 1),
      ),
    )

  const root: OutlineNode[] = []
  const stack: { node: OutlineNode; level: number }[] = []
  for (const h of kept) {
    const level = levelOf(h.size)
    const node: OutlineNode = {
      title: h.text,
      dest: [h.pageIndex, { name: 'XYZ' }, null, h.y + h.size, null],
    }
    while (stack.length > 0 && stack[stack.length - 1]!.level >= level) stack.pop()
    const parent = stack[stack.length - 1]
    if (parent) (parent.node.items ??= []).push(node)
    else root.push(node)
    stack.push({ node, level })
  }
  return root
}

/** Scan the document text (first MAX_PAGES pages) and build the heading outline */
export async function buildHeadingOutline(
  doc: PDFDocumentProxy,
  cancelled: () => boolean = () => false,
): Promise<OutlineNode[] | null> {
  const lines: TextLine[] = []
  const pages = Math.min(doc.numPages, MAX_PAGES)
  for (let n = 1; n <= pages; n++) {
    if (cancelled()) return null
    const page = await doc.getPage(n)
    const content = await page.getTextContent()
    lines.push(...collectLines(content.items as RawTextItem[], n - 1))
  }
  return cancelled() ? null : outlineFromLines(lines)
}

/**
 * Re-point generated nodes after a save that deleted or reordered pages
 * (`pageMap`: old index → new index; missing = deleted, its children hoist up).
 * Embedded bookmarks (non-numeric destinations) pass through untouched.
 */
export function remapOutlinePages(
  nodes: OutlineNode[],
  pageMap: ReadonlyMap<number, number>,
): OutlineNode[] {
  return nodes.flatMap((node) => {
    const items = node.items ? remapOutlinePages(node.items, pageMap) : undefined
    const dest = node.dest
    if (!Array.isArray(dest) || typeof dest[0] !== 'number') {
      return [items ? { ...node, items } : node]
    }
    const page = pageMap.get(dest[0])
    if (page === undefined) return items ?? []
    return [{ ...node, dest: [page, ...dest.slice(1)], ...(items ? { items } : {}) }]
  })
}
