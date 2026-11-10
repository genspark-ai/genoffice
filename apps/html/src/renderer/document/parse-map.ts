import { parse, type DefaultTreeAdapterTypes as T } from 'parse5'

export interface ElementEntry {
  /** stable per-buffer element id; reassigned by sid matching across rebuilds */
  sid: number
  tag: string
  parentSid: number | null
  depth: number
  /** whole element, from `<tag` to the end of the close tag (or of the start tag for void / implied-end elements) */
  range: [number, number]
  startTag: [number, number]
  /** absent for void elements and implied end tags */
  endTag: [number, number] | null
  /** content between start and end tag; equals range end for void elements */
  inner: [number, number]
  /** `body > main > section:nth-of-type(2) > h2` */
  path: string
  /** direct child text nodes with a source location, in document order */
  textNodes: Array<[number, number]>
}

export interface ParseMap {
  version: number
  elements: ElementEntry[]
  bySid: Map<number, ElementEntry>
  /** parse5 recovery diagnostics, useful as a structure-health signal */
  errorCount: number
}

interface BuildState {
  nextSid: number
}

function isElement(node: T.Node): node is T.Element {
  return 'tagName' in node && typeof (node as T.Element).tagName === 'string'
}

function nthOfType(node: T.Element): number {
  const parent = node.parentNode
  if (!parent || !('childNodes' in parent)) return 1
  let n = 0
  for (const sibling of parent.childNodes) {
    if (isElement(sibling) && sibling.tagName === node.tagName) {
      n++
      if (sibling === node) return n
    }
  }
  return n
}

/** Elements with a source location, in document order, deduped by start-tag offset
 * (the adoption agency re-creates formatting elements that share one start tag). */
function collect(
  root: T.Node,
  out: Array<{ node: T.Element; parent: T.Element | null; depth: number }>,
) {
  const seen = new Set<number>()
  const walk = (node: T.Node, parent: T.Element | null, depth: number) => {
    if (isElement(node)) {
      const loc = node.sourceCodeLocation
      if (loc?.startTag) {
        if (!seen.has(loc.startTag.startOffset)) {
          seen.add(loc.startTag.startOffset)
          out.push({ node, parent, depth })
          parent = node
          depth++
        }
      }
    }
    if ('childNodes' in node) for (const child of node.childNodes) walk(child, parent, depth)
    if (isElement(node) && node.tagName === 'template') {
      const content = (node as T.Template).content
      if (content) for (const child of content.childNodes) walk(child, node, depth)
    }
  }
  walk(root, null, 0)
}

/** Previous entries bucketed by tag + path, each bucket sorted by start offset. */
function indexPrevious(previous: ParseMap | null): Map<string, ElementEntry[]> | null {
  if (!previous) return null
  const buckets = new Map<string, ElementEntry[]>()
  for (const entry of previous.elements) {
    const key = bucketKey(entry.tag, entry.path)
    const bucket = buckets.get(key)
    if (bucket) bucket.push(entry)
    else buckets.set(key, [entry])
  }
  // collect() emits strictly increasing start offsets, so this only makes the
  // per-bucket order explicit; equal offsets keep document order.
  for (const bucket of buckets.values()) bucket.sort((a, b) => a.startTag[0] - b.startTag[0])
  return buckets
}

function bucketKey(tag: string, path: string): string {
  return `${tag}\u0000${path}`
}

/** First index in the bucket whose start offset is >= offset. */
function lowerBound(bucket: readonly ElementEntry[], offset: number): number {
  let lo = 0
  let hi = bucket.length
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (bucket[mid]!.startTag[0] < offset) lo = mid + 1
    else hi = mid
  }
  return lo
}

/** Reuse the previous sid for an element with the same tag, same parent path and the
 * closest start offset. Only the two bucket neighbours can be closest, so the
 * nearest-offset search is a binary search instead of a scan of every entry; a
 * matched entry is removed from its bucket so a sid is never reused. Ties go to
 * the smaller offset, which is what the former document-order scan picked. */
function matchSid(
  entry: Omit<ElementEntry, 'sid'>,
  buckets: Map<string, ElementEntry[]> | null,
): number | null {
  if (!buckets) return null
  const bucket = buckets.get(bucketKey(entry.tag, entry.path))
  if (!bucket) return null
  const at = lowerBound(bucket, entry.startTag[0])
  let bestAt = at - 1
  if (at < bucket.length) {
    if (bestAt < 0) bestAt = at
    else {
      const left = Math.abs(bucket[bestAt]!.startTag[0] - entry.startTag[0])
      const right = Math.abs(bucket[at]!.startTag[0] - entry.startTag[0])
      if (right < left) bestAt = at
    }
  }
  if (bestAt < 0) return null
  const [matched] = bucket.splice(bestAt, 1)
  return matched!.sid
}

/** One past the highest sid of the previous build; a loop, not a spread, so a
 * document with more entries than the argument limit cannot throw. */
function nextSidAfter(previous: ParseMap | null): number {
  if (!previous) return 1
  let max = 0
  for (const entry of previous.elements) if (entry.sid > max) max = entry.sid
  return max + 1
}

export function buildParseMap(
  text: string,
  version: number,
  previous: ParseMap | null = null,
): ParseMap {
  let errorCount = 0
  const doc = parse(text, {
    sourceCodeLocationInfo: true,
    onParseError: () => {
      errorCount++
    },
  })
  const found: Array<{ node: T.Element; parent: T.Element | null; depth: number }> = []
  collect(doc, found)

  const state: BuildState = {
    nextSid: nextSidAfter(previous),
  }
  const buckets = indexPrevious(previous)
  const sidByNode = new Map<T.Element, number>()
  const pathByNode = new Map<T.Element, string>()
  const elements: ElementEntry[] = []

  for (const { node, parent, depth } of found) {
    const loc = node.sourceCodeLocation!
    const start: [number, number] = [loc.startTag!.startOffset, loc.startTag!.endOffset]
    const end: [number, number] | null = loc.endTag
      ? [loc.endTag.startOffset, loc.endTag.endOffset]
      : null
    const rangeEnd = end ? end[1] : loc.endOffset
    const inner: [number, number] = [start[1], end ? end[0] : Math.max(start[1], loc.endOffset)]
    const parentPath = parent ? pathByNode.get(parent) : undefined
    const segment =
      node.tagName === 'html' || node.tagName === 'head' || node.tagName === 'body'
        ? node.tagName
        : `${node.tagName}:nth-of-type(${nthOfType(node)})`
    const path = parentPath ? `${parentPath} > ${segment}` : segment
    pathByNode.set(node, path)
    const textNodes: Array<[number, number]> = []
    for (const child of node.childNodes) {
      const loc = (child as T.TextNode).sourceCodeLocation
      if ('value' in child && loc) textNodes.push([loc.startOffset, loc.endOffset])
    }
    const partial: Omit<ElementEntry, 'sid'> = {
      textNodes,
      tag: node.tagName,
      parentSid: parent ? (sidByNode.get(parent) ?? null) : null,
      depth,
      range: [start[0], rangeEnd],
      startTag: start,
      endTag: end,
      inner,
      path,
    }
    const sid = matchSid(partial, buckets) ?? state.nextSid++
    sidByNode.set(node, sid)
    elements.push({ sid, ...partial })
  }

  return {
    version,
    elements,
    bySid: new Map(elements.map((e) => [e.sid, e])),
    errorCount,
  }
}

/** Smallest element whose range covers [from, to]; null when nothing does. */
export function elementCovering(map: ParseMap, from: number, to: number): ElementEntry | null {
  let best: ElementEntry | null = null
  for (const e of map.elements) {
    if (e.range[0] <= from && e.range[1] >= to) {
      if (!best || e.range[1] - e.range[0] <= best.range[1] - best.range[0]) best = e
    }
  }
  return best
}

/** ancestors from the root down to (excluding) the element */
export function ancestorsOf(map: ParseMap, sid: number): ElementEntry[] {
  const out: ElementEntry[] = []
  // The map is parser-built, but a corrupt parentSid cycle (or self-loop)
  // would hang the breadcrumb render in an infinite loop: stop at repeats.
  const seen = new Set<number>([sid])
  let cur = map.bySid.get(sid)
  while (cur && cur.parentSid !== null) {
    if (seen.has(cur.parentSid)) break
    seen.add(cur.parentSid)
    const parent = map.bySid.get(cur.parentSid)
    if (!parent) break
    out.unshift(parent)
    cur = parent
  }
  return out
}

export function childrenOf(map: ParseMap, sid: number): ElementEntry[] {
  return map.elements.filter((e) => e.parentSid === sid)
}

/** plain text of the element's source content, tags stripped and whitespace collapsed */
export function sourceText(text: string, e: ElementEntry): string {
  return text
    .slice(e.inner[0], e.inner[1])
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim()
}
