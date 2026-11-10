// Structured document tags (w:sdt): block-level SDT classification and the
// split of an SDT's run/table content.
import type { SdtShell } from './types'

/** a <w:sdt>…</w:sdt> element's absolute byte ranges inside the scanned xml */
export interface SdtRange {
  /** index of the opening <w:sdt> tag */
  openStart: number
  /** end of the <w:sdtContent> opening tag (the content starts here) */
  openEnd: number
  /** start of the matching </w:sdtContent> closing tag (the content ends here) */
  closeStart: number
  /** end of the closing </w:sdt> tag */
  closeEnd: number
}

/** one level of a nested block-level control chain, outermost first */
export interface SdtChainLevel extends SdtRange {
  /** metadata and this level's own open/close bytes (full slices of the scanned xml) */
  shell: SdtShell
}

/**
 * Every <w:sdt> element inside xml with its own open/close byte ranges,
 * outermost first. Nested controls keep their own ranges, so a specific level
 * can be unwrapped later without guessing which closing tag belongs to it.
 */
export function sdtRanges(xml: string): SdtRange[] {
  const ranges: SdtRange[] = []
  const stack: Array<{ openStart: number; openEnd: number; closeStart: number }> = []
  const tagRe = /<(\/?)([A-Za-z0-9:._-]+)((?:"[^"]*"|'[^']*'|[^"'>])*)>/g
  let m: RegExpExecArray | null
  while ((m = tagRe.exec(xml)) !== null) {
    const closing = m[1] === '/'
    const selfClosing = m[3].endsWith('/')
    const top = stack[stack.length - 1]
    if (!closing && m[2] === 'w:sdt') {
      if (!selfClosing) stack.push({ openStart: m.index, openEnd: -1, closeStart: -1 })
      continue
    }
    if (!closing && m[2] === 'w:sdtContent') {
      if (top) {
        top.openEnd = m.index + m[0].length
        if (selfClosing) top.closeStart = top.openEnd
      }
      continue
    }
    if (closing && m[2] === 'w:sdtContent') {
      if (top) top.closeStart = m.index
      continue
    }
    if (closing && m[2] === 'w:sdt') {
      const level = stack.pop()
      if (level && level.openEnd !== -1 && level.closeStart !== -1) {
        ranges.push({
          openStart: level.openStart,
          openEnd: level.openEnd,
          closeStart: level.closeStart,
          closeEnd: m.index + m[0].length,
        })
      }
    }
  }
  return ranges.sort((a, b) => a.openStart - b.openStart)
}

const shellFromRange = (xml: string, range: SdtRange): SdtShell => ({
  ...sdtMeta(xml.slice(range.openStart, range.closeEnd)),
  openXml: xml.slice(range.openStart, range.openEnd),
  closeXml: xml.slice(range.closeStart, range.closeEnd),
})

/** the controls that enclose [childStart, childEnd), outermost first */
function chainOf(
  xml: string,
  ranges: SdtRange[],
  childStart: number,
  childEnd: number,
): SdtChainLevel[] {
  const chain: SdtChainLevel[] = []
  for (const range of ranges) {
    if (range.openEnd === -1 || range.closeStart === -1) continue
    if (range.openStart >= childStart || range.closeEnd <= childEnd) continue
    chain.push({ ...range, shell: shellFromRange(xml, range) })
  }
  return chain
}

/**
 * Extract a SdtShell from a raw <w:sdt>…</w:sdt> XML string.
 * Returns the outermost shell, the enclosing-control chain, and the first
 * <w:p> found in <w:sdtContent>, or null if no usable paragraph is found.
 */
export function parseSdtBlock(
  sdtXml: string,
): { shell: SdtShell; pXml: string; chain: SdtChainLevel[] } | null {
  // --- find the sdtContent region ---
  // sdtContent is a direct child of w:sdt; its content may be a w:p or w:tbl
  const contentOpen = /<w:sdtContent(?:\s[^>]*)?>/.exec(sdtXml)
  if (!contentOpen) return null

  const contentTagEnd = contentOpen.index + contentOpen[0].length
  // find the matching </w:sdtContent> — last one, so nested content controls
  // don't truncate the region (same boundary rule as splitSdtParts/sdtTableXml)
  const contentClose = sdtXml.lastIndexOf('</w:sdtContent>')
  if (contentClose < contentTagEnd) return null

  // extract the first w:p inside sdtContent
  const innerContent = sdtXml.slice(contentTagEnd, contentClose)
  const pStart = innerContent.search(/<w:p[\s/>]/)
  if (pStart === -1) return null

  // find matching closing </w:p>
  let depth = 0
  let pEnd = -1
  const tagRe = /<\/?w:p(?=[\s/>])/g
  let m: RegExpExecArray | null
  while ((m = tagRe.exec(innerContent)) !== null) {
    if (m[0].startsWith('</')) {
      if (depth === 1) {
        pEnd = m.index + m[0].length + 1 // include '>'
        break
      }
      depth--
    } else {
      depth++
    }
  }
  if (pEnd === -1) {
    // self-closing <w:p/> or no </w:p> found — unlikely but safe
    const selfClose = /<w:p\/>/.exec(innerContent)
    pEnd = selfClose ? selfClose.index + selfClose[0].length : innerContent.length
  }
  const pXml = innerContent.slice(pStart, pEnd)

  // openXml = everything from start of sdt to end of <w:sdtContent> open tag
  const openXml = sdtXml.slice(0, contentTagEnd)
  const closeXml = sdtXml.slice(contentClose) // "</w:sdtContent></w:sdt>"

  const ranges = sdtRanges(sdtXml)
  // the paragraph sits inside the outer control plus any controls nested
  // around it (single-child path: the outer sdtContent holds only an inner sdt)
  let chain = chainOf(sdtXml, ranges, contentTagEnd + pStart, contentTagEnd + pEnd)
  if (chain.length === 0) {
    chain = [
      {
        openStart: 0,
        openEnd: contentTagEnd,
        closeStart: contentClose,
        closeEnd: sdtXml.length,
        shell: { ...sdtMeta(sdtXml), openXml, closeXml },
      },
    ]
  }

  return { shell: chain[0]!.shell, pXml, chain }
}

export function sdtMeta(
  sdtXml: string,
): Pick<SdtShell, 'alias' | 'tag' | 'controlType' | 'lock' | 'id'> {
  const sdtPrXml = /<w:sdtPr>([\s\S]*?)<\/w:sdtPr>/.exec(sdtXml)?.[1] ?? ''
  // w:val may be single- or double-quoted; the caller slices out the single open
  // tag first so the value can't belong to a later sibling element
  const attrVal = (tag: string): string => {
    const m = /\bw:val=(?:"([^"]*)"|'([^']*)')/.exec(tag)
    return m?.[1] ?? m?.[2] ?? ''
  }
  const alias = attrVal(/<w:alias[^>]*>/.exec(sdtPrXml)?.[0] ?? '')
  const tag = attrVal(/<w:tag[^>]*>/.exec(sdtPrXml)?.[0] ?? '')
  // the lock state and the control id: a downstream check finds a control by
  // its tag and confirms it is still locked, so both are read with the tag
  const lock = attrVal(/<w:lock[^>]*>/.exec(sdtPrXml)?.[0] ?? '')
  const id = attrVal(/<w:id[^>]*>/.exec(sdtPrXml)?.[0] ?? '')
  let controlType: SdtShell['controlType'] = 'text'
  if (/<w:date[\s/>]/.test(sdtPrXml)) controlType = 'date'
  else if (/<w:dropDownList[\s/>]|<w:comboBox[\s/>]/.test(sdtPrXml)) controlType = 'dropdown'
  else if (/<w:checkbox[\s/>]/.test(sdtPrXml)) controlType = 'checkbox'
  else if (/<w:text[\s/>]|<w:richText[\s/>]/.test(sdtPrXml)) controlType = 'text'
  return { alias, tag, controlType, lock, id }
}

interface SdtPart {
  name: string
  /** slice of the whole <w:sdt> xml owned by this part (parts partition the sdt exactly) */
  start: number
  end: number
  /** the w:p / w:tbl child inside [start, end) */
  childStart: number
  childEnd: number
  /** every control enclosing this child, outermost first (covers nested controls) */
  chain: SdtChainLevel[]
}

/**
 * When <w:sdtContent> holds several top-level w:p / w:tbl children (Word TOC,
 * multi-paragraph rich-text controls), each child becomes its own block.
 * Returns null for 0-1 children (single-block path keeps its behavior).
 */
export function splitSdtParts(sdtXml: string): SdtPart[] | null {
  const contentOpen = /<w:sdtContent(?:\s[^>]*)?>/.exec(sdtXml)
  if (!contentOpen) return null
  const innerStart = contentOpen.index + contentOpen[0].length
  const innerEnd = sdtXml.lastIndexOf('</w:sdtContent>')
  if (innerEnd <= innerStart) return null

  const children: Array<{ name: string; start: number; end: number }> = []
  const tagRe = /<(\/?)([A-Za-z0-9:._-]+)((?:"[^"]*"|'[^']*'|[^"'>])*)>/g
  tagRe.lastIndex = innerStart
  let depth = 0
  let start = -1
  let name = ''
  let m: RegExpExecArray | null
  while ((m = tagRe.exec(sdtXml)) !== null && m.index < innerEnd) {
    // nested content controls are transparent containers (Word cover-page
    // building blocks wrap each field in its own sdt): descend instead of
    // treating the inner sdt as one opaque child
    if (m[2] === 'w:sdt' || m[2] === 'w:sdtContent') continue
    if (m[1] === '/') {
      depth--
      if (depth === 0) children.push({ name, start, end: m.index + m[0].length })
    } else if (m[3].endsWith('/')) {
      if (depth === 0) children.push({ name: m[2], start: m.index, end: m.index + m[0].length })
    } else {
      if (depth === 0) {
        start = m.index
        name = m[2]
      }
      depth++
    }
  }

  const parts = children.filter((c) => c.name === 'w:p' || c.name === 'w:tbl')
  if (parts.length < 2) return null
  const ranges = sdtRanges(sdtXml)
  return parts.map((c, k) => ({
    name: c.name,
    start: k === 0 ? 0 : c.start,
    end: k === parts.length - 1 ? sdtXml.length : parts[k + 1].start,
    childStart: c.start,
    childEnd: c.end,
    chain: chainOf(sdtXml, ranges, c.start, c.end),
  }))
}

/**
 * When a top-level <w:sdt>'s content begins with a table (no paragraph before
 * it), return the balanced <w:tbl>…</w:tbl> slice, else null.
 */
export function sdtTableXml(sdtXml: string): string | null {
  const contentOpen = /<w:sdtContent(?:\s[^>]*)?>/.exec(sdtXml)
  if (!contentOpen) return null
  const contentTagEnd = contentOpen.index + contentOpen[0].length
  const contentClose = sdtXml.lastIndexOf('</w:sdtContent>')
  if (contentClose <= contentTagEnd) return null
  const inner = sdtXml.slice(contentTagEnd, contentClose)
  const tblStart = inner.search(/<w:tbl[\s>]/)
  if (tblStart === -1) return null
  const pStart = inner.search(/<w:p[\s/>]/)
  if (pStart !== -1 && pStart < tblStart) return null
  // balanced </w:tbl> for the opening tag (tables can nest)
  let depth = 0
  const tagRe = /<\/?w:tbl(?=[\s/>])/g
  tagRe.lastIndex = tblStart
  let m: RegExpExecArray | null
  while ((m = tagRe.exec(inner)) !== null) {
    if (m[0].startsWith('</')) {
      depth--
      if (depth === 0) return inner.slice(tblStart, m.index + '</w:tbl>'.length)
    } else depth++
  }
  return null
}
