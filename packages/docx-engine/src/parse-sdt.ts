// Structured document tags (w:sdt): block-level SDT classification and the
// split of an SDT's run/table content.
import type { SdtShell } from './types'

/**
 * Extract a SdtShell from a raw <w:sdt>…</w:sdt> XML string.
 * Returns the shell metadata + the first <w:p> found in <w:sdtContent>,
 * or null if no usable paragraph is found.
 */
export function parseSdtBlock(sdtXml: string): { shell: SdtShell; pXml: string } | null {
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

  return {
    shell: { ...sdtMeta(sdtXml), openXml, closeXml },
    pXml,
  }
}

export function sdtMeta(sdtXml: string): Pick<SdtShell, 'alias' | 'tag' | 'controlType'> {
  const sdtPrXml = /<w:sdtPr>([\s\S]*?)<\/w:sdtPr>/.exec(sdtXml)?.[1] ?? ''
  const alias = /w:val="([^"]*)"/.exec(/<w:alias[^>]*>/.exec(sdtPrXml)?.[0] ?? '')?.[1] ?? ''
  const tag = /w:val="([^"]*)"/.exec(/<w:tag[^>]*>/.exec(sdtPrXml)?.[0] ?? '')?.[1] ?? ''
  let controlType: SdtShell['controlType'] = 'text'
  if (/<w:date[\s/>]/.test(sdtPrXml)) controlType = 'date'
  else if (/<w:dropDownList[\s/>]|<w:comboBox[\s/>]/.test(sdtPrXml)) controlType = 'dropdown'
  else if (/<w:checkbox[\s/>]/.test(sdtPrXml)) controlType = 'checkbox'
  else if (/<w:text[\s/>]|<w:richText[\s/>]/.test(sdtPrXml)) controlType = 'text'
  return { alias, tag, controlType }
}

interface SdtPart {
  name: string
  /** slice of the whole <w:sdt> xml owned by this part (parts partition the sdt exactly) */
  start: number
  end: number
  /** the w:p / w:tbl child inside [start, end) */
  childStart: number
  childEnd: number
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
  return parts.map((c, k) => ({
    name: c.name,
    start: k === 0 ? 0 : c.start,
    end: k === parts.length - 1 ? sdtXml.length : parts[k + 1].start,
    childStart: c.start,
    childEnd: c.end,
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
