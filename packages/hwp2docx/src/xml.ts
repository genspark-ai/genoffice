/**
 * The ordered XML tree every HWPX part is read through (fast-xml-parser's
 * preserveOrder mode). OWPML paragraphs interleave text with inline controls
 * (tabs, line breaks, tables, pictures), so document order must survive.
 * Matching is on the prefixes Hancom writes (`hp:`, `hh:`, `hc:`, `hs:`,
 * `opf:`) — every producer in the wild (Hangul 2014+, Hancom Office, the
 * public converters) declares exactly these.
 */

import { XMLParser } from 'fast-xml-parser'

/** one preserveOrder node: `{'hp:p': children, ':@': attrs}` or `{'#text': '…'}` */
export type XNode = Record<string, unknown>

const parser = new XMLParser({
  preserveOrder: true,
  ignorePiTags: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  // <hp:t> keeps its spaces verbatim
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: true,
  commentPropName: '#comment',
  cdataPropName: '#cdata',
})

/** Parse one part; null on malformed XML (the caller degrades, never throws). */
export function parseXml(xml: string): XNode[] | null {
  try {
    const out = parser.parse(xml)
    return Array.isArray(out) ? out : null
  } catch {
    return null
  }
}

export function nodeName(node: XNode): string | null {
  for (const key of Object.keys(node)) {
    if (key !== ':@' && key !== '#text' && key !== '#comment' && key !== '#cdata') return key
  }
  return null
}

export function attrs(node: XNode): Record<string, string> {
  return (node[':@'] as Record<string, string> | undefined) ?? {}
}

export function attr(node: XNode | null | undefined, name: string): string | undefined {
  return node ? attrs(node)[name] : undefined
}

export function children(node: XNode): XNode[] {
  const name = nodeName(node)
  if (!name) return []
  const value = node[name]
  return Array.isArray(value) ? (value as XNode[]) : []
}

export function textOf(node: XNode): string | null {
  const cdata = node['#cdata']
  if (typeof cdata === 'string') return cdata
  const text = node['#text']
  return typeof text === 'string' ? text : null
}

/** element children only, as {name, node} pairs */
export function elements(node: XNode | XNode[]): Array<{ name: string; node: XNode }> {
  const list = Array.isArray(node) ? node : children(node)
  const out: Array<{ name: string; node: XNode }> = []
  for (const child of list) {
    const name = nodeName(child)
    if (name && !name.startsWith('#')) out.push({ name, node: child })
  }
  return out
}

export function firstChild(node: XNode | null | undefined, name: string): XNode | null {
  if (!node) return null
  for (const child of children(node)) if (nodeName(child) === name) return child
  return null
}

/** depth-first search for the first element with this name */
export function firstDescendant(node: XNode, name: string): XNode | null {
  for (const child of children(node)) {
    if (nodeName(child) === name) return child
    const deeper = firstDescendant(child, name)
    if (deeper) return deeper
  }
  return null
}

/** every element with this name below `node` (document order) */
export function descendants(node: XNode | XNode[], name: string, out: XNode[] = []): XNode[] {
  const list = Array.isArray(node) ? node : children(node)
  for (const child of list) {
    if (nodeName(child) === name) out.push(child)
    descendants(child, name, out)
  }
  return out
}

/**
 * `<hp:switch><hp:case hp:required-namespace="…/HwpUnitChar">A</hp:case>
 * <hp:default>B</hp:default></hp:switch>`: Hangul writes paragraph margins
 * twice — the HwpUnitChar case in true HWPUNIT, the default branch doubled
 * for legacy readers. The laid-out line positions (hp:lineseg horzpos) match
 * the case values, so a case this filter understands wins.
 */
export function unwrapSwitch(node: XNode): XNode[] {
  const out: XNode[] = []
  for (const child of children(node)) {
    if (nodeName(child) === 'hp:switch') {
      const known = children(child).find(
        (branch) =>
          nodeName(branch) === 'hp:case' &&
          /HwpUnitChar$/.test(attr(branch, 'hp:required-namespace') ?? ''),
      )
      const chosen = known ?? firstChild(child, 'hp:default') ?? firstChild(child, 'hp:case')
      if (chosen) out.push(...children(chosen))
    } else {
      out.push(child)
    }
  }
  return out
}

export function escapeXml(text: string): string {
  return text.replace(/[&<>"]/g, (c) =>
    c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;',
  )
}

/** integer attribute; NaN-safe */
export function intAttr(node: XNode | null | undefined, name: string, fallback = 0): number {
  const value = Number(attr(node, name))
  return Number.isFinite(value) ? value : fallback
}
