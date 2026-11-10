/**
 * The ordered XML tree every ODF part is read through (fast-xml-parser's
 * preserveOrder mode). Nodes keep document order and mixed content intact,
 * which is what ODF text bodies are: a paragraph is text interleaved with
 * styled spans, tabs, breaks. Matching is on the conventional ODF namespace
 * prefixes (`text:`, `office:`, `table:`, `draw:`, `style:`, `fo:`, `svg:`,
 * `presentation:`, `xlink:`) as written — every producer in the wild
 * (LibreOffice, Word, Pandoc) declares exactly these.
 */

import { XMLParser } from 'fast-xml-parser'

/** one preserveOrder node: `{'text:p': children, ':@': attrs}` or `{'#text': '…'}` */
export type XNode = Record<string, unknown>

const parser = new XMLParser({
  preserveOrder: true,
  ignorePiTags: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  // ODF text nodes carry meaningful whitespace; multiple spaces arrive as
  // <text:s> elements, single spaces as literal text
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  processEntities: true,
  commentPropName: '#comment',
  cdataPropName: '#cdata',
})

/** Parse an ODF part; null on malformed XML (the caller degrades, never throws). */
export function parseOdfXml(xml: string): XNode[] | null {
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

export function nodeAttrs(node: XNode): Record<string, string> {
  return (node[':@'] as Record<string, string> | undefined) ?? {}
}

export function nodeChildren(node: XNode): XNode[] {
  const name = nodeName(node)
  if (!name) return []
  const value = node[name]
  return Array.isArray(value) ? (value as XNode[]) : []
}

export function isTextNode(node: XNode): boolean {
  return typeof node['#text'] === 'string' || typeof node['#cdata'] === 'string'
}

export function textOfNode(node: XNode): string {
  const cdata = node['#cdata']
  if (typeof cdata === 'string') return cdata
  const text = node['#text']
  if (typeof text === 'string') return text
  let out = ''
  for (const child of nodeChildren(node)) out += textOfNode(child)
  return out
}

/** First child element with the given name (any level is NOT searched — direct children only). */
export function firstChild(node: XNode, name: string): XNode | null {
  for (const child of nodeChildren(node)) {
    if (nodeName(child) === name) return child
  }
  return null
}

/** Attribute read that tolerates the namespace-less spelling some producers emit. */
export function attr(node: XNode, name: string): string | undefined {
  const value = nodeAttrs(node)[name]
  if (value !== undefined) return value
  const local = name.slice(name.indexOf(':') + 1)
  return nodeAttrs(node)[local]
}
