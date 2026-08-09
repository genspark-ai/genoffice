import { XMLParser } from 'fast-xml-parser'

/** preserveOrder node shape from fast-xml-parser */
export type XNode = Record<string, unknown>

export const xmlParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
})

export function nameOf(node: XNode): string | undefined {
  return Object.keys(node).find((k) => k !== ':@' && k !== '#text')
}

export function childrenOf(node: XNode): XNode[] {
  const name = nameOf(node)
  if (!name) return []
  const value = node[name]
  return Array.isArray(value) ? (value as XNode[]) : []
}

export function attrsOf(node: XNode): Record<string, string> {
  return (node[':@'] as Record<string, string>) ?? {}
}

export function textOf(node: XNode): string {
  let out = ''
  for (const child of childrenOf(node)) {
    if ('#text' in child) out += String(child['#text'])
    else out += textOf(child)
  }
  return out
}

export function findChild(node: XNode, name: string): XNode | undefined {
  return childrenOf(node).find((c) => nameOf(c) === name)
}

export function findChildren(node: XNode, name: string): XNode[] {
  return childrenOf(node).filter((c) => nameOf(c) === name)
}

/**
 * Direct children with `name`, looking through w:sdt → w:sdtContent wrappers
 * (nested sdt included). Structured document tags may wrap table rows, cells
 * or paragraphs at any level; for display purposes the wrapper is transparent
 * (research-report templates wrap every field in an sdt).
 */
export function childrenThroughSdt(node: XNode, name: string | readonly string[]): XNode[] {
  const names = Array.isArray(name) ? (name as readonly string[]) : [name as string]
  const out: XNode[] = []
  const visit = (n: XNode): void => {
    for (const child of childrenOf(n)) {
      const cn = nameOf(child)
      if (cn !== undefined && names.includes(cn)) out.push(child)
      else if (cn === 'w:sdt') {
        const content = findChild(child, 'w:sdtContent')
        if (content) visit(content)
      }
    }
  }
  visit(node)
  return out
}

/** OOXML boolean property: present => true unless w:val says otherwise */
export function boolProp(parent: XNode, name: string): boolean {
  const child = findChild(parent, name)
  if (!child) return false
  const val = attrsOf(child)['w:val']
  if (val === undefined) return true
  return !['0', 'false', 'none', 'off'].includes(val.toLowerCase())
}

/**
 * w:u is NOT an OOXML boolean (CT_OnOff) — it is CT_Underline, where the
 * underline pattern lives entirely in w:val. A <w:u> with no w:val (e.g.
 * `<w:u w:color="415461"/>` as emitted by Pages/LibreOffice) means no
 * underline, matching how Word renders it.
 */
export function underlineProp(parent: XNode): boolean {
  const child = findChild(parent, 'w:u')
  if (!child) return false
  const val = attrsOf(child)['w:val']
  return val !== undefined && val !== 'none'
}

/**
 * XNode → XML text (attribute order = parse order, empty elements self-close). Semantic
 * fidelity, not byte fidelity: used to store parse-tree fragments (e.g. a run's rPr) as
 * writable source slices.
 */
export function serializeXNode(node: XNode): string {
  if ('#text' in node) return escapeXmlText(String(node['#text']))
  const name = nameOf(node)
  if (!name) return ''
  const attrs = Object.entries(attrsOf(node))
    .map(([k, v]) => ` ${k}="${escapeXmlAttr(String(v))}"`)
    .join('')
  const inner = childrenOf(node).map(serializeXNode).join('')
  return inner === '' ? `<${name}${attrs}/>` : `<${name}${attrs}>${inner}</${name}>`
}

// Control characters outside \t \n \r are illegal in XML 1.0 even when escaped
// eslint-disable-next-line no-control-regex
const ILLEGAL_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g

export function escapeXmlText(text: string): string {
  return text
    .replace(ILLEGAL_XML_CHARS, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

export function escapeXmlAttr(text: string): string {
  return escapeXmlText(text).replace(/"/g, '&quot;')
}
