/**
 * A body block of one document made fit to live in another: the raw XML of a
 * shape, text box, equation or field paragraph, with the namespace
 * declarations it took from its own document's root.
 */
import { loadDocxZip } from './zip-load'

const RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const XMLNS = /\sxmlns:([A-Za-z_][\w.-]*)="([^"]*)"/g
/** content that points at another part or story of its document, which does not come along */
const TIED =
  /<w:(?:footnoteReference|endnoteReference|commentReference|commentRangeStart|commentRangeEnd|sectPr|altChunk|object)\b/

/** prefix -> namespace of the declarations on the root element of word/document.xml */
export async function docxRootNamespaces(bytes: Uint8Array): Promise<Record<string, string>> {
  const xml = await (await loadDocxZip(bytes)).file('word/document.xml')?.async('string')
  const root = xml ? /<w:document\b[^>]*>/.exec(xml)?.[0] : null
  const out: Record<string, string> = {}
  for (const m of root?.matchAll(XMLNS) ?? []) out[m[1]] = m[2]
  return out
}

/**
 * The block as a fragment that stands on its own, or null when it refers to
 * something outside itself (a relationship to a picture, chart or link, a
 * note, a comment, a section). List numbering and bookmarks belong to the
 * document they were made in and are taken off.
 */
export function carriableBlockXml(xml: string, namespaces: Record<string, string>): string | null {
  if (!/^<w:p[\s>]/.test(xml) || TIED.test(xml)) return null
  const declared = { ...namespaces }
  for (const m of xml.matchAll(XMLNS)) declared[m[1]] = m[2]
  const relPrefixes = Object.keys(declared).filter((p) => declared[p] === RELATIONSHIPS)
  if (relPrefixes.some((p) => new RegExp(`\\s${p}:[\\w.-]+=`).test(xml))) return null

  const used = new Set<string>()
  for (const m of xml.matchAll(/<\/?([A-Za-z_][\w.-]*):/g)) used.add(m[1])
  for (const m of xml.matchAll(/\s([A-Za-z_][\w.-]*):[\w.-]+=/g)) used.add(m[1])
  const openEnd = xml.indexOf('>')
  const open = xml.slice(0, openEnd)
  let added = ''
  for (const prefix of used) {
    if (prefix === 'xmlns' || prefix === 'xml') continue
    const uri = namespaces[prefix]
    if (!uri) {
      // declared nowhere this fragment can see: it would not parse
      if (!declared[prefix]) return null
      continue
    }
    if (!new RegExp(`\\sxmlns:${prefix}=`).test(open)) added += ` xmlns:${prefix}="${uri}"`
  }
  const selfClosing = open.endsWith('/')
  const head = selfClosing ? `${open.slice(0, -1)}${added}/` : `${open}${added}`
  return (head + xml.slice(openEnd))
    .replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, '')
    .replace(/<w:bookmark(?:Start|End)\b[^>]*\/>/g, '')
}
