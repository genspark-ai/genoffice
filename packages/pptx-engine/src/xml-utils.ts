/** XML escaping utilities (same as docx-engine's, used for patch generation). */

/**
 * A fast-xml-parser element node: attributes are '@_'-prefixed, text is '#text',
 * child elements are nested nodes (a single child collapses to an object instead
 * of an array). Parse trees are inherently untyped, so readers probe them through
 * the narrowing helpers below instead of `any`.
 */
export type XmlNode = Record<string, unknown>

/** View an unknown parse-tree value as an element node; non-objects read as empty. */
export function asXmlNode(v: unknown): XmlNode {
  return typeof v === 'object' && v !== null ? (v as XmlNode) : {}
}

/** Normalize fast-xml-parser's single-child collapse: always get an array of element nodes. */
export function xmlArray(v: unknown): XmlNode[] {
  if (Array.isArray(v)) return v.map(asXmlNode)
  return v ? [asXmlNode(v)] : []
}

/** Decode numeric references only when they name a Unicode scalar value. */
export function decodeNumericCharRefs(text: string): string {
  return text.replace(/&#(?:x([0-9a-fA-F]+)|(\d+));/g, (reference, hex, decimal) => {
    const code = hex === undefined ? Number(decimal) : Number.parseInt(hex, 16)
    return code <= 0x10ffff && (code < 0xd800 || code > 0xdfff)
      ? String.fromCodePoint(code)
      : reference
  })
}

/** XML 1.0 forbids C0 controls (minus tab/LF/CR), U+FFFE/FFFF and lone
    surrogates even when escaped — one such byte makes the whole part
    unparseable and PowerPoint offers repair. VT/FF (common in PDF-extracted
    text) degrade to a space; the rest have no textual meaning and drop. */
// eslint-disable-next-line no-control-regex -- the forbidden chars are the subject
const XML_INVALID_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|\p{Cs}/gu

function sanitizeXmlChars(text: string): string {
  XML_INVALID_CHARS.lastIndex = 0
  if (!XML_INVALID_CHARS.test(text)) return text
  XML_INVALID_CHARS.lastIndex = 0
  return text.replace(XML_INVALID_CHARS, (ch) => (ch === '\u000B' || ch === '\u000C' ? ' ' : ''))
}

export function escapeXmlText(text: string): string {
  return sanitizeXmlChars(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

export function escapeXmlAttr(text: string): string {
  return escapeXmlText(text).replace(/"/g, '&quot;')
}

const RELATIONSHIP_ID = /\bId\s*=\s*(["'])rId(\d+)\1/g

/**
 * Highest rIdN already present in a .rels part. The Id attribute is read
 * quote-agnostically, so a rels file that spells its ids with single quotes
 * still raises the counter instead of handing out a duplicate id.
 */
export function maxRelationshipIdNumber(relsXml: string): number {
  let max = 0
  for (const match of relsXml.matchAll(RELATIONSHIP_ID)) {
    const n = Number(match[2])
    if (n > max) max = n
  }
  return max
}

const PART_NAME = /\bPartName\s*=\s*(["'])(.*?)\1/g

/**
 * Whether [Content_Types].xml already declares an Override for the part.
 * Quote-agnostic for the same reason as maxRelationshipIdNumber: an existing
 * single-quoted PartName must be recognized so no second Override is emitted
 * for the same part (PowerPoint offers to repair such a package).
 */
export function hasContentTypeOverride(contentTypes: string, partPath: string): boolean {
  const wanted = `/${partPath}`
  for (const match of contentTypes.matchAll(PART_NAME)) {
    if (match[2] === wanted) return true
  }
  return false
}

/**
 * a16:creationId extLst for a newborn <p:cNvPr> — durable identity from birth
 * (design step 0): the GUID is written into the file bytes, so ids survive
 * save→reopen, reparse, group/ungroup, and editors that renumber cNvPr ids.
 * The a16 namespace is declared inline so the fragment is valid standalone.
 */
export function creationIdExtXml(): string {
  return (
    '<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">' +
    `<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{${globalThis.crypto.randomUUID().toUpperCase()}}"/>` +
    '</a:ext>'
  )
}

export function creationIdXml(): string {
  return `<a:extLst>${creationIdExtXml()}</a:extLst>`
}
