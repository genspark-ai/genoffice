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

/**
 * An ST_Percentage as a fraction of one. The schema's own form is an integer
 * in 1/1000 % ("50000"), and Google Slides writes the suffixed form ("50%").
 * Reading the suffix as a bare integer turns every such value into a
 * thousandth of what the file meant, so both forms are accepted here.
 */
export function percentFraction(raw: unknown): number | undefined {
  if (raw == null) return undefined
  const text = String(raw).trim()
  const n = Number.parseFloat(text)
  if (!Number.isFinite(n)) return undefined
  return text.endsWith('%') ? n / 100 : n / 100000
}

/** The same value in percent; one division keeps whole percents exact. */
export function percentPoints(raw: unknown): number | undefined {
  if (raw == null) return undefined
  const text = String(raw).trim()
  const n = Number.parseFloat(text)
  if (!Number.isFinite(n)) return undefined
  return text.endsWith('%') ? n : n / 1000
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

/**
 * An `<a:srgbClr val="…">` value as `#RRGGBB`. ST_HexColorRGB is six hex
 * digits; anything else is a broken attribute, and handing it to a lenient
 * hex reader invents a colour ('#UNDEFINED' reads as 00 DE 0F — bright green).
 * An eight-digit value is accepted and its first six digits are the colour.
 */
export function srgbHexValue(raw: unknown): string | undefined {
  const text = String(raw ?? '').trim()
  return /^[0-9a-fA-F]{6}([0-9a-fA-F]{2})?$/.test(text) ? '#' + text.toUpperCase() : undefined
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

const DEFAULT_TAG = /<Default\b[^>]*\/?>/g

/** Any <Default> already declaring this extension, whatever its attribute order or quoting. */
export function hasDefaultFor(ct: string, ext: string): boolean {
  const re = new RegExp(`\\bExtension\\s*=\\s*(["'])${ext}\\1`)
  for (const tag of ct.matchAll(DEFAULT_TAG)) if (re.test(tag[0])) return true
  return false
}

/**
 * Prefix PowerPoint (and this engine's literal `p:`/`a:`/`r:` lookups) expect for each OOXML
 * namespace. `''` marks namespaces that are conventionally the default namespace of their part.
 */
const CANONICAL_NS_PREFIX: Record<string, string> = {
  'http://schemas.openxmlformats.org/presentationml/2006/main': 'p',
  'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
  'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing': 'cdr',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'dgm',
  'http://schemas.openxmlformats.org/drawingml/2006/picture': 'pic',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'http://schemas.microsoft.com/office/drawing/2014/chartex': 'cx',
  'http://schemas.microsoft.com/office/drawing/2008/diagram': 'dsp',
  'http://schemas.microsoft.com/office/drawing/2010/main': 'a14',
  'http://schemas.microsoft.com/office/drawing/2014/main': 'a16',
  'http://schemas.microsoft.com/office/drawing/2016/SVG/main': 'asvg',
  'http://schemas.microsoft.com/office/drawing/2017/model3d': 'am3d',
  'http://schemas.microsoft.com/office/drawing/2012/chart': 'c15',
  'http://schemas.microsoft.com/office/powerpoint/2010/main': 'p14',
  'http://schemas.microsoft.com/office/powerpoint/2012/main': 'p15',
  'http://schemas.microsoft.com/office/powerpoint/2015/main': 'p159',
  'urn:schemas-microsoft-com:vml': 'v',
  'http://schemas.openxmlformats.org/package/2006/relationships': '',
  'http://schemas.openxmlformats.org/package/2006/content-types': '',
}

const ROOT_START_TAG = /<(?!\?|!)[^>]*>/
const XMLNS_DECL = /\sxmlns(?::([A-Za-z_][\w.-]*))?=(["'])([^"']*)\2/g

/**
 * Rewrite namespace prefixes to the ones PowerPoint writes. Tools that re-serialize a part
 * through a generic XML library (Python ElementTree writes `ns0:`, `ns1:`, …) keep the XML
 * namespace-correct, and PowerPoint opens such decks, but every reader here addresses
 * elements by literal prefix and would see an empty presentation. Returns the input unchanged
 * when every declared prefix is already canonical or the rewrite would collide with an
 * existing declaration.
 */
export function canonicalizeNsPrefixes(xml: string): string {
  const root = ROOT_START_TAG.exec(xml)
  if (!root) return xml
  const declared = new Map<string, string>()
  for (const m of root[0].matchAll(XMLNS_DECL)) declared.set(m[1] ?? '', m[3]!)
  const renames = new Map<string, string>()
  for (const [prefix, uri] of declared) {
    const want = CANONICAL_NS_PREFIX[uri]
    if (want === undefined || want === prefix || prefix === '') continue
    if (declared.has(want) && declared.get(want) !== uri) return xml
    renames.set(prefix, want)
  }
  if (!renames.size) return xml
  const alt = [...renames.keys()].map((p) => p.replace(/[.-]/g, '\\$&')).join('|')
  const decl = new RegExp(`(\\sxmlns):(${alt})=`, 'g')
  const names = new RegExp(`(<\\/?|\\s)(${alt}):(?=[A-Za-z_])`, 'g')
  return xml
    .replace(decl, (_, lead: string, p: string) => {
      const want = renames.get(p)!
      return want ? `${lead}:${want}=` : `${lead}=`
    })
    .replace(names, (_, lead: string, p: string) => {
      const want = renames.get(p)!
      return want ? `${lead}${want}:` : lead
    })
}
