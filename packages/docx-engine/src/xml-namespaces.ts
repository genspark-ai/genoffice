/** the prefixes generated markup uses, with the namespace each stands for */
const KNOWN_NAMESPACES: Record<string, string> = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wp14: 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  cx: 'http://schemas.microsoft.com/office/drawing/2014/chartex',
  cx1: 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
  cx2: 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
  cx3: 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex',
  cx4: 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex',
  cx5: 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex',
  cx6: 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex',
  cx7: 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex',
  cx8: 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wpg: 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  wpi: 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
  wne: 'http://schemas.microsoft.com/office/word/2006/wordml',
  aink: 'http://schemas.microsoft.com/office/drawing/2016/ink',
  am3d: 'http://schemas.microsoft.com/office/drawing/2017/model3d',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  w10: 'urn:schemas-microsoft-com:office:word',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
  w15: 'http://schemas.microsoft.com/office/word/2012/wordml',
  w16: 'http://schemas.microsoft.com/office/word/2018/wordml',
  w16cex: 'http://schemas.microsoft.com/office/word/2018/wordml/cex',
  w16cid: 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
  w16du: 'http://schemas.microsoft.com/office/word/2023/wordml/word16du',
  w16sdtdh: 'http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash',
  w16sdtfl: 'http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock',
  w16se: 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
  asvg: 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
}

const SKIPPED = '<!--[\\s\\S]*?-->|<!\\[CDATA\\[[\\s\\S]*?\\]\\]>|<\\?[\\s\\S]*?\\?>'
// comments, CDATA and processing instructions match with no name, so tags inside them are not markup
const TAG = new RegExp(
  `${SKIPPED}|<(\\/?)([A-Za-z_][\\w.-]*)(?::[\\w.-]+)?((?:"[^"]*"|'[^']*'|[^"'>])*)>`,
  'g',
)
const PROLOG = new RegExp(`^(?:\\s*(?:${SKIPPED}))*`)
const ATTR = /\s([A-Za-z_][\w.-]*):([\w.-]+)\s*=/g
// markup-compatibility attributes name prefixes in their value; those must be declared too
const PREFIX_LIST =
  /\s(?:Requires|[A-Za-z_][\w.-]*:(?:Ignorable|MustUnderstand|ProcessContent))\s*=\s*(?:"([^"]*)"|'([^']*)')/g
const ROOT = /<([A-Za-z_][\w.-]*:)?[A-Za-z_][\w.-]*(?:\s(?:"[^"]*"|'[^']*'|[^"'>])*)?>/
const DECLARATION = /\sxmlns:([A-Za-z_][\w.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/g

const rootOf = (xml: string) =>
  ROOT.exec(xml.replace(PROLOG, (prolog) => ' '.repeat(prolog.length)))

/** the prefixes a part's root declares, with their namespaces */
export function rootNamespaces(xml: string): Record<string, string> {
  const out: Record<string, string> = {}
  for (const m of rootOf(xml)?.[0].matchAll(DECLARATION) ?? []) out[m[1]!] = m[2] ?? m[3] ?? ''
  return out
}

/**
 * Markup inserted into a part (a chart, a picture, a shape) uses prefixes such
 * as `wp:` and `r:` that a Word-saved file declares on its root. A file from
 * another generator may not, a part the save creates starts from a short root
 * of its own, and one unbound prefix makes the whole part unparseable: Word
 * refuses to open it. Every known prefix that is used outside the scope of a
 * declaration gets declared on the root; a prefix that is only used where an
 * ancestor declares it is left alone, so the root of a file that needs nothing
 * keeps its bytes.
 *
 * `donors` are the declarations of the part the content may have come from
 * (the main document's root, for a header or a note): they bind a prefix this
 * table does not know.
 */
export function bindUnboundPrefixes(xml: string, donors: Record<string, string> = {}): string {
  const rootMatch = rootOf(xml)
  if (!rootMatch) return xml
  const root = rootMatch[0]
  const namespaces: Record<string, string> = { ...donors, ...KNOWN_NAMESPACES }
  const listed = new Set<string>()
  if (xml.includes('Requires=') || xml.includes(':Ignorable=')) {
    for (const m of xml.matchAll(PREFIX_LIST))
      for (const token of (m[1] ?? m[2] ?? '').split(/\s+/)) listed.add(token.split(':')[0]!)
  }
  // only prefixes the root does not declare can be unbound; most files stop here
  const candidates = Object.keys(namespaces).filter(
    (p) =>
      !root.includes(`xmlns:${p}=`) &&
      (xml.includes(`<${p}:`) || xml.includes(` ${p}:`) || listed.has(p)),
  )
  if (candidates.length === 0) return xml
  const wanted = new Set(candidates)
  const unbound = new Set<string>()
  const scopes: string[][] = []
  const bound = (p: string) => scopes.some((declared) => declared.includes(p))
  TAG.lastIndex = rootMatch.index
  let tag: RegExpExecArray | null
  while ((tag = TAG.exec(xml)) !== null && unbound.size < wanted.size) {
    if (tag[2] === undefined) continue
    if (tag[1] === '/') {
      scopes.pop()
      continue
    }
    const attrs = tag[3] ?? ''
    const declared: string[] = []
    const used: string[] = []
    const prefixed = tag[0][1 + tag[2]!.length] === ':'
    if (prefixed && wanted.has(tag[2]!)) used.push(tag[2]!)
    ATTR.lastIndex = 0
    let attr: RegExpExecArray | null
    while ((attr = ATTR.exec(attrs)) !== null) {
      if (attr[1] === 'xmlns') declared.push(attr[2]!)
      else if (wanted.has(attr[1]!)) used.push(attr[1]!)
    }
    if (listed.size > 0 && (attrs.includes('Requires') || attrs.includes(':Ignorable'))) {
      for (const m of attrs.matchAll(PREFIX_LIST))
        for (const token of (m[1] ?? m[2] ?? '').split(/\s+/)) {
          const p = token.split(':')[0]!
          if (wanted.has(p)) used.push(p)
        }
    }
    for (const p of used) if (!declared.includes(p) && !bound(p)) unbound.add(p)
    if (!attrs.trimEnd().endsWith('/')) scopes.push(declared)
  }
  if (unbound.size === 0) return xml
  const declarations = [...unbound].map((p) => ` xmlns:${p}="${namespaces[p]}"`).join('')
  const at = rootMatch.index + root.length - 1
  return xml.slice(0, at) + declarations + xml.slice(at)
}
