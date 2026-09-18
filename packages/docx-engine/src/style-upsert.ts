import { PPR_CHILD_ORDER, RPR_CHILD_ORDER, splitXmlChildren } from './generate'
import { escapeXmlAttr } from './xml-utils'

/** Run properties of a style definition; a present key sets, null clears, absent keeps. */
export interface StyleRunProps {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  caps?: boolean
  smallCaps?: boolean
  /** hex without '#' */
  color?: string | null
  /** w:highlight color name (yellow, green, ...) */
  highlight?: string | null
  sizeHalfPoints?: number | null
  /** Latin face (w:ascii / w:hAnsi); theme font attributes are dropped so it takes effect */
  font?: string | null
  /** East Asian face (w:eastAsia) */
  eastAsiaFont?: string | null
}

/** Paragraph properties of a style definition; same patch semantics as StyleRunProps. */
export interface StyleParaProps {
  align?: 'left' | 'center' | 'right' | 'justify' | null
  spaceBeforeTwips?: number | null
  spaceAfterTwips?: number | null
  /** line spacing as a multiple (auto) */
  lineSpacing?: number | null
  indentLeftTwips?: number | null
  indentRightTwips?: number | null
  /** positive = first-line indent, negative = hanging indent */
  firstLineTwips?: number | null
  keepNext?: boolean
  keepLines?: boolean
  pageBreakBefore?: boolean
  /** 1-9 (w:outlineLvl val + 1) */
  outlineLevel?: number | null
}

/**
 * Create or modify one style in word/styles.xml. An existing definition is
 * patched property by property: children the patch does not name (link,
 * uiPriority, rsid, tblPr, ...) keep their bytes, and w:spacing / w:ind /
 * w:rFonts keep the attributes the patch does not set.
 */
export interface StyleHeadingInfo {
  name?: string
  basedOn?: string
  headingLevel?: number
  headingLevelInherited?: boolean
  /** own w:outlineLvl 9 (body text); kept by a patch that does not set outlineLevel */
  headingOutlineOff?: boolean
}

/** parse-styles reads "heading N" names and HeadingN ids as that level, before any w:outlineLvl */
function headingLevelByName(name: string | undefined, styleId: string): number | undefined {
  const m = (name && /^heading\s*([1-9])$/i.exec(name)) || /^Heading([1-9])$/.exec(styleId)
  return m ? parseInt(m[1], 10) : undefined
}

/**
 * Heading level a pending paragraph style resolves to once saved, mirroring parse-styles on
 * reopen: heading name or id, else its own outline level (`null` or an existing w:outlineLvl 9
 * = body text), else the nearest basedOn ancestor's. An existing style keeps what it has unless the upsert moves
 * an inherited level to a new parent.
 */
export function pendingHeadingLevel(
  styleId: string,
  pending: (id: string) => StyleUpsert | undefined,
  existing: (id: string) => StyleHeadingInfo | undefined,
  seen: Set<string> = new Set(),
): number | undefined {
  if (seen.has(styleId)) return undefined
  seen.add(styleId)
  const up = pending(styleId)
  const cur = existing(styleId)
  if (!up) return cur?.headingLevel
  const named = headingLevelByName(up.name ?? cur?.name, styleId)
  if (named) return named
  const outline = up.pPr?.outlineLevel
  if (outline !== undefined) return outline ?? undefined
  if (cur?.headingOutlineOff) return undefined
  if (cur && (up.basedOn === undefined || (cur.headingLevel && !cur.headingLevelInherited)))
    return cur.headingLevel
  const parent = up.basedOn === undefined ? cur?.basedOn : (up.basedOn ?? undefined)
  return parent ? pendingHeadingLevel(parent, pending, existing, seen) : undefined
}

export interface StyleUpsert {
  styleId: string
  /** used when creating (default paragraph); an existing style keeps its type */
  type?: 'paragraph' | 'character' | 'table'
  /** display name; defaults to styleId when creating */
  name?: string
  /** null removes w:basedOn */
  basedOn?: string | null
  /** style of the following paragraph; null removes w:next */
  next?: string | null
  /** w:qFormat (quick style gallery); new styles default to true */
  quickFormat?: boolean
  rPr?: StyleRunProps
  pPr?: StyleParaProps
}

const STYLE_CHILD_ORDER = [
  'w:name',
  'w:aliases',
  'w:basedOn',
  'w:next',
  'w:link',
  'w:autoRedefine',
  'w:hidden',
  'w:uiPriority',
  'w:semiHidden',
  'w:unhideWhenUsed',
  'w:qFormat',
  'w:locked',
  'w:personal',
  'w:personalCompose',
  'w:personalReply',
  'w:rsid',
  'w:pPr',
  'w:rPr',
  'w:tblPr',
  'w:trPr',
  'w:tcPr',
  'w:tblStylePr',
]

interface Child {
  name: string
  xml: string
}

type Attrs = Map<string, string>

function parseTag(xml: string): { name: string; attrs: Attrs; selfClosing: boolean } {
  const m = /^<([A-Za-z0-9:._-]+)((?:\s+[^\s=>]+="[^"]*")*)\s*(\/?)>/.exec(xml)
  if (!m) throw new Error(`style-upsert: not an element: ${xml.slice(0, 40)}`)
  const attrs: Attrs = new Map()
  for (const a of m[2]!.matchAll(/([^\s=>]+)="([^"]*)"/g)) attrs.set(a[1]!, a[2]!)
  return { name: m[1]!, attrs, selfClosing: m[3] === '/' }
}

function tag(name: string, attrs: Attrs, inner?: string): string {
  const a = [...attrs].map(([k, v]) => ` ${k}="${v}"`).join('')
  return inner === undefined ? `<${name}${a}/>` : `<${name}${a}>${inner}</${name}>`
}

function innerOf(xml: string): string {
  const { selfClosing } = parseTag(xml)
  if (selfClosing) return ''
  return xml.slice(xml.indexOf('>') + 1, xml.lastIndexOf('</'))
}

/** ordered child list with single-instance replace/remove */
class Children {
  constructor(
    private items: Child[],
    private order: readonly string[],
  ) {}

  get(name: string): Child | undefined {
    return this.items.find((c) => c.name === name)
  }

  set(name: string, xml: string | null): void {
    const at = this.items.findIndex((c) => c.name === name)
    if (xml === null) {
      if (at >= 0) this.items.splice(at, 1)
    } else if (at >= 0) this.items[at] = { name, xml }
    else this.items.push({ name, xml })
  }

  /** self-closing on/off element: true = present, false = w:val="0" */
  flag(name: string, on: boolean | undefined): void {
    if (on === undefined) return
    this.set(name, on ? `<${name}/>` : `<${name} w:val="0"/>`)
  }

  /** merge attributes into a self-closing element; null attribute values remove */
  attrs(name: string, patch: Record<string, string | null>): void {
    const cur = this.get(name)
    const attrs = cur ? parseTag(cur.xml).attrs : new Map<string, string>()
    for (const [k, v] of Object.entries(patch)) {
      if (v === null) attrs.delete(k)
      else attrs.set(k, v)
    }
    this.set(name, attrs.size === 0 ? null : tag(name, attrs))
  }

  toXml(): string {
    const rank = (n: string) => {
      const i = this.order.indexOf(n)
      return i < 0 ? this.order.length : i
    }
    return [...this.items]
      .sort((a, b) => rank(a.name) - rank(b.name))
      .map((c) => c.xml)
      .join('')
  }

  get size(): number {
    return this.items.length
  }
}

const num = (n: number) => String(Math.round(n))

function patchRun(children: Children, rp: StyleRunProps): void {
  children.flag('w:b', rp.bold)
  children.flag('w:bCs', rp.bold)
  children.flag('w:i', rp.italic)
  children.flag('w:iCs', rp.italic)
  children.flag('w:strike', rp.strike)
  children.flag('w:caps', rp.caps)
  children.flag('w:smallCaps', rp.smallCaps)
  if (rp.underline !== undefined)
    children.set('w:u', `<w:u w:val="${rp.underline ? 'single' : 'none'}"/>`)
  if (rp.color !== undefined) {
    children.set(
      'w:color',
      rp.color === null ? null : `<w:color w:val="${escapeXmlAttr(rp.color)}"/>`,
    )
  }
  if (rp.highlight !== undefined) {
    children.set(
      'w:highlight',
      rp.highlight === null ? null : `<w:highlight w:val="${escapeXmlAttr(rp.highlight)}"/>`,
    )
  }
  if (rp.sizeHalfPoints !== undefined) {
    const sz = rp.sizeHalfPoints === null ? null : num(rp.sizeHalfPoints)
    children.set('w:sz', sz === null ? null : `<w:sz w:val="${sz}"/>`)
    children.set('w:szCs', sz === null ? null : `<w:szCs w:val="${sz}"/>`)
  }
  const fonts: Record<string, string | null> = {}
  if (rp.font !== undefined) {
    const f = rp.font === null ? null : escapeXmlAttr(rp.font)
    Object.assign(fonts, {
      'w:ascii': f,
      'w:hAnsi': f,
      'w:asciiTheme': null,
      'w:hAnsiTheme': null,
    })
  }
  if (rp.eastAsiaFont !== undefined) {
    fonts['w:eastAsia'] = rp.eastAsiaFont === null ? null : escapeXmlAttr(rp.eastAsiaFont)
    fonts['w:eastAsiaTheme'] = null
  }
  if (Object.keys(fonts).length > 0) children.attrs('w:rFonts', fonts)
}

function patchPara(children: Children, pp: StyleParaProps): void {
  if (pp.align !== undefined) {
    children.set(
      'w:jc',
      pp.align === null ? null : `<w:jc w:val="${pp.align === 'justify' ? 'both' : pp.align}"/>`,
    )
  }
  const spacing: Record<string, string | null> = {}
  if (pp.spaceBeforeTwips !== undefined)
    spacing['w:before'] = pp.spaceBeforeTwips === null ? null : num(pp.spaceBeforeTwips)
  if (pp.spaceAfterTwips !== undefined)
    spacing['w:after'] = pp.spaceAfterTwips === null ? null : num(pp.spaceAfterTwips)
  if (pp.lineSpacing !== undefined) {
    spacing['w:line'] = pp.lineSpacing === null ? null : num(pp.lineSpacing * 240)
    spacing['w:lineRule'] = pp.lineSpacing === null ? null : 'auto'
  }
  if (Object.keys(spacing).length > 0) children.attrs('w:spacing', spacing)
  const ind: Record<string, string | null> = {}
  if (pp.indentLeftTwips !== undefined) {
    ind['w:left'] = pp.indentLeftTwips === null ? null : num(pp.indentLeftTwips)
    ind['w:start'] = null
  }
  if (pp.indentRightTwips !== undefined) {
    ind['w:right'] = pp.indentRightTwips === null ? null : num(pp.indentRightTwips)
    ind['w:end'] = null
  }
  if (pp.firstLineTwips !== undefined) {
    const v = pp.firstLineTwips
    ind['w:firstLine'] = v !== null && v > 0 ? num(v) : null
    ind['w:hanging'] = v !== null && v < 0 ? num(-v) : null
  }
  if (Object.keys(ind).length > 0) children.attrs('w:ind', ind)
  children.flag('w:keepNext', pp.keepNext)
  children.flag('w:keepLines', pp.keepLines)
  children.flag('w:pageBreakBefore', pp.pageBreakBefore)
  if (pp.outlineLevel !== undefined) {
    children.set(
      'w:outlineLvl',
      pp.outlineLevel === null ? null : `<w:outlineLvl w:val="${num(pp.outlineLevel - 1)}"/>`,
    )
  }
}

/** The w:style element after applying `up` to `existing` (null = create). */
export function mergeStyleXml(existing: string | null, up: StyleUpsert): string {
  const creating = existing === null
  const attrs: Attrs = creating
    ? new Map([
        ['w:type', up.type ?? 'paragraph'],
        ['w:styleId', escapeXmlAttr(up.styleId)],
        ['w:customStyle', '1'],
      ])
    : parseTag(existing).attrs
  const children = new Children(
    creating ? [] : splitXmlChildren(innerOf(existing)),
    STYLE_CHILD_ORDER,
  )
  const name = up.name ?? (creating ? up.styleId : undefined)
  if (name !== undefined) children.set('w:name', `<w:name w:val="${escapeXmlAttr(name)}"/>`)
  if (up.basedOn !== undefined) {
    children.set(
      'w:basedOn',
      up.basedOn === null ? null : `<w:basedOn w:val="${escapeXmlAttr(up.basedOn)}"/>`,
    )
  }
  if (up.next !== undefined) {
    children.set('w:next', up.next === null ? null : `<w:next w:val="${escapeXmlAttr(up.next)}"/>`)
  }
  const quick = up.quickFormat ?? (creating ? true : undefined)
  if (quick !== undefined) children.set('w:qFormat', quick ? '<w:qFormat/>' : null)
  if (up.pPr) {
    const cur = children.get('w:pPr')
    const inner = new Children(cur ? splitXmlChildren(innerOf(cur.xml)) : [], PPR_CHILD_ORDER)
    patchPara(inner, up.pPr)
    children.set('w:pPr', inner.size ? `<w:pPr>${inner.toXml()}</w:pPr>` : null)
  }
  if (up.rPr) {
    const cur = children.get('w:rPr')
    const inner = new Children(cur ? splitXmlChildren(innerOf(cur.xml)) : [], RPR_CHILD_ORDER)
    patchRun(inner, up.rPr)
    children.set('w:rPr', inner.size ? `<w:rPr>${inner.toXml()}</w:rPr>` : null)
  }
  return tag('w:style', attrs, children.toXml())
}

/** Patch only the requested default font slots, preserving all other defaults. */
export type DefaultFonts = Pick<StyleRunProps, 'font' | 'eastAsiaFont'>
export function mergeDefaultFontsXml(xml: string, fonts: DefaultFonts): string {
  const declaration = xml.slice(0, xml.indexOf('<w:styles'))
  const root = xml.slice(xml.indexOf('<w:styles'))
  const styles = new Children(splitXmlChildren(innerOf(root)), [
    'w:docDefaults',
    'w:latentStyles',
    'w:style',
  ])
  const defaults = new Children(
    splitXmlChildren(innerOf(styles.get('w:docDefaults')?.xml ?? '<w:docDefaults/>')),
    ['w:rPrDefault', 'w:pPrDefault'],
  )
  const runDefault = new Children(
    splitXmlChildren(innerOf(defaults.get('w:rPrDefault')?.xml ?? '<w:rPrDefault/>')),
    ['w:rPr'],
  )
  const run = new Children(
    splitXmlChildren(innerOf(runDefault.get('w:rPr')?.xml ?? '<w:rPr/>')),
    RPR_CHILD_ORDER,
  )
  patchRun(run, fonts)
  runDefault.set('w:rPr', `<w:rPr>${run.toXml()}</w:rPr>`)
  defaults.set('w:rPrDefault', `<w:rPrDefault>${runDefault.toXml()}</w:rPrDefault>`)
  styles.set('w:docDefaults', `<w:docDefaults>${defaults.toXml()}</w:docDefaults>`)
  return declaration + tag('w:styles', parseTag(root).attrs, styles.toXml())
}
