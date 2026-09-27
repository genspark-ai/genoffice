import JSZip from 'jszip'
import type { NumberingLevel } from './types'

/**
 * Blank document template for "new document" and AI from-scratch generation.
 *
 * The generate layer only references styles / numbering that already exist in
 * the package, so a fresh document must ship with the standard set: Normal,
 * Heading1-6, ListParagraph, Hyperlink, TOC1-9 (for generated TOC fields),
 * plus bullet (numId 1) and decimal (numId 2) numbering definitions.
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

const DOC_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'

/** numId of the template's bullet list definition */
export const BLANK_BULLET_NUM_ID = '1'
/** numId of the template's ordered (decimal) list definition */
export const BLANK_ORDERED_NUM_ID = '2'

/** TOC entry styles referenced by generateTocFieldXml (indent grows per level) */
function tocStyle(level: number): string {
  const ind = level > 1 ? `<w:ind w:left="${220 * (level - 1)}"/>` : ''
  return (
    `<w:style w:type="paragraph" w:styleId="TOC${level}"><w:name w:val="toc ${level}"/>` +
    '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/>' +
    `<w:pPr><w:spacing w:after="100"/>${ind}</w:pPr></w:style>`
  )
}

function headingStyle(level: number, sizeHalfPoints: number): string {
  // Word lists Heading 3+ in the gallery only once the document uses them
  const hidden = level >= 3 ? '<w:semiHidden/><w:unhideWhenUsed/>' : ''
  return (
    `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/>` +
    `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/>${hidden}<w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="${level <= 2 ? 240 : 160}" w:after="120"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
    `<w:rPr><w:b/><w:sz w:val="${sizeHalfPoints}"/><w:szCs w:val="${sizeHalfPoints}"/></w:rPr></w:style>`
  )
}

/** accent1 (4472C4) shade BF / text1 tints of the Office theme, as Word caches them in w:val */
const ACCENT_DARK = 'w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"'
const TEXT_BF = 'w:val="404040" w:themeColor="text1" w:themeTint="BF"'
const TEXT_A6 = 'w:val="595959" w:themeColor="text1" w:themeTint="A6"'
const TEXT_A5 = 'w:val="5A5A5A" w:themeColor="text1" w:themeTint="A5"'

function paraStyle(id: string, name: string, priority: number, pPr: string, rPr = ''): string {
  return (
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/>` +
    `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="${priority}"/><w:qFormat/>` +
    `<w:pPr>${pPr}</w:pPr>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}</w:style>`
  )
}

function charStyle(id: string, name: string, priority: number, rPr: string): string {
  return (
    `<w:style w:type="character" w:styleId="${id}"><w:name w:val="${name}"/>` +
    `<w:uiPriority w:val="${priority}"/><w:qFormat/><w:rPr>${rPr}</w:rPr></w:style>`
  )
}

/** Word's Quick Style set (Normal and the headings are defined separately), in gallery order */
const QUICK_STYLES_XML =
  '<w:style w:type="paragraph" w:styleId="NoSpacing"><w:name w:val="No Spacing"/><w:uiPriority w:val="1"/><w:qFormat/>' +
  '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>' +
  paraStyle(
    'Title',
    'Title',
    10,
    '<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>',
    '<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>',
  ) +
  paraStyle(
    'Subtitle',
    'Subtitle',
    11,
    '<w:spacing w:after="160"/>',
    `<w:color ${TEXT_A6}/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/>`,
  ) +
  charStyle('SubtleEmphasis', 'Subtle Emphasis', 19, `<w:i/><w:iCs/><w:color ${TEXT_BF}/>`) +
  charStyle('Emphasis', 'Emphasis', 20, '<w:i/><w:iCs/>') +
  charStyle('IntenseEmphasis', 'Intense Emphasis', 21, `<w:i/><w:iCs/><w:color ${ACCENT_DARK}/>`) +
  charStyle('Strong', 'Strong', 22, '<w:b/><w:bCs/>') +
  paraStyle(
    'Quote',
    'Quote',
    29,
    '<w:spacing w:before="200" w:after="480"/><w:jc w:val="center"/>',
    `<w:i/><w:iCs/><w:color ${TEXT_BF}/>`,
  ) +
  paraStyle(
    'IntenseQuote',
    'Intense Quote',
    30,
    `<w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="2F5496" w:themeColor="accent1" w:themeShade="BF"/>` +
      `<w:bottom w:val="single" w:sz="4" w:space="10" w:color="2F5496" w:themeColor="accent1" w:themeShade="BF"/></w:pBdr>` +
      '<w:spacing w:before="360" w:after="440"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>',
    `<w:i/><w:iCs/><w:color ${ACCENT_DARK}/>`,
  ) +
  charStyle('SubtleReference', 'Subtle Reference', 31, `<w:smallCaps/><w:color ${TEXT_A5}/>`) +
  charStyle(
    'IntenseReference',
    'Intense Reference',
    32,
    `<w:b/><w:bCs/><w:smallCaps/><w:color ${ACCENT_DARK}/><w:spacing w:val="5"/>`,
  ) +
  charStyle('BookTitle', 'Book Title', 33, '<w:b/><w:bCs/><w:i/><w:iCs/><w:spacing w:val="5"/>') +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>' +
  '<w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'

const stylesXml = (eastAsiaFont?: string): string =>
  XML_DECL +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr>' +
  `<w:rFonts w:ascii="Calibri"${eastAsiaFont ? ` w:eastAsia="${eastAsiaFont}"` : ''} w:hAnsi="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/>` +
  '</w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  headingStyle(1, 32) +
  headingStyle(2, 28) +
  headingStyle(3, 26) +
  headingStyle(4, 24) +
  headingStyle(5, 22) +
  headingStyle(6, 22) +
  QUICK_STYLES_XML +
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>' +
  '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
  Array.from({ length: 9 }, (_, i) => tocStyle(i + 1)).join('') +
  '</w:styles>'

function bulletLevels(): string {
  let xml = ''
  for (let ilvl = 0; ilvl < 5; ilvl++) {
    xml +=
      `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/>` +
      `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/></w:pPr>` +
      '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>'
  }
  return xml
}

function decimalLevels(): string {
  let xml = ''
  for (let ilvl = 0; ilvl < 5; ilvl++) {
    xml +=
      `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%${ilvl + 1}."/>` +
      `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (ilvl + 1)}" w:hanging="360"/></w:pPr></w:lvl>`
  }
  return xml
}

/** Custom numbering level (one level of a multilevel list / bullet library / numbering library definition) */
export interface CustomNumberingLevel {
  /** w:numFmt:decimal/bullet/lowerLetter/upperRoman/chineseCountingThousand… */
  numFmt: string
  /** w:lvlText: a pattern like "%1." or a literal bullet symbol */
  lvlText: string
  /** w:ind w:left(twips) */
  indentLeft: number
  /** w:ind w:hanging (twips, default 360) */
  hanging?: number
  /** w:start (default 1) */
  start?: number
  /** w:suff (default tab) */
  suff?: 'tab' | 'space' | 'nothing'
  /** w:lvlJc (default left) */
  lvlJc?: 'left' | 'center' | 'right'
  /** w:pPr/w:tabs: an explicit tab stop after the marker (twips) */
  tabStop?: number
  /** w:rPr of the marker: font family, size (half-points), color (hex), weight */
  font?: string
  szHalfPoints?: number
  color?: string
  bold?: boolean
  italic?: boolean
  /** w:pStyle: paragraphs in this style take the level's numbering */
  pStyle?: string
  /** w:isLgl: legal-style numbering (every placeholder shows as decimal) */
  isLgl?: boolean
  /** w:lvlRestart: restart this level after the given (higher) level; 0 = never */
  lvlRestart?: number
  /** w:lvlPicBulletId: draw a w:numPicBullet image instead of lvlText */
  picBulletId?: number
}

/** A parsed level in the shape the define/adjust dialogs edit */
export function customLevelFromNumberingLevel(level: NumberingLevel): CustomNumberingLevel {
  const out: CustomNumberingLevel = {
    numFmt: level.numFmt,
    lvlText: level.lvlText,
    indentLeft: level.indentLeft ?? 0,
    hanging: level.hanging ?? (level.indentLeft === undefined ? 360 : 0),
    start: level.start,
  }
  if (level.suff) out.suff = level.suff
  if (level.lvlJc) out.lvlJc = level.lvlJc
  if (level.font) out.font = level.font
  if (level.szHalfPoints) out.szHalfPoints = level.szHalfPoints
  if (level.color) out.color = level.color
  if (level.isLgl) out.isLgl = true
  if (level.picBulletId !== undefined) out.picBulletId = level.picBulletId
  if (level.bold) out.bold = true
  if (level.italic) out.italic = true
  if (level.tabStop !== undefined) out.tabStop = level.tabStop
  if (level.pStyle) out.pStyle = level.pStyle
  if (level.lvlRestart !== undefined) out.lvlRestart = level.lvlRestart
  return out
}

/** The renderer's overlay shape for a level the dialogs produced (display before save) */
export function numberingLevelFromCustom(
  l: CustomNumberingLevel,
  picBulletSrc?: string,
): NumberingLevel {
  const out: NumberingLevel = {
    numFmt: l.numFmt,
    lvlText: l.lvlText,
    start: l.start ?? 1,
    indentLeft: l.indentLeft,
    hanging: l.hanging ?? 360,
  }
  if (l.suff) out.suff = l.suff
  if (l.lvlJc) out.lvlJc = l.lvlJc
  if (l.font) out.font = l.font
  if (l.szHalfPoints) out.szHalfPoints = l.szHalfPoints
  if (l.color) out.color = l.color
  if (l.isLgl) out.isLgl = true
  if (l.picBulletId !== undefined) out.picBulletId = l.picBulletId
  if (picBulletSrc) out.picBulletSrc = picBulletSrc
  if (l.bold) out.bold = true
  if (l.italic) out.italic = true
  if (l.tabStop !== undefined) out.tabStop = l.tabStop
  if (l.pStyle) out.pStyle = l.pStyle
  if (l.lvlRestart !== undefined) out.lvlRestart = l.lvlRestart
  return out
}

const LVL_CHILD_ORDER = [
  'w:start',
  'w:numFmt',
  'w:lvlRestart',
  'w:pStyle',
  'w:isLgl',
  'w:suff',
  'w:lvlText',
  'w:lvlPicBulletId',
  'w:legacy',
  'w:lvlJc',
  'w:pPr',
  'w:rPr',
]

/** top-level children of a small element as [name, xml] pairs (no same-name nesting inside a w:lvl) */
function splitChildren(inner: string): Array<[string, string]> {
  const out: Array<[string, string]> = []
  const re = /<(w:[\w]+)\b(?:[^>]*?\/>|[^>]*>[\s\S]*?<\/\1>)/g
  for (const m of inner.matchAll(re)) out.push([m[1], m[0]])
  return out
}

function innerOf(xml: string): string {
  const open = xml.indexOf('>')
  const close = xml.lastIndexOf('</')
  return xml.endsWith('/>') || close < 0 ? '' : xml.slice(open + 1, close)
}

/**
 * Merge a dialog-produced level into an existing w:lvl: the fields the level
 * carries replace their elements, everything else in the original (legacy,
 * w14 extensions, rPr children the dialogs do not know) keeps its bytes.
 */
export function mergeLevelXml(existing: string, l: CustomNumberingLevel, ilvl: number): string {
  const children = new Map(splitChildren(innerOf(existing)))
  const set = (name: string, xml: string | null) => {
    if (xml === null) children.delete(name)
    else children.set(name, xml)
  }
  set('w:start', `<w:start w:val="${l.start ?? 1}"/>`)
  set('w:numFmt', `<w:numFmt w:val="${escAttr(l.numFmt)}"/>`)
  if (l.lvlRestart !== undefined) set('w:lvlRestart', `<w:lvlRestart w:val="${l.lvlRestart}"/>`)
  if (l.pStyle !== undefined)
    set('w:pStyle', l.pStyle ? `<w:pStyle w:val="${escAttr(l.pStyle)}"/>` : null)
  if (l.isLgl !== undefined) set('w:isLgl', l.isLgl ? '<w:isLgl/>' : null)
  if (l.suff !== undefined) set('w:suff', l.suff === 'tab' ? null : `<w:suff w:val="${l.suff}"/>`)
  set('w:lvlText', `<w:lvlText w:val="${escAttr(l.lvlText)}"/>`)
  if (l.picBulletId !== undefined)
    set('w:lvlPicBulletId', `<w:lvlPicBulletId w:val="${l.picBulletId}"/>`)
  set('w:lvlJc', `<w:lvlJc w:val="${l.lvlJc ?? 'left'}"/>`)

  const pPr = new Map(splitChildren(innerOf(children.get('w:pPr') ?? '')))
  const hanging = Math.round(l.hanging ?? 360)
  pPr.set('w:ind', `<w:ind w:left="${Math.round(l.indentLeft)}" w:hanging="${hanging}"/>`)
  if (l.tabStop !== undefined)
    pPr.set('w:tabs', `<w:tabs><w:tab w:val="num" w:pos="${Math.round(l.tabStop)}"/></w:tabs>`)
  const pPrOrder = ['w:tabs', 'w:ind']
  const pPrXml = [...pPr.entries()]
    .sort((a, b) => pPrOrder.indexOf(a[0]) - pPrOrder.indexOf(b[0]))
    .map(([, xml]) => xml)
    .join('')
  set('w:pPr', `<w:pPr>${pPrXml}</w:pPr>`)

  const rPr = new Map(splitChildren(innerOf(children.get('w:rPr') ?? '')))
  if (l.font !== undefined) {
    const f = escAttr(l.font)
    if (l.font)
      rPr.set(
        'w:rFonts',
        `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}" w:hint="default"/>`,
      )
    else rPr.delete('w:rFonts')
  }
  const toggle = (name: string, on: boolean | undefined, xml: string) => {
    if (on === undefined) return
    if (on) rPr.set(name, xml)
    else rPr.delete(name)
  }
  toggle('w:b', l.bold, '<w:b/>')
  toggle('w:i', l.italic, '<w:i/>')
  if (l.color !== undefined) {
    if (l.color) rPr.set('w:color', `<w:color w:val="${escAttr(l.color)}"/>`)
    else rPr.delete('w:color')
  }
  if (l.szHalfPoints !== undefined) {
    if (l.szHalfPoints) {
      rPr.set('w:sz', `<w:sz w:val="${l.szHalfPoints}"/>`)
      rPr.set('w:szCs', `<w:szCs w:val="${l.szHalfPoints}"/>`)
    } else {
      rPr.delete('w:sz')
      rPr.delete('w:szCs')
    }
  }
  const rPrOrder = ['w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:color', 'w:sz', 'w:szCs']
  const rPrXml = [...rPr.entries()]
    .sort((a, b) => {
      const ia = rPrOrder.indexOf(a[0])
      const ib = rPrOrder.indexOf(b[0])
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib)
    })
    .map(([, xml]) => xml)
    .join('')
  set('w:rPr', rPrXml ? `<w:rPr>${rPrXml}</w:rPr>` : null)

  const body = [...children.entries()]
    .sort((a, b) => {
      const ia = LVL_CHILD_ORDER.indexOf(a[0])
      const ib = LVL_CHILD_ORDER.indexOf(b[0])
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib)
    })
    .map(([, xml]) => xml)
    .join('')
  const openTag = /<w:lvl\b[^>]*>/.exec(existing)?.[0] ?? `<w:lvl w:ilvl="${ilvl}">`
  return `${openTag.replace(/\/>$/, '>')}${body}</w:lvl>`
}

const escAttr = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/** One w:lvl in CT_Lvl child order */
export function customLevelXml(l: CustomNumberingLevel, ilvl: number): string {
  const parts = [`<w:start w:val="${l.start ?? 1}"/>`, `<w:numFmt w:val="${escAttr(l.numFmt)}"/>`]
  if (l.lvlRestart !== undefined) parts.push(`<w:lvlRestart w:val="${l.lvlRestart}"/>`)
  if (l.pStyle) parts.push(`<w:pStyle w:val="${escAttr(l.pStyle)}"/>`)
  if (l.isLgl) parts.push('<w:isLgl/>')
  if (l.suff && l.suff !== 'tab') parts.push(`<w:suff w:val="${l.suff}"/>`)
  parts.push(`<w:lvlText w:val="${escAttr(l.lvlText)}"/>`)
  if (l.picBulletId !== undefined) parts.push(`<w:lvlPicBulletId w:val="${l.picBulletId}"/>`)
  parts.push(`<w:lvlJc w:val="${l.lvlJc ?? 'left'}"/>`)
  const tabs =
    l.tabStop !== undefined
      ? `<w:tabs><w:tab w:val="num" w:pos="${Math.round(l.tabStop)}"/></w:tabs>`
      : ''
  const hanging = Math.round(l.hanging ?? 360)
  parts.push(
    `<w:pPr>${tabs}<w:ind w:left="${Math.round(l.indentLeft)}" w:hanging="${hanging}"/></w:pPr>`,
  )
  const rPr: string[] = []
  if (l.font) {
    const f = escAttr(l.font)
    rPr.push(
      `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}" w:hint="default"/>`,
    )
  }
  if (l.bold) rPr.push('<w:b/>')
  if (l.italic) rPr.push('<w:i/>')
  if (l.color) rPr.push(`<w:color w:val="${escAttr(l.color)}"/>`)
  if (l.szHalfPoints)
    rPr.push(`<w:sz w:val="${l.szHalfPoints}"/><w:szCs w:val="${l.szHalfPoints}"/>`)
  if (rPr.length) parts.push(`<w:rPr>${rPr.join('')}</w:rPr>`)
  return `<w:lvl w:ilvl="${ilvl}">${parts.join('')}</w:lvl>`
}

function customLevels(levels: CustomNumberingLevel[]): string {
  return levels.map((l, ilvl) => customLevelXml(l, ilvl)).join('')
}

/** w:numPicBullet with a DrawingML picture; the image relationship lives in numbering.xml.rels */
export function numPicBulletXml(id: number, rId: string): string {
  return (
    `<w:numPicBullet w:numPicBulletId="${id}"><w:drawing>` +
    `<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="152400" cy="152400"/><wp:docPr id="${id}" name="Picture bullet ${id}"/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="Picture bullet ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="152400" cy="152400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:numPicBullet>`
  )
}

/** One abstractNum definition; uses custom levels when provided, otherwise the blank-template style (5 levels) */
export function abstractNumXml(
  abstractNumId: string,
  kind: 'bullet' | 'ordered',
  levels?: CustomNumberingLevel[],
): string {
  const body = levels?.length
    ? customLevels(levels)
    : kind === 'bullet'
      ? bulletLevels()
      : decimalLevels()
  return `<w:abstractNum w:abstractNumId="${abstractNumId}">${body}</w:abstractNum>`
}

/** The blank template's numbering.xml (bullet numId 1 / decimal numId 2);
 *  base content used when the document has no such part */
export const BLANK_NUMBERING_XML =
  XML_DECL +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  `<w:abstractNum w:abstractNumId="0">${bulletLevels()}</w:abstractNum>` +
  `<w:abstractNum w:abstractNumId="1">${decimalLevels()}</w:abstractNum>` +
  `<w:num w:numId="${BLANK_BULLET_NUM_ID}"><w:abstractNumId w:val="0"/></w:num>` +
  `<w:num w:numId="${BLANK_ORDERED_NUM_ID}"><w:abstractNumId w:val="1"/></w:num>` +
  '</w:numbering>'

const NUMBERING_XML = BLANK_NUMBERING_XML

export type BlankPaperSize = 'A4' | 'Letter'

/** Word starts new documents on Letter in these OS regions and on A4 everywhere else. */
const LETTER_REGIONS = new Set([
  'US',
  'CA',
  'MX',
  'PH',
  'LR',
  'MM',
  'PR',
  'CL',
  'CO',
  'VE',
  'GT',
  'CR',
  'PA',
  'DO',
])

export function paperSizeForRegion(region: string | null | undefined): BlankPaperSize {
  return region && LETTER_REGIONS.has(region.toUpperCase()) ? 'Letter' : 'A4'
}

/** Paper size for a BCP 47 system locale such as `en-US` or `zh-Hans-CN`. */
export function paperSizeForLocale(locale: string | null | undefined): BlankPaperSize {
  if (!locale) return 'A4'
  let region: string | undefined
  try {
    region = new Intl.Locale(locale.replace(/_/g, '-')).maximize().region
  } catch {
    region = undefined
  }
  return paperSizeForRegion(region)
}

const PAPER_TWIPS: Record<BlankPaperSize, { w: number; h: number }> = {
  A4: { w: 11906, h: 16838 },
  Letter: { w: 12240, h: 15840 },
}

export interface BlankDocxOptions {
  /**
   * docDefaults East Asian font (w:eastAsia); pick per UI language so e.g.
   * Japanese/Korean users don't start with a Simplified-Chinese face.
   * When omitted no w:eastAsia is written — like an en-US Word document,
   * Word/our renderer then substitute per script when CJK text appears.
   */
  eastAsiaFont?: string
  /** Page size of the single section; defaults to A4. */
  paperSize?: BlankPaperSize
}

/** Build a minimal valid .docx: one empty paragraph, portrait page, standard styles. */
export async function buildBlankDocx(options?: BlankDocxOptions): Promise<Uint8Array> {
  const zip = new JSZip()

  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '</Types>',
  )

  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )

  zip.file(
    'word/_rels/document.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
      '</Relationships>',
  )

  zip.file('word/styles.xml', stylesXml(options?.eastAsiaFont))
  zip.file('word/numbering.xml', NUMBERING_XML)

  const paper = PAPER_TWIPS[options?.paperSize ?? 'A4']
  const sectPr =
    `<w:sectPr><w:pgSz w:w="${paper.w}" w:h="${paper.h}"/>` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr>'
  zip.file(
    'word/document.xml',
    `${XML_DECL}<w:document ${DOC_NS}><w:body><w:p/>${sectPr}</w:body></w:document>`,
  )

  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
