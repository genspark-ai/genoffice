import type {
  PageNumberChapter,
  ParsedDoc,
  SectionInfo,
  SectionSettings,
  DocGrid,
  NoteProps,
  LineNumbering,
} from './types'
import { onOffTagIn } from './parse-xml-text'
import { escapeXmlAttr } from './xml-utils'

/** US Letter, portrait, 1-inch margins */
export const DEFAULT_SECTION: SectionSettings = {
  pageWidth: 12240,
  pageHeight: 15840,
  orientation: 'portrait',
  marginTop: 1440,
  marginRight: 1440,
  marginBottom: 1440,
  marginLeft: 1440,
  pageBorder: false,
  columns: 1,
  headerDist: 720,
  footerDist: 720,
}

/** Vertical alignment of page content (sectPr w:vAlign); top/default returns undefined */
function vAlignOf(xml: string): 'center' | 'both' | 'bottom' | undefined {
  const v = strAttr(/<w:vAlign[^>]*\/?>/.exec(xml)?.[0] ?? '', 'w:val')
  return v === 'center' || v === 'both' || v === 'bottom' ? v : undefined
}

/** String attribute of a tag; matches either quote style. */
function strAttr(tag: string, name: string): string | undefined {
  return new RegExp(`${name}\\s*=\\s*["']([^"']*)["']`).exec(tag)?.[1]
}

/**
 * Integer attribute of a tag; XML allows either quote style, so match both.
 * Fractional twips (converter output like w:left="1987.1997") round like the
 * paragraph spacing/indent readers do instead of dropping to the default.
 */
function intAttr(tag: string, name: string, fallback: number): number {
  const m = new RegExp(`${name}=["'](-?\\d+(?:\\.\\d+)?)["']`).exec(tag)
  const v = m ? Math.round(parseFloat(m[1])) : NaN
  return Number.isFinite(v) ? v : fallback
}

/**
 * True when w:pgBorders declares at least one visible side. Word writes explicit
 * "no border" as side elements with w:val="none" (or "nil"), so the mere presence
 * of w:pgBorders must not be treated as a drawn border.
 */
function hasVisiblePageBorder(xml: string): boolean {
  const pgBorders = /<w:pgBorders[^>]*\/>|<w:pgBorders[\s\S]*?<\/w:pgBorders>/.exec(xml)?.[0]
  if (!pgBorders) return false
  const sides = pgBorders.match(/<w:(?:top|left|bottom|right)\b[^>]*\/?>/g) ?? []
  return sides.some((side) => {
    const val = strAttr(side, 'w:val')
    return val !== undefined && val !== 'none' && val !== 'nil'
  })
}

function onOffAttr(tag: string, attr: string): boolean {
  const v = strAttr(tag, attr)
  return v === '1' || v === 'true' || v === 'on'
}

function colsUnequal(colsTag: string): boolean {
  return /^(?:0|false|off)$/.test(strAttr(colsTag, 'w:equalWidth') ?? '')
}

/** ST_Border line styles; every other w:val is an art (bitmap) border. */
export const LINE_BORDER_VALS = new Set([
  'single',
  'thick',
  'double',
  'dotted',
  'dashed',
  'dotDash',
  'dotDotDash',
  'triple',
  'thinThickSmallGap',
  'thickThinSmallGap',
  'thinThickThinSmallGap',
  'thinThickMediumGap',
  'thickThinMediumGap',
  'thinThickThinMediumGap',
  'thinThickLargeGap',
  'thickThinLargeGap',
  'thinThickThinLargeGap',
  'wave',
  'doubleWave',
  'dashSmallGap',
  'dashDotStroked',
  'threeDEmboss',
  'threeDEngrave',
  'outset',
  'inset',
])

/** Styling details of a visible w:pgBorders box (undefined when no visible side). */
function pageBorderPropsOf(xml: string): SectionSettings['pageBorderProps'] {
  const pgBorders = /<w:pgBorders[^>]*\/>|<w:pgBorders[\s\S]*?<\/w:pgBorders>/.exec(xml)?.[0]
  if (!pgBorders || !hasVisiblePageBorder(xml)) return undefined
  const display = strAttr(pgBorders, 'w:display') as 'firstPage' | 'notFirstPage' | undefined
  const offsetFrom = strAttr(pgBorders, 'w:offsetFrom') as 'page' | 'text' | undefined
  const zOrder = /\bw:zOrder\s*=\s*["']back["']/.test(pgBorders) ? 'back' : undefined
  let spacePt = 0
  let widthPt = 0
  let color: string | undefined
  const sides: NonNullable<SectionSettings['pageBorderProps']>['sides'] = {}
  for (const side of pgBorders.match(/<w:(?:top|left|bottom|right)\b[^>]*\/?>/g) ?? []) {
    const val = strAttr(side, 'w:val')
    if (!val || val === 'none' || val === 'nil') continue
    const name = /<w:(top|left|bottom|right)\b/.exec(side)![1] as keyof typeof sides
    const sideColor = /w:color\s*=\s*["']([0-9A-Fa-f]{6})["']/.exec(side)?.[1]
    // art borders: w:sz is the tiled pattern height in points, not eighth-points
    const art = !LINE_BORDER_VALS.has(val)
    const sz = intAttr(side, 'w:sz', 0)
    sides[name] = {
      val,
      widthPt: art ? sz : sz / 8,
      spacePt: intAttr(side, 'w:space', 0),
      ...(sideColor ? { color: sideColor } : {}),
      ...(art ? { art } : {}),
      ...(onOffAttr(side, 'w:shadow') ? { shadow: true } : {}),
      ...(onOffAttr(side, 'w:frame') ? { frame: true } : {}),
    }
    spacePt = Math.max(spacePt, intAttr(side, 'w:space', 0))
    widthPt = Math.max(widthPt, sides[name]!.widthPt)
    color ??= sideColor
  }
  return {
    ...(display ? { display } : {}),
    ...(offsetFrom ? { offsetFrom } : {}),
    ...(zOrder ? { zOrder } : {}),
    spacePt,
    widthPt,
    ...(color ? { color } : {}),
    sides,
  }
}

/** Page setup from one w:sectPr XML slice. */
export function sectionSettingsFromXml(
  xml: string,
  opts: { gutterAtTop?: boolean } = {},
): SectionSettings {
  const pgSz = /<w:pgSz[^>]*\/?>/.exec(xml)?.[0] ?? ''
  const pgMar = /<w:pgMar[^>]*\/?>/.exec(xml)?.[0] ?? ''

  // Parse docGrid (w:docGrid)
  let docGrid: DocGrid | undefined
  const docGridTag = /<w:docGrid[^>]*\/?>/.exec(xml)?.[0]
  if (docGridTag) {
    const gridType = (strAttr(docGridTag, 'w:type') ?? 'default') as DocGrid['type']
    const validTypes: DocGrid['type'][] = ['default', 'lines', 'linesAndChars', 'snapToChars']
    const linePitch = intAttr(docGridTag, 'w:linePitch', -1)
    const charSpace = intAttr(docGridTag, 'w:charSpace', Number.NaN)
    docGrid = {
      type: validTypes.includes(gridType) ? gridType : 'default',
      ...(linePitch >= 0 ? { linePitch } : {}),
      ...(Number.isFinite(charSpace) ? { charSpace } : {}),
    }
  }

  // explicit unequal column widths (w:cols > w:col children). w:col is an empty
  // element, so matching its start tag covers both the self-closing and the paired
  // spelling, and either attribute quote style
  const colsElement = /<w:cols[^>]*>[\s\S]*?<\/w:cols>/.exec(xml)?.[0]
  const colTags = (colsElement?.match(/<w:col [^>]*w:w=["'][\d.]+["'][^>]*>/g) ?? []).filter(
    (tag) => intAttr(tag, 'w:w', 0) > 0,
  )
  const colWidths = colTags.map((tag) => intAttr(tag, 'w:w', 0))
  const colSpaces = colTags.slice(0, -1).map((tag) => intAttr(tag, 'w:space', 0))

  const colsTag = /<w:cols\b[^>]*>/.exec(xml)?.[0] ?? ''
  const pageBorderProps = pageBorderPropsOf(xml)
  const footnotePr = notePropsFromXml(liveSectPr(xml), 'w:footnotePr')
  const endnotePr = notePropsFromXml(liveSectPr(xml), 'w:endnotePr')
  const lineNumbers = lineNumberingOf(xml)
  const marginTop = intAttr(pgMar, 'w:top', DEFAULT_SECTION.marginTop)
  const marginBottom = intAttr(pgMar, 'w:bottom', DEFAULT_SECTION.marginBottom)
  // Word widens the binding-side margin by the gutter (left, or top with gutterAtTop;
  // mirrored inside margins are not modelled)
  const gutter = Math.max(0, intAttr(pgMar, 'w:gutter', 0))
  const gutterAtTop = gutter > 0 && !!opts.gutterAtTop

  return {
    pageWidth: intAttr(pgSz, 'w:w', DEFAULT_SECTION.pageWidth),
    pageHeight: intAttr(pgSz, 'w:h', DEFAULT_SECTION.pageHeight),
    orientation: strAttr(pgSz, 'w:orient') === 'landscape' ? 'landscape' : 'portrait',
    marginTop: Math.abs(marginTop) + (gutterAtTop ? gutter : 0),
    marginRight: intAttr(pgMar, 'w:right', DEFAULT_SECTION.marginRight),
    marginBottom: Math.abs(marginBottom),
    marginLeft: intAttr(pgMar, 'w:left', DEFAULT_SECTION.marginLeft) + (gutterAtTop ? 0 : gutter),
    ...(marginTop < 0 ? { marginTopFixed: true } : {}),
    ...(marginBottom < 0 ? { marginBottomFixed: true } : {}),
    ...(gutter > 0 ? { gutter } : {}),
    ...(gutterAtTop ? { gutterAtTop: true } : {}),
    headerDist: intAttr(pgMar, 'w:header', 720),
    footerDist: intAttr(pgMar, 'w:footer', 720),
    ...(vAlignOf(xml) ? { vAlign: vAlignOf(xml) } : {}),
    pageBorder: hasVisiblePageBorder(xml),
    ...(pageBorderProps ? { pageBorderProps } : {}),
    columns: intAttr(/<w:cols[^>]*\/?>/.exec(xml)?.[0] ?? '', 'w:num', 1),
    colSpace: intAttr(/<w:cols[^>]*\/?>/.exec(xml)?.[0] ?? '', 'w:space', 720),
    ...(colWidths.length >= 2 ? { colWidths, colSpaces } : {}),
    ...(colsUnequal(colsTag) ? { colEqualWidth: false } : {}),
    ...(onOffAttr(colsTag, 'w:sep') ? { colSep: true } : {}),
    ...(lineNumbers ? { lineNumbers } : {}),
    ...(onOffTagIn(xml, 'w:bidi') ? { bidi: true } : {}),
    ...(docGrid ? { docGrid } : {}),
    ...(textDirectionOf(xml) ? { textDirection: textDirectionOf(xml) } : {}),
    ...(onOffTagIn(xml, 'w:noEndnote') ? { noEndnote: true } : {}),
    ...(footnotePr ? { footnotePr } : {}),
    ...(endnotePr ? { endnotePr } : {}),
  }
}

/** a sectPr without the previous one a tracked w:sectPrChange nests in it */
function liveSectPr(xml: string): string {
  const at = xml.indexOf('<w:sectPrChange')
  return at >= 0 ? xml.slice(0, at) : xml
}

/**
 * Fields declared by a w:footnotePr / w:endnotePr element (sectPr or settings.xml);
 * undefined when the element is absent or declares none of them.
 */
export function notePropsFromXml(
  xml: string,
  tag: 'w:footnotePr' | 'w:endnotePr',
): NoteProps | undefined {
  const el = new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`).exec(xml)?.[1]
  if (!el) return undefined
  const val = (name: string) => {
    const tag = new RegExp(`<w:${name}\\b[^>]*\\/?>`).exec(el)?.[0] ?? ''
    return strAttr(tag, 'w:val')
  }
  const pos = val('pos')
  const numFmt = val('numFmt')
  const numStart = val('numStart')
  const numRestart = val('numRestart')
  const out: NoteProps = {
    ...(pos === 'pageBottom' || pos === 'beneathText' || pos === 'sectEnd' || pos === 'docEnd'
      ? { pos }
      : {}),
    ...(numFmt ? { numFmt } : {}),
    ...(numStart && /^\d+$/.test(numStart) ? { numStart: parseInt(numStart, 10) } : {}),
    ...(numRestart === 'continuous' || numRestart === 'eachSect' || numRestart === 'eachPage'
      ? { numRestart }
      : {}),
  }
  return Object.keys(out).length > 0 ? out : undefined
}

const NOTE_PR_FIELDS = ['pos', 'numFmt', 'numStart', 'numRestart'] as const

/**
 * Set the fields `props` carries on a w:footnotePr / w:endnotePr element's inner
 * XML, in schema order; its other children (the separator references of the
 * settings part) stay as they are.
 */
function notePrInner(inner: string, props: NoteProps): string {
  const kept = new Map<string, string>()
  let rest = inner
  for (const name of NOTE_PR_FIELDS) {
    const re = new RegExp(`<w:${name}\\b[^>]*?(?:/>|>\\s*</w:${name}>)`)
    const hit = re.exec(rest)?.[0]
    if (hit) {
      kept.set(name, hit)
      rest = rest.replace(re, '')
    }
  }
  const lead = NOTE_PR_FIELDS.map((name) =>
    props[name] === undefined
      ? (kept.get(name) ?? '')
      : `<w:${name} w:val="${escapeXmlAttr(String(props[name]))}"/>`,
  ).join('')
  return lead + rest
}

type NotePrTag = 'w:footnotePr' | 'w:endnotePr'
const notePrRe = (tag: NotePrTag) => new RegExp(`<${tag}\\s*/>|<${tag}>([\\s\\S]*?)</${tag}>`)

/** settings.xml with the document-wide note properties set; `insert` places a new element */
export function applySettingsNoteProps(
  xml: string,
  tag: NotePrTag,
  props: NoteProps,
  insert: (xml: string, childXml: string) => string,
): string {
  const hit = notePrRe(tag).exec(xml)
  const element = `<${tag}>${notePrInner(hit?.[1] ?? '', props)}</${tag}>`
  return hit ? xml.replace(hit[0], element) : insert(xml, element)
}

/**
 * Every live section of the body takes the note properties: a sectPr's own
 * w:footnotePr / w:endnotePr overrides the document-wide one, so a setting made
 * for the whole document has to reach each of them. The previous sectPr inside a
 * tracked w:sectPrChange is left alone.
 */
export function applySectionNoteProps(
  documentXml: string,
  props: { footnote?: NoteProps; endnote?: NoteProps },
): string {
  return documentXml.replace(
    /<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g,
    (sectPr) => {
      // a w:sectPrChange nests the previous sectPr: the match then ends at that one's close tag
      const open = /^<w:sectPr\b[^>]*?(\/?)>/.exec(sectPr)!
      let body = open[1] ? '' : sectPr.slice(open[0].length, -'</w:sectPr>'.length)
      const changeAt = body.indexOf('<w:sectPrChange')
      const tail = changeAt >= 0 ? body.slice(changeAt) : ''
      if (changeAt >= 0) body = body.slice(0, changeAt)
      let endnotePr = ''
      let footnotePr = ''
      for (const [tag, own] of [
        ['w:endnotePr', props.endnote],
        ['w:footnotePr', props.footnote],
      ] as const) {
        const hit = notePrRe(tag).exec(body)
        if (hit) body = body.replace(hit[0], '')
        const element =
          own && Object.keys(own).length > 0
            ? `<${tag}>${notePrInner(hit?.[1] ?? '', own)}</${tag}>`
            : (hit?.[0] ?? '')
        if (tag === 'w:endnotePr') endnotePr = element
        else footnotePr = element
      }
      // CT_SectPr: header / footer references, then footnotePr, endnotePr
      const refs =
        /^(?:\s*<w:(?:headerReference|footerReference)\b[^>]*?(?:\/>|>\s*<\/w:(?:headerReference|footerReference)>))*/.exec(
          body,
        )![0]
      const openTag = open[1] ? open[0].slice(0, -2) + '>' : open[0]
      return `${openTag}${refs}${footnotePr}${endnotePr}${body.slice(refs.length)}${tail}</w:sectPr>`
    },
  )
}

/** w:lnNumType of a sectPr slice; undefined when the section has no line numbering */
export function lineNumberingOf(xml: string): LineNumbering | undefined {
  const tag = /<w:lnNumType\b[^>]*\/?>/.exec(xml)?.[0]
  if (!tag) return undefined
  const restart = strAttr(tag, 'w:restart')
  const distance = intAttr(tag, 'w:distance', -1)
  return {
    countBy: Math.max(1, intAttr(tag, 'w:countBy', 1)),
    // Word skips w:start lines before the first label and restarts per page unless told otherwise
    start: Math.max(0, intAttr(tag, 'w:start', 0)) + 1,
    ...(distance >= 0 ? { distance } : {}),
    restart: restart === 'continuous' || restart === 'newSection' ? restart : 'newPage',
  }
}

function textDirectionOf(xml: string): string | undefined {
  const val = strAttr(/<w:textDirection[^>]*\/?>/.exec(xml)?.[0] ?? '', 'w:val')
  return val && val !== 'lrTb' ? val : undefined
}

/** Read page setup from the trailing (hidden) w:sectPr. */
export function readSectionSettings(parsed: ParsedDoc): SectionSettings {
  const sectBlock = parsed.blocks.find((b) => b.hidden && b.originalXml?.includes('<w:sectPr'))
  return sectionSettingsFromXml(sectBlock?.originalXml ?? '', { gutterAtTop: parsed.gutterAtTop })
}

const SECT_PR_RE = /<w:sectPr[^>]*\/>|<w:sectPr[\s\S]*?<\/w:sectPr>/

/** On/off flag element (w:titlePg, w:evenAndOddHeaders, …) present and not switched
 *  off: matches self-closing and paired forms, w:val="0|false|off" counts as off. */
export function xmlFlagOn(xml: string, tag: string): boolean {
  for (const m of xml.matchAll(new RegExp(`<${tag}(?=[\\s/>])[^>]*>`, 'g'))) {
    const val = strAttr(m[0], 'w:val')
    if (val === undefined || !/^(?:0|false|off)$/.test(val)) return true
  }
  return false
}

/**
 * Header/footer reference elements of a sectPr, in document order. Matching is
 * quote-agnostic and covers the empty element pair LibreOffice writes, so a
 * reference a producer spelled differently still resolves: matching only
 * self-closing double-quoted tags dropped the part from the model, the save
 * plan and the section info at once.
 */
export function hfReferenceTags(xml: string, kind: 'header' | 'footer'): string[] {
  return (
    xml.match(
      new RegExp(`<w:${kind}Reference\\b[^>]*(?:\\/>|>\\s*<\\/w:${kind}Reference>)`, 'g'),
    ) ?? []
  )
}

/** w:type of a header/footer reference element (undefined when undeclared) */
export function hfReferenceType(tag: string): string | undefined {
  return /\bw:type\s*=\s*["']([^"']*)["']/.exec(tag)?.[1]
}

/** r:id of a header/footer reference element (undefined when undeclared) */
export function hfReferenceRId(tag: string): string | undefined {
  return /\br:id\s*=\s*["']([^"']+)["']/.exec(tag)?.[1]
}

function hfRefs(
  xml: string,
  kind: 'header' | 'footer',
): Partial<Record<'default' | 'first' | 'even', string>> {
  const refs: Partial<Record<'default' | 'first' | 'even', string>> = {}
  for (const ref of hfReferenceTags(xml, kind)) {
    const declared = hfReferenceType(ref)
    const type = declared === 'first' || declared === 'even' ? declared : 'default'
    const rId = hfReferenceRId(ref)
    if (rId) refs[type] = rId
  }
  return refs
}

export function sectionFromSectPr(
  sectPrXml: string,
  firstBlockIndex: number,
  lastBlockIndex: number,
  gutterAtTop?: boolean,
): SectionInfo {
  const pgNumTag = /<w:pgNumType[^>]*\/?>/.exec(sectPrXml)?.[0] ?? ''
  const type = strAttr(/<w:type[^>]*\/?>/.exec(sectPrXml)?.[0] ?? '', 'w:val')
  const startType =
    type === 'continuous' ||
    type === 'evenPage' ||
    type === 'oddPage' ||
    type === 'nextColumn' ||
    type === 'nextPage'
      ? type
      : 'nextPage'
  const pgNumStart = /w:start\s*=\s*["'](\d+)["']/.exec(pgNumTag)?.[1]
  const pgNumFmt = /w:fmt\s*=\s*["']([^"']+)["']/.exec(pgNumTag)?.[1]
  const chapStyle = /w:chapStyle\s*=\s*["'](\d+)["']/.exec(pgNumTag)?.[1]
  const chapSep = /w:chapSep\s*=\s*["'](\w+)["']/.exec(pgNumTag)?.[1]
  return {
    settings: sectionSettingsFromXml(sectPrXml, { gutterAtTop }),
    startType,
    firstBlockIndex,
    lastBlockIndex,
    sectPrXml,
    titlePg: xmlFlagOn(sectPrXml, 'w:titlePg'),
    ...(pgNumStart !== undefined ? { pageNumberStart: parseInt(pgNumStart, 10) } : {}),
    ...(pgNumFmt !== undefined ? { pageNumberFmt: pgNumFmt } : {}),
    ...(chapStyle !== undefined
      ? { pageNumberChapter: { style: parseInt(chapStyle, 10), sep: chapSep ?? 'hyphen' } }
      : {}),
    headerRefs: hfRefs(sectPrXml, 'header'),
    footerRefs: hfRefs(sectPrXml, 'footer'),
  }
}

/**
 * Remove every occurrence of an element written either self-closing
 * (<w:pgNumType .../>) or as an empty element pair (<w:pgNumType ...></w:pgNumType>).
 * Both spellings are valid OOXML, and stripping only the self-closing one left the
 * paired copy behind, so the tag got written twice into a single sectPr.
 */
function stripElement(xml: string, tag: string): string {
  return xml.replace(new RegExp(`<${tag}[^>]*\\/>|<${tag}[^>]*>[\\s\\S]*?<\\/${tag}>`, 'g'), '')
}

/**
 * Rewrite the sectPr page numbering w:pgNumType (fmt = number format, start = starting
 * page number; undefined fields are omitted, and the tag is removed when nothing is left).
 * `chapter`: the heading level whose number leads the page number and what separates
 * them; null takes it off, undefined leaves the file's own.
 * Schema order: pgNumType comes after pgMar/pgBorders and before cols/docGrid.
 */
export function applyPageNumType(
  sectPrXml: string,
  fmt: string | undefined,
  start: number | undefined,
  chapter?: PageNumberChapter | null,
): string {
  const kept =
    chapter === undefined
      ? (/<w:pgNumType\b[^>]*>/.exec(sectPrXml)?.[0].match(/\sw:chap\w+="[^"]*"/g) ?? []).join('')
      : chapter === null
        ? ''
        : ` w:chapStyle="${chapter.style}"${chapter.sep === 'hyphen' ? '' : ` w:chapSep="${chapter.sep}"`}`
  const xml = stripElement(sectPrXml, 'w:pgNumType')
  if (fmt === undefined && start === undefined && !kept) return xml
  const tag = `<w:pgNumType${fmt !== undefined ? ` w:fmt="${fmt}"` : ''}${start !== undefined ? ` w:start="${start}"` : ''}${kept}/>`
  return insertBefore(xml, tag, PG_NUM_TYPE_FOLLOWERS)
}

/**
 * Insert tag as the first child of a sectPr. A self-closing <w:sectPr/> is
 * expanded to a pair first: the open-tag anchor matched that element whole, so
 * the child was written after it, as a sibling inside w:body / w:pPr.
 */
export function injectIntoSectPr(xml: string, tag: string): string {
  const m = /<w:sectPr(?:\s[^>]*?)?\/>|<w:sectPr(?:\s[^>]*?)?>/.exec(xml)
  if (!m) return xml
  const open = m[0].endsWith('/>') ? `${m[0].slice(0, -2)}>` : m[0]
  return xml.replace(m[0], () => `${open}${tag}${m![0].endsWith('/>') ? '</w:sectPr>' : ''}`)
}

/** CT_SectPr children that follow pgNumType / titlePg, in schema order */
const PG_NUM_TYPE_FOLLOWERS =
  /<w:(?:cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings)[\s/>]/
const LN_NUM_TYPE_FOLLOWERS =
  /<w:(?:pgNumType|cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings)[\s/>]/
const V_ALIGN_FOLLOWERS =
  /<w:(?:noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings)[\s/>]/
const NO_ENDNOTE_FOLLOWERS =
  /<w:(?:titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)[\s/>]/
const TITLE_PG_FOLLOWERS = /<w:(?:textDirection|bidi|rtlGutter|docGrid|printerSettings)[\s/>]/
const TEXT_DIRECTION_FOLLOWERS = /<w:(?:bidi|rtlGutter|docGrid|printerSettings|sectPrChange)[\s/>]/
/** w:docGrid as a file may spell it: self-closing or paired */
const DOC_GRID_ELEMENT = /<w:docGrid\b[^>]*(?:\/>|>\s*<\/w:docGrid>)/
const DOC_GRID_FOLLOWERS = /<w:(?:printerSettings|sectPrChange)[\s/>]/
/** the line pitch Word writes for a new Latin document */
export const DEFAULT_LINE_PITCH_TWIPS = 360

function insertBefore(sectPrXml: string, tag: string, followers: RegExp): string {
  const m = followers.exec(sectPrXml)
  if (m) return `${sectPrXml.slice(0, m.index)}${tag}${sectPrXml.slice(m.index)}`
  return sectPrXml.replace(/<\/w:sectPr>/, `${tag}</w:sectPr>`)
}

/**
 * Enumerate every section in document order. A paragraph-level w:sectPr
 * (section-break paragraph) closes its section; the trailing hidden sectPr closes the last.
 * Block ranges use docxIndex (body child order), boundaries inclusive.
 */
export function readSections(parsed: ParsedDoc): SectionInfo[] {
  const sections: SectionInfo[] = []
  let first = 0
  for (const block of parsed.blocks) {
    const xml = block.originalXml ?? ''
    if (block.docxIndex == null || !xml.includes('<w:sectPr')) continue
    const sectPrXml = SECT_PR_RE.exec(xml)?.[0]
    if (!sectPrXml) continue
    sections.push(sectionFromSectPr(sectPrXml, first, block.docxIndex, parsed.gutterAtTop))
    first = block.docxIndex + 1
  }
  if (sections.length === 0) {
    const last = parsed.blocks[parsed.blocks.length - 1]?.docxIndex ?? 0
    sections.push(sectionFromSectPr('', 0, last))
  }
  return sections
}

/**
 * A kept w:cols element whose values were loaded fractional: round them in
 * place so the saved markup is schema-valid without touching its structure.
 */
function roundColsInPlace(xml: string, cols: string): string {
  const rounded = cols.replace(
    / w:(num|space|w)="(\d+\.\d+)"/g,
    (_m, name: string, v: string) => ` w:${name}="${Math.round(parseFloat(v))}"`,
  )
  return rounded === cols ? xml : xml.replace(cols, rounded)
}

/**
 * w:pgBorders for pageBorderProps. Without props the legacy single-line box
 * (½ pt, 24 pt from the page edge) is written; props with no sides yield ''.
 */
export function pageBordersXml(props: SectionSettings['pageBorderProps']): string {
  const sides = props?.sides
  if (props && !(sides && Object.keys(sides).length > 0)) return ''
  if (!props || !sides) {
    const side = (name: string) =>
      `<w:${name} w:val="single" w:sz="4" w:space="24" w:color="auto"/>`
    return (
      '<w:pgBorders w:offsetFrom="page">' +
      `${side('top')}${side('left')}${side('bottom')}${side('right')}` +
      '</w:pgBorders>'
    )
  }
  let out = '<w:pgBorders'
  if (props.zOrder === 'back') out += ' w:zOrder="back"'
  if (props.display) out += ` w:display="${props.display}"`
  out += ` w:offsetFrom="${props.offsetFrom ?? 'text'}">`
  for (const name of ['top', 'left', 'bottom', 'right'] as const) {
    const side = sides[name]
    if (!side) continue
    const sz = side.art ? Math.round(side.widthPt) : Math.round(side.widthPt * 8)
    out += `<w:${name} w:val="${side.val}" w:sz="${sz}" w:space="${Math.round(side.spacePt)}"`
    out += ` w:color="${side.color ?? 'auto'}"`
    if (side.shadow) out += ' w:shadow="1"'
    if (side.frame) out += ' w:frame="1"'
    out += '/>'
  }
  return out + '</w:pgBorders>'
}

/** Rewrite pgSz / pgMar inside a w:sectPr XML slice, preserving everything else. */
export function applySectionSettings(sectPrXml: string, settings: SectionSettings): string {
  const orient = settings.orientation === 'landscape' ? ' w:orient="landscape"' : ''
  const pgSz = `<w:pgSz w:w="${settings.pageWidth}" w:h="${settings.pageHeight}"${orient}/>`
  let xml = sectPrXml
  if (/<w:pgSz[^>]*\/?>/.test(xml)) {
    xml = xml.replace(/<w:pgSz[^>]*\/?>/, pgSz)
  } else {
    xml = injectIntoSectPr(xml, pgSz)
  }
  const replaceMarAttr = (tag: string, name: string, value: number): string => {
    if (new RegExp(`${name}="`).test(tag)) {
      return tag.replace(new RegExp(`${name}="-?[\\d.]+"`), `${name}="${value}"`)
    }
    return tag.replace(/\/>$/, ` ${name}="${value}"/>`)
  }
  const marMatch = /<w:pgMar[^>]*\/>/.exec(xml)
  if (marMatch) {
    let tag = marMatch[0]
    // the gutter was folded into the parsed margin; a user-set margin below it writes 0, not a negative value
    const gutter = settings.gutter ?? 0
    const top = Math.max(0, settings.marginTop - (settings.gutterAtTop ? gutter : 0))
    tag = replaceMarAttr(tag, 'w:top', settings.marginTopFixed ? -top : top)
    tag = replaceMarAttr(tag, 'w:right', settings.marginRight)
    tag = replaceMarAttr(
      tag,
      'w:bottom',
      settings.marginBottomFixed ? -settings.marginBottom : settings.marginBottom,
    )
    tag = replaceMarAttr(
      tag,
      'w:left',
      Math.max(0, settings.marginLeft - (settings.gutterAtTop ? 0 : gutter)),
    )
    if (settings.headerDist !== undefined)
      tag = replaceMarAttr(tag, 'w:header', settings.headerDist)
    if (settings.footerDist !== undefined)
      tag = replaceMarAttr(tag, 'w:footer', settings.footerDist)
    if (settings.gutter !== undefined) tag = replaceMarAttr(tag, 'w:gutter', gutter)
    xml = xml.replace(marMatch[0], tag)
  } else {
    const gutter = settings.gutter ?? 0
    const pgMar =
      `<w:pgMar w:top="${settings.marginTop - (settings.gutterAtTop ? gutter : 0)}" w:right="${settings.marginRight}"` +
      ` w:bottom="${settings.marginBottom}" w:left="${settings.marginLeft - (settings.gutterAtTop ? 0 : gutter)}"` +
      ` w:header="708" w:footer="708" w:gutter="${gutter}"/>`
    xml = xml.replace(/<\/w:sectPr>/, `${pgMar}</w:sectPr>`)
  }

  // page border (pgBorders sits between pgMar and cols in CT_SectPr)
  xml = xml.replace(/<w:pgBorders[^>]*\/>|<w:pgBorders[\s\S]*?<\/w:pgBorders>/, '')
  if (settings.pageBorder) {
    const pgBorders = pageBordersXml(settings.pageBorderProps)
    if (pgBorders) xml = xml.replace(/(<w:pgMar[^>]*\/>)/, `$1${pgBorders}`)
  }

  // line numbering (w:lnNumType, between pgBorders and pgNumType): rewritten
  // only when it differs from what the section already says
  const sameNumbering = (x?: LineNumbering, y?: LineNumbering): boolean =>
    x === y ||
    (!!x &&
      !!y &&
      x.countBy === y.countBy &&
      x.start === y.start &&
      x.restart === y.restart &&
      x.distance === y.distance)
  if (!sameNumbering(settings.lineNumbers, lineNumberingOf(xml))) {
    xml = xml.replace(/<w:lnNumType\b[^>]*(?:\/>|>\s*<\/w:lnNumType>)/, '')
    const ln = settings.lineNumbers
    if (ln) {
      const attrs =
        ` w:countBy="${ln.countBy}"` +
        (ln.start > 1 ? ` w:start="${ln.start - 1}"` : '') +
        (ln.distance !== undefined ? ` w:distance="${ln.distance}"` : '') +
        ` w:restart="${ln.restart}"`
      xml = insertBefore(xml, `<w:lnNumType${attrs}/>`, LN_NUM_TYPE_FOLLOWERS)
    }
  }

  // columns (w:cols may be self-closing or carry explicit <w:col> children)
  const COLS_ELEMENT = /<w:cols[^>]*\/>|<w:cols[^>]*>[\s\S]*?<\/w:cols>/
  const equalCols = settings.colEqualWidth === true
  const declared = COLS_ELEMENT.exec(xml)
  if (equalCols && declared && (!declared[0].endsWith('/>') || colsUnequal(declared[0]))) {
    const open = /^<w:cols[^>]*>/.exec(declared[0])![0]
    xml = xml.replace(
      declared[0],
      open.replace(/ w:equalWidth=["'][^"']*["']/, '').replace(/\/?>$/, '/>'),
    )
  }
  const colsMatch = COLS_ELEMENT.exec(xml)
  const numAttr = settings.columns > 1 ? ` w:num="${settings.columns}"` : ''
  const colsAnchor = (tag: string): string => {
    // w:cols follows whichever of these the section has last (CT_SectPr order)
    const before = [
      ...xml.matchAll(
        /<w:pgBorders[\s\S]*?<\/w:pgBorders>|<w:(pgMar|lnNumType|pgNumType)\b[^>]*(?:\/>|>\s*<\/w:\1>)/g,
      ),
    ].at(-1)
    if (before) {
      const at = before.index + before[0].length
      return `${xml.slice(0, at)}${tag}${xml.slice(at)}`
    }
    return xml.replace(/<\/w:sectPr>/, `${tag}</w:sectPr>`)
  }
  if (
    !equalCols &&
    settings.colWidths !== undefined &&
    settings.columns > 1 &&
    settings.colWidths.length === settings.columns
  ) {
    // explicit unequal widths: rebuild the element (opt-in via colWidths) —
    // unless the document already carries exactly these values (round-trip)
    const currentTags = colsMatch?.[0].match(/<w:col [^>]*w:w=["'][\d.]+["'][^>]*>/g) ?? []
    const currentWidths = currentTags.map((t) => intAttr(t, 'w:w', 0))
    const space = settings.colSpace ?? 720
    const spaces = settings.colSpaces
    const unchanged =
      colsMatch !== null &&
      intAttr(colsMatch[0], 'w:num', 1) === settings.columns &&
      currentWidths.length === settings.colWidths.length &&
      currentWidths.every((w, i) => w === settings.colWidths![i]) &&
      (spaces === undefined ||
        currentTags.slice(0, -1).every((t, i) => intAttr(t, 'w:space', 0) === (spaces[i] ?? space)))
    if (!unchanged) {
      const children = settings.colWidths
        .map((w, i) =>
          i < settings.colWidths!.length - 1
            ? `<w:col w:w="${w}" w:space="${spaces?.[i] ?? space}"/>`
            : `<w:col w:w="${w}"/>`,
        )
        .join('')
      const tag = `<w:cols${numAttr} w:space="${space}" w:equalWidth="0">${children}</w:cols>`
      xml = colsMatch ? xml.replace(colsMatch[0], tag) : colsAnchor(tag)
    } else if (colsMatch) {
      xml = roundColsInPlace(xml, colsMatch[0])
    }
  } else if (colsMatch) {
    const openTag = /^<w:cols[^>]*>/.exec(colsMatch[0])?.[0] ?? colsMatch[0]
    const selfClosing = colsMatch[0].endsWith('/>')
    const currentNum = String(intAttr(openTag, 'w:num', 1))
    if (selfClosing || currentNum !== String(settings.columns)) {
      // explicit per-column widths only stay valid while the count is unchanged
      let tag = openTag
        .replace(/ w:num="[\d.]+"/, '')
        .replace(/ w:equalWidth=["'][^"']*["']/, '')
        .replace(/\/?>$/, '/>')
      if (numAttr) tag = tag.replace(/^<w:cols/, `<w:cols${numAttr}`)
      // honor an explicitly different column gap (720 = OOXML default); equal
      // values leave the tag byte-identical for round-trip safety
      if (settings.colSpace !== undefined && settings.colSpace !== intAttr(tag, 'w:space', 720)) {
        tag = / w:space="[\d.]+"/.test(tag)
          ? tag.replace(/ w:space="[\d.]+"/, ` w:space="${settings.colSpace}"`)
          : tag.replace(/\/>$/, ` w:space="${settings.colSpace}"/>`)
      }
      xml = xml.replace(colsMatch[0], tag)
    } else {
      xml = roundColsInPlace(xml, colsMatch[0])
    }
  } else if (numAttr) {
    xml = colsAnchor(`<w:cols${numAttr} w:space="${settings.colSpace ?? 425}"/>`)
  }

  // the rule between columns (w:cols w:sep): an unchanged one stays as the file has it
  const colsOpen = /<w:cols\b[^>]*>/.exec(xml)?.[0]
  if (colsOpen && !!settings.colSep !== onOffAttr(colsOpen, 'w:sep')) {
    const bare = colsOpen.replace(/ w:sep=["'][^"']*["']/, '')
    xml = xml.replace(colsOpen, settings.colSep ? bare.replace(/(\/?>)$/, ' w:sep="1"$1') : bare)
  }

  // vertical alignment of the page (w:vAlign, after cols and formProt): top is
  // written by leaving the tag out; an unchanged one stays as the file has it
  const vAlign = settings.vAlign === 'top' ? undefined : settings.vAlign
  if (vAlign !== vAlignOf(xml)) {
    xml = xml.replace(/<w:vAlign\b[^>]*(?:\/>|>\s*<\/w:vAlign>)/, '')
    if (vAlign) xml = insertBefore(xml, `<w:vAlign w:val="${vAlign}"/>`, V_ALIGN_FOLLOWERS)
  }

  // w:noEndnote (after vAlign): an unchanged one stays as the file has it
  if (!!settings.noEndnote !== !!onOffTagIn(xml, 'w:noEndnote')) {
    xml = stripElement(xml, 'w:noEndnote')
    if (settings.noEndnote) xml = insertBefore(xml, '<w:noEndnote/>', NO_ENDNOTE_FOLLOWERS)
  }

  // the section's own note properties: the fields it states that the sectPr does not have yet
  const noteDelta = (tag: 'w:footnotePr' | 'w:endnotePr', own?: NoteProps): NoteProps => {
    const has = notePropsFromXml(liveSectPr(xml), tag)
    return Object.fromEntries(
      NOTE_PR_FIELDS.filter((f) => own?.[f] !== undefined && own[f] !== has?.[f]).map((f) => [
        f,
        own![f],
      ]),
    )
  }
  const footnote = noteDelta('w:footnotePr', settings.footnotePr)
  const endnote = noteDelta('w:endnotePr', settings.endnotePr)
  if (Object.keys(footnote).length + Object.keys(endnote).length > 0)
    xml = applySectionNoteProps(xml, { footnote, endnote })

  // text flow (w:textDirection, after titlePg): horizontal is written by leaving
  // the tag out; an unchanged one stays as the file has it
  const flow = settings.textDirection === 'lrTb' ? undefined : settings.textDirection
  if (flow !== textDirectionOf(xml) && (flow === undefined || /^[A-Za-z]+$/.test(flow))) {
    xml = stripElement(xml, 'w:textDirection')
    if (flow)
      xml = insertBefore(xml, `<w:textDirection w:val="${flow}"/>`, TEXT_DIRECTION_FOLLOWERS)
  }

  // document grid type (w:docGrid w:type, default written by leaving the attribute
  // out): an unchanged one stays as the file has it, and so do its other attributes
  const gridTag = DOC_GRID_ELEMENT.exec(xml)?.[0]
  const gridType = settings.docGrid?.type ?? 'default'
  if (gridType !== ((gridTag && strAttr(gridTag, 'w:type')) || 'default')) {
    const pitch = settings.docGrid?.linePitch ?? DEFAULT_LINE_PITCH_TWIPS
    let tag = (gridTag ?? '<w:docGrid/>')
      .replace(/>\s*<\/w:docGrid>$/, '/>')
      .replace(/ w:type=["'][^"']*["']/, '')
    if (gridType !== 'default') tag = tag.replace('<w:docGrid', `<w:docGrid w:type="${gridType}"`)
    if (!/ w:linePitch=/.test(tag)) tag = tag.replace(/\/>$/, ` w:linePitch="${pitch}"/>`)
    xml = gridTag ? xml.replace(gridTag, tag) : insertBefore(xml, tag, DOC_GRID_FOLLOWERS)
  }
  // the line pitch of a grid the section states: an unchanged one stays as the file has it
  const pitchTag = DOC_GRID_ELEMENT.exec(xml)?.[0]
  const linePitch = settings.docGrid?.linePitch
  if (pitchTag && linePitch !== undefined && linePitch !== intAttr(pitchTag, 'w:linePitch', -1)) {
    const bare = pitchTag
      .replace(/>\s*<\/w:docGrid>$/, '/>')
      .replace(/ w:linePitch=["'][^"']*["']/, '')
    xml = xml.replace(pitchTag, bare.replace(/\/>$/, ` w:linePitch="${linePitch}"/>`))
  }
  // the character pitch of a grid the section states (w:charSpace, 0 written by leaving it out)
  const spaceTag = DOC_GRID_ELEMENT.exec(xml)?.[0]
  const charSpace = settings.docGrid?.charSpace
  if (spaceTag && charSpace !== undefined && charSpace !== intAttr(spaceTag, 'w:charSpace', 0)) {
    const bare = spaceTag
      .replace(/>\s*<\/w:docGrid>$/, '/>')
      .replace(/ w:charSpace=["'][^"']*["']/, '')
    xml = xml.replace(
      spaceTag,
      charSpace === 0 ? bare : bare.replace(/\/>$/, ` w:charSpace="${Math.round(charSpace)}"/>`),
    )
  }

  // section direction (w:bidi, after cols in CT_SectPr): undefined = keep the
  // document's tag untouched; true/false = ensure present/absent
  // w:bidi as Word may write it: self-closing, paired, or carrying an explicit w:val.
  // The reader sees all three spellings, so the writer has to match and remove all
  // three or a round-trip over a paired element leaves a second w:bidi behind.
  const BIDI_ELEMENT = /<w:bidi(?=[\s/>])[^>]*(?:\/>|>[\s\S]*?<\/w:bidi>)/i
  if (settings.bidi !== undefined) {
    const hasBidi = BIDI_ELEMENT.test(xml)
    if (settings.bidi && hasBidi && !onOffTagIn(xml, 'w:bidi')) {
      xml = xml.replace(BIDI_ELEMENT, '<w:bidi/>')
    } else if (settings.bidi && !hasBidi) {
      if (/<w:docGrid/.test(xml)) xml = xml.replace(/(<w:docGrid)/, '<w:bidi/>$1')
      else xml = xml.replace(/<\/w:sectPr>/, '<w:bidi/></w:sectPr>')
    } else if (!settings.bidi && hasBidi) {
      xml = xml.replace(BIDI_ELEMENT, '')
    }
  }
  return xml
}

/**
 * Rewrite the sectPr start type w:type (how this section starts relative to the previous
 * one). nextPage is the default and is written by removing the tag; w:type comes before
 * pgSz in CT_SectPr.
 */
export function applySectionStartType(
  sectPrXml: string,
  type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn',
): string {
  let xml = stripElement(sectPrXml, 'w:type')
  if (type === 'nextPage') return xml
  const tag = `<w:type w:val="${type}"/>`
  if (/<w:pgSz/.test(xml)) xml = xml.replace(/(<w:pgSz)/, `${tag}$1`)
  else xml = injectIntoSectPr(xml, tag)
  return xml
}

/** set or remove w:titlePg (different first page) at its CT_SectPr position */
export function applyTitlePg(sectPrXml: string, on: boolean): string {
  const xml = stripElement(sectPrXml, 'w:titlePg')
  return on ? insertBefore(xml, '<w:titlePg/>', TITLE_PG_FOLLOWERS) : xml
}

/** Read the page color (w:background) from document.xml; null when unset. */
export function readPageColor(parsed: ParsedDoc): string | null {
  const m = /<w:background[^>]*w:color="([0-9A-Fa-f]{6})"/.exec(parsed.internal.documentXml)
  return m ? m[1].toUpperCase() : null
}
