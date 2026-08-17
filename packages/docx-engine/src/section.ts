import type { ParsedDoc, SectionInfo, SectionSettings, DocGrid } from './types'

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
  const v = /<w:vAlign w:val="(center|both|bottom)"\s*\/>/.exec(xml)?.[1]
  return v as 'center' | 'both' | 'bottom' | undefined
}

function intAttr(tag: string, name: string, fallback: number): number {
  const m = new RegExp(`${name}="(-?\\d+)"`).exec(tag)
  const v = m ? parseInt(m[1], 10) : NaN
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
    const val = /w:val="([^"]*)"/.exec(side)?.[1]
    return val !== undefined && val !== 'none' && val !== 'nil'
  })
}

/** Page setup from one w:sectPr XML slice. */
export function sectionSettingsFromXml(xml: string): SectionSettings {
  const pgSz = /<w:pgSz[^>]*\/?>/.exec(xml)?.[0] ?? ''
  const pgMar = /<w:pgMar[^>]*\/?>/.exec(xml)?.[0] ?? ''

  // Parse docGrid (w:docGrid)
  let docGrid: DocGrid | undefined
  const docGridTag = /<w:docGrid[^>]*\/?>/.exec(xml)?.[0]
  if (docGridTag) {
    const typeMatch = /w:type="([^"]+)"/.exec(docGridTag)
    const linePitchMatch = /w:linePitch="(\d+)"/.exec(docGridTag)
    const charSpaceMatch = /w:charSpace="(-?\d+)"/.exec(docGridTag)
    const gridType = (typeMatch?.[1] ?? 'default') as DocGrid['type']
    const validTypes: DocGrid['type'][] = ['default', 'lines', 'linesAndChars', 'snapToChars']
    docGrid = {
      type: validTypes.includes(gridType) ? gridType : 'default',
      ...(linePitchMatch ? { linePitch: parseInt(linePitchMatch[1], 10) } : {}),
      ...(charSpaceMatch ? { charSpace: parseInt(charSpaceMatch[1], 10) } : {}),
    }
  }

  return {
    pageWidth: intAttr(pgSz, 'w:w', DEFAULT_SECTION.pageWidth),
    pageHeight: intAttr(pgSz, 'w:h', DEFAULT_SECTION.pageHeight),
    orientation: pgSz.includes('w:orient="landscape"') ? 'landscape' : 'portrait',
    marginTop: intAttr(pgMar, 'w:top', DEFAULT_SECTION.marginTop),
    marginRight: intAttr(pgMar, 'w:right', DEFAULT_SECTION.marginRight),
    marginBottom: intAttr(pgMar, 'w:bottom', DEFAULT_SECTION.marginBottom),
    marginLeft: intAttr(pgMar, 'w:left', DEFAULT_SECTION.marginLeft),
    headerDist: intAttr(pgMar, 'w:header', 720),
    footerDist: intAttr(pgMar, 'w:footer', 720),
    ...(vAlignOf(xml) ? { vAlign: vAlignOf(xml) } : {}),
    pageBorder: hasVisiblePageBorder(xml),
    columns: intAttr(/<w:cols[^>]*\/?>/.exec(xml)?.[0] ?? '', 'w:num', 1),
    colSpace: intAttr(/<w:cols[^>]*\/?>/.exec(xml)?.[0] ?? '', 'w:space', 720),
    ...(docGrid ? { docGrid } : {}),
    ...(textDirectionOf(xml) ? { textDirection: textDirectionOf(xml) } : {}),
  }
}

function textDirectionOf(xml: string): string | undefined {
  const val = /<w:textDirection[^>]*w:val="([^"]+)"/.exec(xml)?.[1]
  return val && val !== 'lrTb' ? val : undefined
}

/** Read page setup from the trailing (hidden) w:sectPr. */
export function readSectionSettings(parsed: ParsedDoc): SectionSettings {
  const sectBlock = parsed.blocks.find((b) => b.hidden && b.originalXml?.includes('<w:sectPr'))
  return sectionSettingsFromXml(sectBlock?.originalXml ?? '')
}

const SECT_PR_RE = /<w:sectPr[^>]*\/>|<w:sectPr[\s\S]*?<\/w:sectPr>/

/** On/off flag element (w:titlePg, w:evenAndOddHeaders, …) present and not switched
 *  off: matches self-closing and paired forms, w:val="0|false|off" counts as off. */
export function xmlFlagOn(xml: string, tag: string): boolean {
  for (const m of xml.matchAll(new RegExp(`<${tag}(?=[\\s/>])[^>]*>`, 'g'))) {
    const val = /w:val="([^"]*)"/.exec(m[0])?.[1]
    if (val === undefined || !/^(?:0|false|off)$/.test(val)) return true
  }
  return false
}

function hfRefs(
  xml: string,
  kind: 'header' | 'footer',
): Partial<Record<'default' | 'first' | 'even', string>> {
  const refs: Partial<Record<'default' | 'first' | 'even', string>> = {}
  for (const ref of xml.match(new RegExp(`<w:${kind}Reference[^>]*/>`, 'g')) ?? []) {
    const type = /w:type="(default|first|even)"/.exec(ref)?.[1] ?? 'default'
    const rId = /r:id="([^"]+)"/.exec(ref)?.[1]
    if (rId) refs[type as 'default' | 'first' | 'even'] = rId
  }
  return refs
}

function sectionFromSectPr(
  sectPrXml: string,
  firstBlockIndex: number,
  lastBlockIndex: number,
): SectionInfo {
  const type = /<w:type[^>]*w:val="(nextPage|continuous|evenPage|oddPage|nextColumn)"/.exec(
    sectPrXml,
  )?.[1]
  const pgNumStart = /<w:pgNumType[^>]*w:start="(\d+)"/.exec(sectPrXml)?.[1]
  const pgNumFmt = /<w:pgNumType[^>]*w:fmt="([^"]+)"/.exec(sectPrXml)?.[1]
  return {
    settings: sectionSettingsFromXml(sectPrXml),
    // nextColumn is rare; treat it as continuous (no forced page break)
    startType:
      type === 'nextColumn' ? 'continuous' : ((type as SectionInfo['startType']) ?? 'nextPage'),
    firstBlockIndex,
    lastBlockIndex,
    sectPrXml,
    titlePg: xmlFlagOn(sectPrXml, 'w:titlePg'),
    ...(pgNumStart !== undefined ? { pageNumberStart: parseInt(pgNumStart, 10) } : {}),
    ...(pgNumFmt !== undefined ? { pageNumberFmt: pgNumFmt } : {}),
    headerRefs: hfRefs(sectPrXml, 'header'),
    footerRefs: hfRefs(sectPrXml, 'footer'),
  }
}

/**
 * Rewrite the sectPr page numbering w:pgNumType (fmt = number format, start = starting
 * page number; undefined fields are omitted, and the tag is removed when both are unset).
 * Schema order: pgNumType comes after pgMar/pgBorders and before cols/docGrid.
 */
export function applyPageNumType(
  sectPrXml: string,
  fmt: string | undefined,
  start: number | undefined,
): string {
  let xml = sectPrXml.replace(/<w:pgNumType[^>]*\/>/, '')
  if (fmt === undefined && start === undefined) return xml
  const tag = `<w:pgNumType${fmt !== undefined ? ` w:fmt="${fmt}"` : ''}${start !== undefined ? ` w:start="${start}"` : ''}/>`
  if (/<w:cols[\s/>]/.test(xml)) xml = xml.replace(/(<w:cols[\s/>])/, `${tag}$1`)
  else if (/<w:docGrid/.test(xml)) xml = xml.replace(/(<w:docGrid)/, `${tag}$1`)
  else xml = xml.replace(/<\/w:sectPr>/, `${tag}</w:sectPr>`)
  return xml
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
    sections.push(sectionFromSectPr(sectPrXml, first, block.docxIndex))
    first = block.docxIndex + 1
  }
  if (sections.length === 0) {
    const last = parsed.blocks[parsed.blocks.length - 1]?.docxIndex ?? 0
    sections.push(sectionFromSectPr('', 0, last))
  }
  return sections
}

/** Rewrite pgSz / pgMar inside a w:sectPr XML slice, preserving everything else. */
export function applySectionSettings(sectPrXml: string, settings: SectionSettings): string {
  const orient = settings.orientation === 'landscape' ? ' w:orient="landscape"' : ''
  const pgSz = `<w:pgSz w:w="${settings.pageWidth}" w:h="${settings.pageHeight}"${orient}/>`
  let xml = sectPrXml
  if (/<w:pgSz[^>]*\/?>/.test(xml)) {
    xml = xml.replace(/<w:pgSz[^>]*\/?>/, pgSz)
  } else {
    xml = xml.replace(/(<w:sectPr[^>]*>)/, `$1${pgSz}`)
  }
  const replaceMarAttr = (tag: string, name: string, value: number): string => {
    if (new RegExp(`${name}="`).test(tag)) {
      return tag.replace(new RegExp(`${name}="-?\\d+"`), `${name}="${value}"`)
    }
    return tag.replace(/\/>$/, ` ${name}="${value}"/>`)
  }
  const marMatch = /<w:pgMar[^>]*\/>/.exec(xml)
  if (marMatch) {
    let tag = marMatch[0]
    tag = replaceMarAttr(tag, 'w:top', settings.marginTop)
    tag = replaceMarAttr(tag, 'w:right', settings.marginRight)
    tag = replaceMarAttr(tag, 'w:bottom', settings.marginBottom)
    tag = replaceMarAttr(tag, 'w:left', settings.marginLeft)
    xml = xml.replace(marMatch[0], tag)
  } else {
    const pgMar =
      `<w:pgMar w:top="${settings.marginTop}" w:right="${settings.marginRight}"` +
      ` w:bottom="${settings.marginBottom}" w:left="${settings.marginLeft}"` +
      ' w:header="708" w:footer="708" w:gutter="0"/>'
    xml = xml.replace(/<\/w:sectPr>/, `${pgMar}</w:sectPr>`)
  }

  // page border: single-line box measured from the page edge (pgBorders sits
  // between pgMar and cols in CT_SectPr)
  xml = xml.replace(/<w:pgBorders[^>]*\/>|<w:pgBorders[\s\S]*?<\/w:pgBorders>/, '')
  if (settings.pageBorder) {
    const side = (name: string) =>
      `<w:${name} w:val="single" w:sz="4" w:space="24" w:color="auto"/>`
    const pgBorders =
      '<w:pgBorders w:offsetFrom="page">' +
      `${side('top')}${side('left')}${side('bottom')}${side('right')}` +
      '</w:pgBorders>'
    xml = xml.replace(/(<w:pgMar[^>]*\/>)/, `$1${pgBorders}`)
  }

  // columns (w:cols may be self-closing or carry explicit <w:col> children)
  const colsMatch = /<w:cols[^>]*\/>|<w:cols[^>]*>[\s\S]*?<\/w:cols>/.exec(xml)
  const numAttr = settings.columns > 1 ? ` w:num="${settings.columns}"` : ''
  if (colsMatch) {
    const openTag = /^<w:cols[^>]*>/.exec(colsMatch[0])?.[0] ?? colsMatch[0]
    const selfClosing = colsMatch[0].endsWith('/>')
    const currentNum = / w:num="(\d+)"/.exec(openTag)?.[1] ?? '1'
    if (selfClosing || currentNum !== String(settings.columns)) {
      // explicit per-column widths only stay valid while the count is unchanged
      let tag = openTag.replace(/ w:num="\d+"/, '').replace(/\/?>$/, '/>')
      if (numAttr) tag = tag.replace(/^<w:cols/, `<w:cols${numAttr}`)
      xml = xml.replace(colsMatch[0], tag)
    }
  } else if (numAttr) {
    const anchor = /(<w:pgBorders[\s\S]*?<\/w:pgBorders>|<w:pgMar[^>]*\/>)/.exec(xml)
    const cols = `<w:cols${numAttr} w:space="425"/>`
    if (anchor) xml = xml.replace(anchor[0], `${anchor[0]}${cols}`)
    else xml = xml.replace(/<\/w:sectPr>/, `${cols}</w:sectPr>`)
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
  type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage',
): string {
  let xml = sectPrXml.replace(/<w:type[^>]*\/>/, '')
  if (type === 'nextPage') return xml
  const tag = `<w:type w:val="${type}"/>`
  if (/<w:pgSz/.test(xml)) xml = xml.replace(/(<w:pgSz)/, `${tag}$1`)
  else xml = xml.replace(/(<w:sectPr[^>]*>)/, `$1${tag}`)
  return xml
}

/** Read the page color (w:background) from document.xml; null when unset. */
export function readPageColor(parsed: ParsedDoc): string | null {
  const m = /<w:background[^>]*w:color="([0-9A-Fa-f]{6})"/.exec(parsed.internal.documentXml)
  return m ? m[1].toUpperCase() : null
}
