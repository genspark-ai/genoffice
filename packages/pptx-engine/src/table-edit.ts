/**
 * Table style editing — surgically patches <a:tblPr> style attributes
 * (firstRow/bandRow/tableStyleId) and each cell's <a:tcPr> fill/borders.
 *
 * Design philosophy: reuse the existing anchor.originalXml surgical-patch pattern,
 * replacing only the <a:tblPr> and <a:tcPr> nodes and keeping all other bytes
 * intact. Tables are "modeled" in this app (not passthrough), so editing is allowed.
 */
import type { Slide } from './types'
import { escapeXmlAttr } from './xml-utils'

/** Table style edit parameters (engine side of the IPC EditTableStyleOp) */
export interface TableStyleEdit {
  /** Whole-block tblPr XML replacement applied directly (used for styleName presets) */
  tblPrXml?: string
  /** Change only the firstRow flag */
  firstRow?: boolean
  /** Change only the bandRow flag */
  bandRow?: boolean
  /** Right-to-left table (tblPr rtl: mirrored grid); false removes the attribute */
  rtl?: boolean
  /** Shading color #RRGGBB or 'none' (<a:solidFill> / <a:noFill> per tc) */
  shadingColor?: string | null
  /** Border color + width (all = all borders; none = clear) */
  borderColor?: string | null
  borderWidthEmu?: number | null
  borderPreset?: 'all' | 'none' | null
  /** Clear all tcPr direct fills/borders (what PowerPoint does when applying a gallery preset; otherwise direct cell formatting hides the style) */
  clearDirectFormatting?: boolean
  /** Cells the shading/border edit applies to, as (row, tc index); undefined = whole table */
  cells?: Array<{ row: number; col: number }>
}

// ── Preset styles ────────────────────────────────────────────────────
// Color presets do not reference PowerPoint built-in style GUIDs (built-ins follow
// the file theme's accent colors, so a "dark blue header" would come out orange in
// an orange theme). Instead, fixed-color custom <a:tblStyle> definitions are
// written into ppt/tableStyles.xml — the button swatch, this app's rendering, and
// PowerPoint all agree.

const inLn = (tag: 'insideH' | 'insideV', color: string) =>
  `<a:${tag}><a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></a:${tag}>`

/** Style XML for one table region (wholeTbl/band1H/firstRow…). */
function partXml(
  tag: string,
  o: { fill?: string; text?: string; bold?: boolean; insideH?: string; insideV?: string },
): string {
  const tx =
    o.text || o.bold
      ? `<a:tcTxStyle${o.bold ? ' b="on"' : ''}>${o.text ? `<a:srgbClr val="${o.text}"/>` : ''}</a:tcTxStyle>`
      : ''
  const bdr =
    o.insideH || o.insideV
      ? `<a:tcBdr>${o.insideH ? inLn('insideH', o.insideH) : ''}${o.insideV ? inLn('insideV', o.insideV) : ''}</a:tcBdr>`
      : ''
  const fill = o.fill
    ? `<a:fill><a:solidFill><a:srgbClr val="${o.fill}"/></a:solidFill></a:fill>`
    : ''
  const tc = bdr || fill ? `<a:tcStyle>${bdr}${fill}</a:tcStyle>` : ''
  return `<a:${tag}>${tx}${tc}</a:${tag}>`
}

/** Fixed-color banded/header style (band = alternate-row shading; header = header fill). */
function customStyle(
  styleId: string,
  name: string,
  o: { band?: string; header?: { fill: string; text: string }; insideH?: string; whole?: string },
): string {
  // CT_TableStyle child order: wholeTbl → band1H → … → firstRow
  return (
    `<a:tblStyle styleId="${styleId}" styleName="${name}">` +
    partXml('wholeTbl', { text: '000000', fill: o.whole, insideH: o.insideH }) +
    (o.band ? partXml('band1H', { fill: o.band }) : '') +
    partXml('firstRow', {
      bold: true,
      ...(o.header ? { fill: o.header.fill, text: o.header.text } : {}),
    }) +
    '</a:tblStyle>'
  )
}

const ID_ZEBRA_BLUE = '{A10FF1CE-0000-4000-9000-000000000001}'
const ID_ZEBRA_GRAY = '{A10FF1CE-0000-4000-9000-000000000002}'
const ID_HEADER_DARKBLUE = '{A10FF1CE-0000-4000-9000-000000000003}'
const ID_HEADER_ORANGE = '{A10FF1CE-0000-4000-9000-000000000004}'
const ID_NO_BORDER = '{A10FF1CE-0000-4000-9000-000000000005}'

export interface TableStylePreset {
  tblPrXml: string
  description: string
  /** Custom style: must first be ensured into ppt/tableStyles.xml */
  styleId?: string
  styleDefXml?: string
  /** Grid presets: add border lines to every cell directly (incl. outer frame; the style mechanism only covers inner lines) */
  border?: { color: string; widthEmu: number }
}

const tblPr = (styleId: string, flags = '') =>
  `<a:tblPr${flags}><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>`
// Official "No Style, No Grid" GUID (we previously misused a made-up ...307D; the read side stays compatible with it)
const NO_STYLE = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'

/** The 8 preset styles (keys map to the ribbon style gallery). */
export const TABLE_STYLE_PRESETS: Record<string, TableStylePreset> = {
  none: { tblPrXml: tblPr(NO_STYLE), description: 'No style' },
  lightGrid: {
    tblPrXml: tblPr(NO_STYLE),
    description: 'Light grid',
    border: { color: '#BFBFBF', widthEmu: 12700 },
  },
  zebraBlue: {
    tblPrXml: tblPr(ID_ZEBRA_BLUE, ' firstRow="1" bandRow="1"'),
    description: 'Banded blue',
    styleId: ID_ZEBRA_BLUE,
    styleDefXml: customStyle(ID_ZEBRA_BLUE, 'Banded blue', {
      whole: 'FFFFFF',
      band: 'D6E4F0',
      header: { fill: '4472C4', text: 'FFFFFF' },
    }),
  },
  zebraGray: {
    tblPrXml: tblPr(ID_ZEBRA_GRAY, ' firstRow="1" bandRow="1"'),
    description: 'Banded gray',
    styleId: ID_ZEBRA_GRAY,
    styleDefXml: customStyle(ID_ZEBRA_GRAY, 'Banded gray', {
      whole: 'FFFFFF',
      band: 'EDEDED',
      header: { fill: '595959', text: 'FFFFFF' },
    }),
  },
  headerDarkBlue: {
    tblPrXml: tblPr(ID_HEADER_DARKBLUE, ' firstRow="1"'),
    description: 'Dark blue header',
    styleId: ID_HEADER_DARKBLUE,
    styleDefXml: customStyle(ID_HEADER_DARKBLUE, 'Dark blue header', {
      whole: 'FFFFFF',
      insideH: 'D9D9D9',
      band: 'E9EDF5',
      header: { fill: '1F3864', text: 'FFFFFF' },
    }),
  },
  headerOrange: {
    tblPrXml: tblPr(ID_HEADER_ORANGE, ' firstRow="1"'),
    description: 'Orange header',
    styleId: ID_HEADER_ORANGE,
    styleDefXml: customStyle(ID_HEADER_ORANGE, 'Orange header', {
      whole: 'FFFFFF',
      insideH: 'D9D9D9',
      band: 'FBE5D6',
      header: { fill: 'ED7D31', text: 'FFFFFF' },
    }),
  },
  noBorder: {
    tblPrXml: tblPr(ID_NO_BORDER, ' firstRow="1" bandRow="1"'),
    description: 'Minimal (no borders)',
    styleId: ID_NO_BORDER,
    styleDefXml: customStyle(ID_NO_BORDER, 'Minimal (no borders)', { band: 'F2F2F2' }),
  },
  fullBorder: {
    tblPrXml: tblPr(NO_STYLE),
    description: 'All borders',
    border: { color: '#000000', widthEmu: 12700 },
  },
}

const EMPTY_TABLE_STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'

/** Ensure the given style definition exists in tableStyles.xml; returned as-is if already present. */
export function ensureTableStyleXml(
  tableStylesXml: string | null | undefined,
  styleId: string,
  styleDefXml: string,
): string {
  const xml = tableStylesXml?.includes('<a:tblStyleLst') ? tableStylesXml : EMPTY_TABLE_STYLES_XML
  if (xml.includes(styleId)) return xml
  const sc = /<a:tblStyleLst([^>]*)\/>/.exec(xml)
  if (sc) return xml.replace(sc[0], `<a:tblStyleLst${sc[1]}>${styleDefXml}</a:tblStyleLst>`)
  return xml.replace('</a:tblStyleLst>', `${styleDefXml}</a:tblStyleLst>`)
}

/**
 * Surgically patch a table element's originalXml (graphicFrame fragment), applying
 * the style edit. Returns the patched XML; the caller writes it back to
 * anchor.originalXml and sets dirty = true.
 */
export function patchTableStyleXml(originalXml: string, edit: TableStyleEdit): string {
  let xml = originalXml

  // ── 1. Replace <a:tblPr> ──────────────────────────────────────────
  if (edit.tblPrXml !== undefined) {
    // Whole-block replacement. Presets carry their own firstRow/bandRow but never rtl —
    // table direction is orthogonal state and survives a style change (PowerPoint behavior)
    let next = edit.tblPrXml
    const rtl = /<a:tblPr[^>]*\srtl="([^"]*)"/.exec(xml)?.[1]
    if ((rtl === '1' || rtl === 'true') && !/\srtl="/.test(next)) {
      next = next.replace(/^<a:tblPr/, '<a:tblPr rtl="1"')
    }
    xml = replaceTblPr(xml, next)
  } else if (edit.firstRow !== undefined || edit.bandRow !== undefined || edit.rtl !== undefined) {
    // Change only the flags, keeping the rest (styleId etc.)
    const tblPrMatch = /<a:tblPr(\s[^>]*)?(\/?>)/s.exec(xml)
    if (tblPrMatch) {
      // With attributes present the greedy [^>]* swallows a trailing "/" into group 1,
      // so self-closing must be detected on the whole tag and the slash stripped
      const selfClosing = tblPrMatch[0].endsWith('/>')
      let attrs = (tblPrMatch[1] ?? '').replace(/\/\s*$/, '')
      if (edit.firstRow !== undefined)
        attrs = setAttr(attrs, 'firstRow', edit.firstRow ? '1' : undefined)
      if (edit.bandRow !== undefined)
        attrs = setAttr(attrs, 'bandRow', edit.bandRow ? '1' : undefined)
      if (edit.rtl !== undefined) attrs = setAttr(attrs, 'rtl', edit.rtl ? '1' : undefined)
      // Self-closing expands into <a:tblPr...></a:tblPr>
      xml =
        xml.slice(0, tblPrMatch.index) +
        `<a:tblPr${attrs}>` +
        (selfClosing ? '</a:tblPr>' : '') +
        xml.slice(tblPrMatch.index + tblPrMatch[0].length)
    }
  }

  // ── 2. Shading / borders: patch the targeted <a:tcPr> nodes ────────
  if (
    edit.shadingColor !== undefined ||
    edit.borderPreset !== undefined ||
    edit.clearDirectFormatting
  ) {
    xml = edit.cells ? patchCellsTcPr(xml, edit) : patchAllTcPr(xml, edit)
  }

  return xml
}

/** Replace <a:tblPr>…</a:tblPr> (or its self-closing form) in the XML with a new value. */
function replaceTblPr(xml: string, newTblPr: string): string {
  // Self-closing
  const sc = /<a:tblPr[^>]*\/>/.exec(xml)
  if (sc) return xml.slice(0, sc.index) + newTblPr + xml.slice(sc.index + sc[0].length)
  // Open/close tag pair
  const open = xml.indexOf('<a:tblPr')
  if (open < 0) {
    // No tblPr: insert right after <a:tbl>
    const tbl = xml.indexOf('<a:tbl>')
    if (tbl >= 0)
      return xml.slice(0, tbl + '<a:tbl>'.length) + newTblPr + xml.slice(tbl + '<a:tbl>'.length)
    return xml
  }
  const close = xml.indexOf('</a:tblPr>', open)
  if (close < 0) return xml
  return xml.slice(0, open) + newTblPr + xml.slice(close + '</a:tblPr>'.length)
}

/** Set or remove key="value" in an XML attribute string (remove: value=undefined). */
function setAttr(attrs: string, key: string, value: string | undefined): string {
  const re = new RegExp(`\\s${key}="[^"]*"`)
  const cleaned = attrs.replace(re, '')
  if (value === undefined) return cleaned
  return `${cleaned} ${key}="${escapeXmlAttr(value)}"`
}

/** Apply the shading/border/clear parts of an edit to one tcPr's children. */
/** One ln* element in either form — self-closing or paired (they cannot nest). */
const LN_ELEMENT_RE =
  /<a:(lnL|lnR|lnT|lnB|lnTlToBr|lnBlToTr)(?:\s[^>]*)?\/>|<a:(lnL|lnR|lnT|lnB|lnTlToBr|lnBlToTr)(?:\s[^>]*)?>.*?<\/a:\2>/gs

/** Transform only the segments outside ln* elements: a:ln* carry their own
    solidFill children, so cell-fill edits must not reach inside them. */
function outsideLns(src: string, fn: (seg: string) => string): string {
  let out = ''
  let cursor = 0
  for (const m of src.matchAll(LN_ELEMENT_RE)) {
    out += fn(src.slice(cursor, m.index)) + m[0]
    cursor = m.index + m[0].length
  }
  return out + fn(src.slice(cursor))
}

function applyTcPrEdit(inner: string, edit: TableStyleEdit): string {
  if (edit.clearDirectFormatting) {
    inner = inner.replace(/<a:(solidFill|gradFill|pattFill|blipFill)(?:\s[^>]*)?>.*?<\/a:\1>/gs, '')
    inner = inner.replace(/<a:(noFill|grpFill)\/>/g, '')
    inner = inner.replace(/<a:(lnL|lnR|lnT|lnB|lnTlToBr|lnBlToTr)(\s[^>]*)?\/>/g, '')
    inner = inner.replace(/<a:(lnL|lnR|lnT|lnB|lnTlToBr|lnBlToTr)(\s[^>]*)?>.*?<\/a:\1>/gs, '')
  }
  // CT_TableCellProperties is a sequence: ln* → cell3D → fill → headers/extLst.
  // Misplaced children make the part schema-invalid and PowerPoint offers repair.
  const insertFill = (fill: string) => {
    const at = inner.search(/<a:(headers|extLst)[\s>/]/)
    return at >= 0 ? inner.slice(0, at) + fill + inner.slice(at) : inner + fill
  }
  const insertLns = (lns: string) => {
    const at = inner.search(
      /<a:(lnTlToBr|lnBlToTr|cell3D|solidFill|gradFill|pattFill|blipFill|noFill|grpFill|headers|extLst)[\s>/]/,
    )
    return at >= 0 ? inner.slice(0, at) + lns + inner.slice(at) : inner + lns
  }
  // Replace existing cell-level fill nodes (border fills stay untouched)
  if (edit.shadingColor !== undefined) {
    inner = outsideLns(inner, (seg) =>
      seg
        .replace(/<a:(solidFill|gradFill|pattFill|blipFill)(?:\s[^>]*)?>.*?<\/a:\1>/gs, '')
        .replace(/<a:(noFill|grpFill)\/>/g, '')
        .replace(/<a:(noFill|grpFill)><\/a:\1>/g, ''),
    )
    if (edit.shadingColor === 'none') {
      inner = insertFill('<a:noFill/>')
    } else if (edit.shadingColor) {
      const c = edit.shadingColor.replace('#', '').toUpperCase()
      inner = insertFill(`<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>`)
    }
  }
  // Border handling
  if (edit.borderPreset !== undefined) {
    inner = inner.replace(/<a:(lnL|lnR|lnT|lnB)(?:\s[^>]*)?\/>/g, '')
    inner = inner.replace(/<a:(lnL|lnR|lnT|lnB)(?:\s[^>]*)?>.*?<\/a:\1>/gs, '')
    if (edit.borderPreset === 'all' && edit.borderColor) {
      const c = edit.borderColor.replace('#', '').toUpperCase()
      const w = edit.borderWidthEmu ?? 12700
      const lnXml = (tag: string) =>
        `<${tag} w="${w}"><a:solidFill><a:srgbClr val="${c}"/></a:solidFill></${tag}>`
      inner = insertLns(lnXml('a:lnL') + lnXml('a:lnR') + lnXml('a:lnT') + lnXml('a:lnB'))
    } else if (edit.borderPreset === 'none') {
      inner = insertLns(
        '<a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT><a:lnB w="0"><a:noFill/></a:lnB>',
      )
    }
  }
  return inner
}

/** Patch the fill/borders of every <a:tcPr> in the XML. */
function patchAllTcPr(xml: string, edit: TableStyleEdit): string {
  const out: string[] = []
  let cursor = 0
  const re = /<a:tcPr([^>]*)>(.*?)<\/a:tcPr>|<a:tcPr([^>]*)\/>/gs
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null) {
    out.push(xml.slice(cursor, m.index))
    // Existing children (non-self-closing form)
    const inner = applyTcPrEdit(m[2] ?? '', edit)
    const attrs = m[1] ?? m[3] ?? ''
    out.push(`<a:tcPr${attrs}>${inner}</a:tcPr>`)
    cursor = m.index + m[0].length
  }
  out.push(xml.slice(cursor))
  return out.join('')
}

/** Span of the nth <tag> element in [from,to); only for tags that cannot nest (a:tr / a:tc). */
function nthSpan(
  xml: string,
  tag: string,
  n: number,
  from = 0,
  to = xml.length,
): { start: number; end: number } | null {
  const re = new RegExp(`<${tag}[\\s>]`, 'g')
  re.lastIndex = from
  let i = 0
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) !== null && m.index < to) {
    if (i === n) {
      const close = xml.indexOf(`</${tag}>`, m.index)
      if (close < 0 || close + tag.length + 3 > to) return null
      return { start: m.index, end: close + tag.length + 3 }
    }
    i++
  }
  return null
}

/** Patch the fill/borders of the <a:tcPr> of just the cells in edit.cells (creating a tcPr when a cell has none). */
function patchCellsTcPr(xml: string, edit: TableStyleEdit): string {
  // Back to front so earlier spans stay valid while splicing
  const cells = [...edit.cells!].sort((a, b) => b.row - a.row || b.col - a.col)
  for (const { row, col } of cells) {
    if (row < 0 || col < 0) continue
    const tr = nthSpan(xml, 'a:tr', row)
    if (!tr) continue
    const tc = nthSpan(xml, 'a:tc', col, tr.start, tr.end)
    if (!tc) continue
    let tcXml = xml.slice(tc.start, tc.end)
    const m = /<a:tcPr([^>]*)>(.*?)<\/a:tcPr>|<a:tcPr([^>]*)\/>/s.exec(tcXml)
    if (m) {
      const inner = applyTcPrEdit(m[2] ?? '', edit)
      const attrs = m[1] ?? m[3] ?? ''
      tcXml =
        tcXml.slice(0, m.index) +
        `<a:tcPr${attrs}>${inner}</a:tcPr>` +
        tcXml.slice(m.index + m[0].length)
    } else {
      // tcPr follows txBody in CT_TableCell
      tcXml = tcXml.replace(/<\/a:tc>$/, `<a:tcPr>${applyTcPrEdit('', edit)}</a:tcPr></a:tc>`)
    }
    xml = xml.slice(0, tc.start) + tcXml + xml.slice(tc.end)
  }
  return xml
}

/** Find the table element with the given sourceId in the slide; returns its anchor.originalXml location. */
export function findTableElementInSlide(
  slide: Slide,
  sourceId: string,
): import('./types').TableElement | null {
  for (const el of slide.elements) {
    if (el.id === sourceId && el.type === 'table') {
      return el as import('./types').TableElement
    }
  }
  return null
}
