// styles.xml: paragraph / character / table style resolution into the
// ParsedDoc style tables.
import JSZip from 'jszip'

import {
  attrsOf,
  boolProp,
  findChild,
  findChildren,
  nameOf,
  xmlParser,
  type XNode,
} from './xml-utils'
import { colorFrom, lineTwipsOf, onOffOf, stripHash } from './parse-xml-text'
import {
  EA_LANG_DEFAULT_FONT,
  autoSpaceOf,
  cellMarginsOf,
  charIndentsOf,
  mergeCharIndents,
  mergedBorderLinesOf,
  shdDisplayFill,
  tabStopsOf,
  themeLangEaSlotFont,
  themedRFonts,
} from './parse-props'
import type {
  DocDefaults,
  StyleDisplay,
  StyleInfo,
  TableStyleDisplay,
  ThemeColors,
  ThemeFonts,
} from './types'

export async function parseStyles(
  zip: JSZip,
  theme?: ThemeColors | null,
  themeFonts?: ThemeFonts | null,
): Promise<{ styles: Map<string, StyleInfo>; docDefaults?: DocDefaults }> {
  const styles = new Map<string, StyleInfo>()
  const file = zip.file('word/styles.xml')
  if (!file) return { styles }
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(await file.async('string')) as XNode[]
  } catch (err) {
    console.warn('styles.xml unparseable, styles degraded to empty:', err)
    return { styles }
  }
  const root = parsed.find((n) => nameOf(n) === 'w:styles')
  if (!root) return { styles }

  let docDefaults: DocDefaults | undefined
  const defaultsNode = findChild(root, 'w:docDefaults')
  if (defaultsNode) {
    const dd: DocDefaults = {}
    const rPr = findChild(findChild(defaultsNode, 'w:rPrDefault') ?? {}, 'w:rPr')
    const sz = rPr ? attrsOf(findChild(rPr, 'w:sz') ?? {})['w:val'] : undefined
    if (sz) dd.sizeHalfPoints = parseInt(sz, 10) || undefined
    const ddRf = themedRFonts(rPr ? attrsOf(findChild(rPr, 'w:rFonts') ?? {}) : {}, themeFonts)
    if (ddRf.ascii ?? ddRf.hAnsi) dd.asciiFont = ddRf.ascii ?? ddRf.hAnsi
    // docDefaults keeps the lang-based backfill below for the empty-slot case
    if (ddRf.eastAsia && !ddRf.eaSlotEmpty) dd.eastAsiaFont = ddRf.eastAsia
    // Empty EA slot + w:lang w:eastAsia backfill: when the backfill would fire,
    // a face settings.xml themeFontLang resolves (script table / probed locale
    // defaults) outranks the often-stale docDefaults w:lang. Without a firing
    // backfill the slot stays empty — themeFontLang alone must not invent a
    // doc-level EA face (it would reroute PUA/CJK fallback in Latin documents).
    const eaLang = rPr ? attrsOf(findChild(rPr, 'w:lang') ?? {})['w:eastAsia'] : undefined
    if (!dd.eastAsiaFont && eaLang) {
      const eaDefault = EA_LANG_DEFAULT_FONT[eaLang.toLowerCase()]
      if (eaDefault) {
        const ddEaTheme =
          ddRf.eaSlotEmpty && themeFonts
            ? themeLangEaSlotFont(
                themeFonts,
                rPr ? attrsOf(findChild(rPr, 'w:rFonts') ?? {})['w:eastAsiaTheme'] : undefined,
              )
            : undefined
        dd.eastAsiaFont = ddEaTheme ?? eaDefault
        dd.eaFromLang = true
        if (ddRf.eaSlotEmpty) dd.eaSlotEmpty = true
      }
    }
    if (rPr) {
      const onFlag = (tag: string) => {
        const node = findChild(rPr, tag)
        if (!node) return undefined
        const val = attrsOf(node)['w:val']
        return val === '0' || val === 'false' ? undefined : true
      }
      if (onFlag('w:b')) dd.bold = true
      if (onFlag('w:i')) dd.italic = true
      const color = colorFrom(rPr, theme)
      if (color) dd.color = color
      const lang = attrsOf(findChild(rPr, 'w:lang') ?? {})['w:val']
      if (lang) dd.lang = lang
    }
    const pPr = findChild(findChild(defaultsNode, 'w:pPrDefault') ?? {}, 'w:pPr')
    const spacingAttrs = pPr ? attrsOf(findChild(pPr, 'w:spacing') ?? {}) : {}
    if (spacingAttrs['w:line']) {
      const line = lineTwipsOf(spacingAttrs['w:line'])
      const rule = (spacingAttrs['w:lineRule'] ?? 'auto') as 'auto' | 'atLeast' | 'exact'
      if (line > 0) {
        dd.lineRawTwips = line
        dd.lineRule = rule
        if (rule === 'auto') dd.lineSpacing = line / 240
      }
    }
    if (spacingAttrs['w:before'] !== undefined) {
      dd.spaceBeforeTwips = parseInt(spacingAttrs['w:before'], 10) || 0
    }
    if (spacingAttrs['w:after'] !== undefined) {
      dd.spaceAfterTwips = parseInt(spacingAttrs['w:after'], 10) || 0
    }
    if (spacingAttrs['w:beforeAutospacing'] !== undefined)
      dd.spaceBeforeAuto =
        spacingAttrs['w:beforeAutospacing'] === '1' ||
        spacingAttrs['w:beforeAutospacing'] === 'true'
    if (spacingAttrs['w:afterAutospacing'] !== undefined)
      dd.spaceAfterAuto =
        spacingAttrs['w:afterAutospacing'] === '1' || spacingAttrs['w:afterAutospacing'] === 'true'
    if (pPr && onOffOf(pPr, 'w:suppressAutoHyphens')) dd.suppressAutoHyphens = true
    if (Object.keys(dd).length > 0) docDefaults = dd
  }

  const basedOnIds = new Map<string, string>()
  const linkedIds = new Map<string, string>()
  // styles with an explicit w:outlineLvl 9 (body text, e.g. TOCHeading basedOn Heading1)
  const outlineOffIds = new Set<string>()
  for (const styleNode of findChildren(root, 'w:style')) {
    const attrs = attrsOf(styleNode)
    const type = attrs['w:type']
    if (type !== 'paragraph' && type !== 'character' && type !== 'table') continue
    const styleId = attrs['w:styleId']
    if (!styleId) continue
    const name = attrsOf(findChild(styleNode, 'w:name') ?? {})['w:val'] ?? styleId
    let headingLevel: number | undefined
    if (type === 'paragraph') {
      const nameMatch = /^heading\s*([1-9])$/i.exec(name) ?? /^Heading([1-9])$/.exec(styleId)
      if (nameMatch) headingLevel = parseInt(nameMatch[1], 10)
      else {
        const pPr = findChild(styleNode, 'w:pPr')
        const outline = pPr ? attrsOf(findChild(pPr, 'w:outlineLvl') ?? {})['w:val'] : undefined
        if (outline !== undefined) {
          const lvl = parseInt(outline, 10)
          if (lvl >= 0 && lvl <= 8) headingLevel = lvl + 1
          else outlineOffIds.add(styleId)
        }
      }
    }
    const basedOn = attrsOf(findChild(styleNode, 'w:basedOn') ?? {})['w:val']
    if (basedOn) basedOnIds.set(styleId, basedOn)
    const link = attrsOf(findChild(styleNode, 'w:link') ?? {})['w:val']
    if (link) linkedIds.set(styleId, link)
    const onFlag = (tag: string): boolean | undefined => {
      const node = findChild(styleNode, tag)
      if (!node) return undefined
      const val = attrsOf(node)['w:val']
      return val === '0' || val === 'false' ? undefined : true
    }
    let numPr: StyleInfo['numPr']
    if (type === 'paragraph') {
      const styleNumPr = findChild(findChild(styleNode, 'w:pPr') ?? {}, 'w:numPr')
      if (styleNumPr) {
        const numId = attrsOf(findChild(styleNumPr, 'w:numId') ?? {})['w:val']
        if (numId === '0') {
          numPr = 'none'
        } else if (numId) {
          const ilvl = parseInt(attrsOf(findChild(styleNumPr, 'w:ilvl') ?? {})['w:val'] ?? '0', 10)
          numPr = { numId, ilvl: ilvl || 0 }
        }
      }
    }
    styles.set(styleId, {
      styleId,
      name,
      type,
      headingLevel,
      semiHidden: onFlag('w:semiHidden'),
      qFormat: onFlag('w:qFormat'),
      display: type === 'table' ? undefined : styleDisplayOf(styleNode, theme, themeFonts),
      tableDisplay: type === 'table' ? tableStyleDisplayOf(styleNode, theme) : undefined,
      numPr,
      isDefault: attrs['w:default'] === '1' || attrs['w:default'] === 'true' ? true : undefined,
    })
  }

  // Word's effective default per style type: the last w:default="1" wins. When a
  // type declares none, Word does NOT use ECMA-376's first-of-type rule — it falls
  // back to a style id/named "Normal", else to built-in defaults (docDefaults only).
  {
    const declared = new Map<string, StyleInfo>()
    const normalOfType = new Map<string, StyleInfo>()
    for (const info of styles.values()) {
      if (info.isDefault) declared.set(info.type, info)
      if (
        !normalOfType.has(info.type) &&
        (info.styleId.toLowerCase() === 'normal' || info.name.toLowerCase() === 'normal')
      ) {
        normalOfType.set(info.type, info)
      }
      info.isDefault = undefined
    }
    for (const type of new Set([...declared.keys(), ...normalOfType.keys()])) {
      const pick = declared.get(type) ?? normalOfType.get(type)
      if (pick) pick.isDefault = true
    }
  }

  // resolve basedOn chains: a style inherits every display prop it doesn't set itself
  const resolved = new Set<string>()
  const resolve = (styleId: string, seen: Set<string>): StyleInfo | undefined => {
    const info = styles.get(styleId)
    if (!info) return undefined
    const parentId = basedOnIds.get(styleId)
    if (resolved.has(styleId) || !parentId || seen.has(styleId)) return info
    seen.add(styleId)
    const parent = resolve(parentId, seen)
    resolved.add(styleId)
    if (parent?.display) {
      const own = info.display
      info.display = { ...parent.display, ...(own ?? {}) }
      // w:ind character attributes layer per component: a child's twips-only
      // w:ind keeps the parent's *Chars, a child's explicit zero cancels it
      if (parent.display.indentChars && own?.indentChars) {
        info.display.indentChars = mergeCharIndents(parent.display.indentChars, own.indentChars)
      }
      if (Object.keys(info.display).length === 0) info.display = undefined
    }
    if (parent?.tableDisplay) {
      info.tableDisplay = mergeTableDisplay(parent.tableDisplay, info.tableDisplay)
    }
    if (
      info.type === 'paragraph' &&
      info.headingLevel === undefined &&
      !outlineOffIds.has(styleId) &&
      parent?.headingLevel
    ) {
      info.headingLevel = parent.headingLevel
    }
    if (info.type === 'paragraph' && !info.numPr && parent?.numPr) info.numPr = parent.numPr
    return info
  }
  for (const styleId of styles.keys()) resolve(styleId, new Set())

  // linkedStyle (w:link): a paragraph style and a character style form one unit (Word
  // "linked styles"). Fill in run-level display properties in both directions (never
  // overriding a style's own) — the common gap is a character-style shell with no rPr,
  // where all run properties live on the linked paragraph style.
  const RUN_KEYS = [
    'sizeHalfPoints',
    'color',
    'bold',
    'italic',
    'boldCs',
    'italicCs',
    'sizeCsHalfPoints',
    'rtl',
    'underline',
    'strike',
    'font',
    'fontAscii',
    'csFont',
    'caps',
  ] as const
  for (const [fromId, toId] of linkedIds) {
    const a = styles.get(fromId)
    const b = styles.get(toId)
    if (!a || !b) continue
    for (const [self, other] of [
      [a, b],
      [b, a],
    ] as const) {
      if (self.type !== 'character' && self.type !== 'paragraph') continue
      const fill: Partial<StyleDisplay> = {}
      for (const key of RUN_KEYS) {
        if (self.display?.[key] === undefined && other.display?.[key] !== undefined) {
          ;(fill as Record<string, unknown>)[key] = other.display[key]
        }
      }
      if (Object.keys(fill).length > 0) self.display = { ...fill, ...(self.display ?? {}) }
    }
    if (a.type === 'character' && b.type === 'paragraph') a.linkedCharShell = true
    if (b.type === 'character' && a.type === 'paragraph') b.linkedCharShell = true
  }

  return { styles, docDefaults }
}

function mergeTableDisplay(
  parent: TableStyleDisplay,
  child: TableStyleDisplay | undefined,
): TableStyleDisplay | undefined {
  const merged: TableStyleDisplay = { ...parent, ...(child ?? {}) }
  const DEEP = ['wholeTable', 'firstRow', 'firstCol', 'lastCol', 'lastRow', 'paraSpacing'] as const
  for (const key of DEEP) {
    if (parent[key] || child?.[key]) {
      merged[key] = { ...(parent[key] ?? {}), ...(child?.[key] ?? {}) } as never
    }
  }
  return Object.keys(merged).length > 0 ? merged : undefined
}

/** fills / first-row formatting a table style contributes on screen */
function tableStyleDisplayOf(
  styleNode: XNode,
  theme?: ThemeColors | null,
): TableStyleDisplay | undefined {
  const display: TableStyleDisplay = {}
  const shdFill = (node: XNode | undefined): string | undefined => {
    const fill = node ? attrsOf(findChild(node, 'w:shd') ?? {})['w:fill'] : undefined
    return fill && fill !== 'auto' ? stripHash(fill) : undefined
  }
  const baseFill = shdFill(findChild(styleNode, 'w:tcPr'))
  if (baseFill) display.fill = baseFill
  const szHalfOf = (rPr: XNode | undefined): number | undefined => {
    const val = parseInt(attrsOf(findChild(rPr ?? {}, 'w:sz') ?? {})['w:val'] ?? '', 10)
    return val > 0 ? val : undefined
  }
  const styleRPr = findChild(styleNode, 'w:rPr')
  if (styleRPr) {
    const wholeTable: NonNullable<TableStyleDisplay['wholeTable']> = {}
    const color = colorFrom(styleRPr, theme)
    if (color) wholeTable.color = color
    if (boolProp(styleRPr, 'w:b')) wholeTable.bold = true
    if (boolProp(styleRPr, 'w:i')) wholeTable.italic = true
    const sz = szHalfOf(styleRPr)
    if (sz) wholeTable.sizeHalfPoints = sz
    if (Object.keys(wholeTable).length > 0) display.wholeTable = wholeTable
  }
  for (const cond of findChildren(styleNode, 'w:tblStylePr')) {
    const type = attrsOf(cond)['w:type']
    const tcPr = findChild(cond, 'w:tcPr')
    const fill = shdFill(tcPr)
    if (type === 'firstRow' || type === 'firstCol' || type === 'lastCol' || type === 'lastRow') {
      const rPr = findChild(cond, 'w:rPr')
      const fmt: NonNullable<TableStyleDisplay['firstRow']> = {}
      if (fill) fmt.fill = fill
      if (rPr && boolProp(rPr, 'w:b')) fmt.bold = true
      const color = colorFrom(rPr, theme)
      if (color) fmt.color = color
      const sz = szHalfOf(rPr)
      if (sz) fmt.sizeHalfPoints = sz
      if (Object.keys(fmt).length > 0) display[type] = fmt
    } else if (type === 'band1Horz' && fill) {
      display.band1Fill = fill
    } else if (type === 'band2Horz' && fill) {
      display.band2Fill = fill
    }
  }
  const styleTblPr = findChild(styleNode, 'w:tblPr')
  const borders = mergedBorderLinesOf(styleTblPr, 'w:tblBorders', true)
  if (borders) display.borders = borders
  const cellMar = cellMarginsOf(findChild(styleTblPr ?? {}, 'w:tblCellMar'))
  if (cellMar) display.cellMarTwips = cellMar
  const stylePPr = findChild(styleNode, 'w:pPr')
  const jc = attrsOf(findChild(stylePPr ?? {}, 'w:jc') ?? {})['w:val']
  if (jc) display.paraJc = jc
  const stylePPrSpacing = findChild(stylePPr ?? {}, 'w:spacing')
  if (stylePPrSpacing) {
    const a = attrsOf(stylePPrSpacing)
    const ps: NonNullable<TableStyleDisplay['paraSpacing']> = {}
    const before = parseInt(a['w:before'] ?? '', 10)
    if (before >= 0 && a['w:before'] !== undefined) ps.beforeTwips = before
    const after = parseInt(a['w:after'] ?? '', 10)
    if (after >= 0 && a['w:after'] !== undefined) ps.afterTwips = after
    const line = lineTwipsOf(a['w:line'])
    if (line > 0) {
      ps.lineRawTwips = line
      const rule = (a['w:lineRule'] ?? 'auto') as 'auto' | 'atLeast' | 'exact'
      ps.lineRule = rule
      if (rule === 'auto') ps.lineSpacing = Math.round((line / 240) * 100) / 100
    }
    if (Object.keys(ps).length > 0) display.paraSpacing = ps
  }
  return Object.keys(display).length > 0 ? display : undefined
}

/**
 * Style-chain run props under Word's rtl selection (probed, Word for Mac 2026-08):
 * rtl runs read only the Cs twins (w:bCs/w:iCs/w:szCs), non-rtl runs read only the
 * base props — no cross-fallback. Pass the run's cs flag; callers without run
 * context (style gallery previews, caret defaults) pass false (= non-rtl).
 */
export function styleRunFormat(
  display: StyleDisplay | undefined,
  cs: boolean,
): Pick<StyleDisplay, 'bold' | 'italic' | 'sizeHalfPoints'> {
  if (!display) return {}
  return cs
    ? { bold: display.boldCs, italic: display.italicCs, sizeHalfPoints: display.sizeCsHalfPoints }
    : { bold: display.bold, italic: display.italic, sizeHalfPoints: display.sizeHalfPoints }
}

/** display-only formatting the style contributes on screen (Word renders these from styles.xml) */
function styleDisplayOf(
  styleNode: XNode,
  theme?: ThemeColors | null,
  themeFonts?: ThemeFonts | null,
): StyleDisplay | undefined {
  const display: StyleDisplay = {}
  const rPr = findChild(styleNode, 'w:rPr')
  if (rPr) {
    const sz = attrsOf(findChild(rPr, 'w:sz') ?? {})['w:val']
    if (sz) display.sizeHalfPoints = parseInt(sz, 10) || undefined
    const color = colorFrom(rPr, theme)
    if (color) display.color = color
    const bold = onOffOf(rPr, 'w:b')
    if (bold !== undefined) display.bold = bold
    const italic = onOffOf(rPr, 'w:i')
    if (italic !== undefined) display.italic = italic
    // Cs twins carried separately: the consuming run picks the set by its rtl flag
    // (styleRunFormat); consumers without run context read the base props (= non-rtl)
    const boldCs = onOffOf(rPr, 'w:bCs')
    if (boldCs !== undefined) display.boldCs = boldCs
    const italicCs = onOffOf(rPr, 'w:iCs')
    if (italicCs !== undefined) display.italicCs = italicCs
    const szCs = attrsOf(findChild(rPr, 'w:szCs') ?? {})['w:val']
    if (szCs) display.sizeCsHalfPoints = parseInt(szCs, 10) || undefined
    const rtl = onOffOf(rPr, 'w:rtl')
    if (rtl !== undefined) display.rtl = rtl
    const u = attrsOf(findChild(rPr, 'w:u') ?? {})['w:val']
    if (u) display.underline = u !== 'none'
    const strike = onOffOf(rPr, 'w:strike')
    if (strike !== undefined) display.strike = strike
    const rf = themedRFonts(attrsOf(findChild(rPr, 'w:rFonts') ?? {}), themeFonts)
    const font = rf.eastAsia ?? rf.ascii ?? rf.hAnsi
    const fontAscii = rf.ascii ?? rf.hAnsi
    if (fontAscii) display.fontAscii = fontAscii
    if (rf.cs) display.csFont = rf.cs
    if (font) display.font = font
    if (rf.eaSlotEmpty && font && font === rf.eastAsia) display.eaSlotEmpty = true
    const spc = parseInt(attrsOf(findChild(rPr, 'w:spacing') ?? {})['w:val'] ?? '', 10)
    if (spc) display.charSpacingTwips = spc
    const capsOn = onOffOf(rPr, 'w:caps')
    const smallCapsOn = onOffOf(rPr, 'w:smallCaps')
    if (capsOn) display.caps = 'all'
    else if (smallCapsOn) display.caps = 'small'
    else if (capsOn === false || smallCapsOn === false) display.caps = 'none'
    // w:specVanish marks a style separator, not hidden text
    const vanish = onOffOf(rPr, 'w:vanish')
    if (vanish !== undefined && onOffOf(rPr, 'w:specVanish') !== true) display.vanish = vanish
  }
  const pPr = findChild(styleNode, 'w:pPr')
  if (pPr) {
    const spacing = attrsOf(findChild(pPr, 'w:spacing') ?? {})
    const line = lineTwipsOf(spacing['w:line'])
    if (line > 0) {
      const rule = (spacing['w:lineRule'] ?? 'auto') as 'auto' | 'atLeast' | 'exact'
      display.lineRule = rule
      display.lineRawTwips = line
      if (rule === 'auto') {
        display.lineSpacing = line / 240
      }
    }
    if (spacing['w:before'] !== undefined) {
      display.spaceBeforeTwips = parseInt(spacing['w:before'], 10) || 0
    }
    if (spacing['w:after'] !== undefined) {
      display.spaceAfterTwips = parseInt(spacing['w:after'], 10) || 0
    }
    // tri-state so a child style's explicit "0" overrides the basedOn chain's auto
    if (spacing['w:beforeAutospacing'] !== undefined)
      display.spaceBeforeAuto =
        spacing['w:beforeAutospacing'] === '1' || spacing['w:beforeAutospacing'] === 'true'
    if (spacing['w:afterAutospacing'] !== undefined)
      display.spaceAfterAuto =
        spacing['w:afterAutospacing'] === '1' || spacing['w:afterAutospacing'] === 'true'
    if (boolProp(pPr, 'w:keepNext')) display.keepNext = true
    if (boolProp(pPr, 'w:keepLines')) display.keepLines = true
    {
      const pbb = onOffOf(pPr, 'w:pageBreakBefore')
      if (pbb !== undefined) display.pageBreakBefore = pbb
    }
    {
      const sah = onOffOf(pPr, 'w:suppressAutoHyphens')
      if (sah !== undefined) display.suppressAutoHyphens = sah
    }
    {
      // tri-state so a child style's explicit off survives the basedOn merge
      const ctx = onOffOf(pPr, 'w:contextualSpacing')
      if (ctx !== undefined) display.contextualSpacing = ctx
    }
    const autoSpace = autoSpaceOf(pPr)
    if (autoSpace !== undefined) display.autoSpace = autoSpace
    const jc = attrsOf(findChild(pPr, 'w:jc') ?? {})['w:val']
    if (jc === 'center' || jc === 'right' || jc === 'left' || jc === 'justify') display.align = jc
    else if (jc === 'both' || jc === 'distribute') display.align = 'justify'
    const shdDisp = shdDisplayFill(findChild(pPr, 'w:shd'))
    if (shdDisp) display.shadingFill = shdDisp
    const stops = tabStopsOf(pPr)
    if (stops) display.tabStops = stops
    const ind = findChild(pPr, 'w:ind')
    if (ind) {
      const a = attrsOf(ind)
      const left = parseInt(a['w:left'] ?? a['w:start'] ?? '', 10)
      if (Number.isFinite(left) && left !== 0) display.indentLeftTwips = left
      const right = parseInt(a['w:right'] ?? a['w:end'] ?? '', 10)
      if (Number.isFinite(right) && right !== 0) display.indentRightTwips = right
      const firstLine = parseInt(a['w:firstLine'] ?? '', 10)
      const hanging = parseInt(a['w:hanging'] ?? '', 10)
      if (hanging > 0) display.indentFirstLineTwips = -hanging
      else if (firstLine > 0) display.indentFirstLineTwips = firstLine
      // character-unit indents depend on each paragraph's text; kept raw (explicit
      // zeros included, they cancel an inherited value) for the parser
      const chars = charIndentsOf(ind)
      if (chars) display.indentChars = chars
    }
  }
  return Object.keys(display).length > 0 ? display : undefined
}
