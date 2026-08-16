/**
 * Table style resolution — tableStyleId referenced by a:tblPr → final fill/text styles
 * for each table region.
 *
 * Two sources:
 * 1. <a:tblStyle> in ppt/tableStyles.xml (custom styles embedded in the file);
 * 2. PowerPoint built-in styles (tableStyles.xml usually has only the def id, no body):
 *    All 74 built-in GUIDs are generated from family patterns (Themed 1-2 / Light 1-3 /
 *    Medium 1-4 / Dark 1-2 x no-accent + Accent1-6). The GUID table comes from official
 *    Microsoft docs; family color rules are converted against the LibreOffice oox
 *    predefined-table-styles implementation; banded colors with alpha are pre-blended
 *    into solid colors over the white/table background. Unknown GUIDs return undefined
 *    (keeping the unstyled rendering).
 */
import { XMLParser } from 'fast-xml-parser'
import { resolveSchemeColor, type Theme } from './theme'
import type { Fill, Stroke } from './types'
import { asXmlNode, xmlArray, type XmlNode } from './xml-utils'

/** Style for one table region (whole table / banding / first row, etc.). */
export interface TablePartStyle {
  fill?: Fill
  bold?: boolean
  textColor?: string
}

export interface TableStyleDef {
  wholeTbl?: TablePartStyle
  band1H?: TablePartStyle
  band2H?: TablePartStyle
  band1V?: TablePartStyle
  band2V?: TablePartStyle
  firstRow?: TablePartStyle
  lastRow?: TablePartStyle
  firstCol?: TablePartStyle
  lastCol?: TablePartStyle
  /** Inner separator lines between cells (lt1 white lines for the Medium family) */
  insideH?: Stroke
  insideV?: Stroke
  /** Four sides of the whole-table outer border */
  outer?: { l?: Stroke; r?: Stroke; t?: Stroke; b?: Stroke }
  /** First-row bottom edge / last-row top edge (header separator lines, used when the firstRow/lastRow flags are on) */
  firstRowBottom?: Stroke
  lastRowTop?: Stroke
  /** <a:tblBg> direct fill: whole-table background under transparent/alpha cell fills */
  tblBg?: Fill
  /** <a:tblBg><a:fillRef>: theme fillStyleLst template ref (parse.ts instantiates it with phClr) */
  tblBgRef?: { idx: number; phClr?: string }
}

/** Region toggles from a:tblPr. */
export interface TableStyleFlags {
  firstRow: boolean
  lastRow: boolean
  firstCol: boolean
  lastCol: boolean
  bandRow: boolean
  bandCol: boolean
}

type BuiltinFamily =
  | 'themed1'
  | 'themed2'
  | 'light1'
  | 'light2'
  | 'light3'
  | 'medium1'
  | 'medium2'
  | 'medium3'
  | 'medium4'
  | 'dark1'
  | 'dark2'

// 74 built-in style GUIDs -> family + accent (GUID table from official Microsoft docs hh273476).
// themed1 without accent = No Style, No Grid; themed2 without accent = No Style, Table Grid.
const BUILTIN: Record<string, { family: BuiltinFamily; accent?: string }> = {
  '{2D5ABB26-0587-4C30-8999-92F81FD0307C}': { family: 'themed1' },
  '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}': { family: 'themed1', accent: 'accent1' },
  '{284E427A-3D55-4303-BF80-6455036E1DE7}': { family: 'themed1', accent: 'accent2' },
  '{69C7853C-536D-4A76-A0AE-DD22124D55A5}': { family: 'themed1', accent: 'accent3' },
  '{775DCB02-9BB8-47FD-8907-85C794F793BA}': { family: 'themed1', accent: 'accent4' },
  '{35758FB7-9AC5-4552-8A53-C91805E547FA}': { family: 'themed1', accent: 'accent5' },
  '{08FB837D-C827-4EFA-A057-4D05807E0F7C}': { family: 'themed1', accent: 'accent6' },
  '{5940675A-B579-460E-94D1-54222C63F5DA}': { family: 'themed2' },
  '{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}': { family: 'themed2', accent: 'accent1' },
  '{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}': { family: 'themed2', accent: 'accent2' },
  '{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}': { family: 'themed2', accent: 'accent3' },
  '{E269D01E-BC32-4049-B463-5C60D7B0CCD2}': { family: 'themed2', accent: 'accent4' },
  '{327F97BB-C833-4FB7-BDE5-3F7075034690}': { family: 'themed2', accent: 'accent5' },
  '{638B1855-1B75-4FBE-930C-398BA8C253C6}': { family: 'themed2', accent: 'accent6' },
  '{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}': { family: 'light1' },
  '{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}': { family: 'light1', accent: 'accent1' },
  '{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}': { family: 'light1', accent: 'accent2' },
  '{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}': { family: 'light1', accent: 'accent3' },
  '{D27102A9-8310-4765-A935-A1911B00CA55}': { family: 'light1', accent: 'accent4' },
  '{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}': { family: 'light1', accent: 'accent5' },
  '{68D230F3-CF80-4859-8CE7-A43EE81993B5}': { family: 'light1', accent: 'accent6' },
  '{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}': { family: 'light2' },
  '{69012ECD-51FC-41F1-AA8D-1B2483CD663E}': { family: 'light2', accent: 'accent1' },
  '{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}': { family: 'light2', accent: 'accent2' },
  '{F2DE63D5-997A-4646-A377-4702673A728D}': { family: 'light2', accent: 'accent3' },
  '{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}': { family: 'light2', accent: 'accent4' },
  '{5A111915-BE36-4E01-A7E5-04B1672EAD32}': { family: 'light2', accent: 'accent5' },
  '{912C8C85-51F0-491E-9774-3900AFEF0FD7}': { family: 'light2', accent: 'accent6' },
  '{616DA210-FB5B-4158-B5E0-FEB733F419BA}': { family: 'light3' },
  '{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}': { family: 'light3', accent: 'accent1' },
  '{5DA37D80-6434-44D0-A028-1B22A696006F}': { family: 'light3', accent: 'accent2' },
  '{8799B23B-EC83-4686-B30A-512413B5E67A}': { family: 'light3', accent: 'accent3' },
  '{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}': { family: 'light3', accent: 'accent4' },
  '{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}': { family: 'light3', accent: 'accent5' },
  '{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}': { family: 'light3', accent: 'accent6' },
  '{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}': { family: 'medium1' },
  '{B301B821-A1FF-4177-AEE7-76D212191A09}': { family: 'medium1', accent: 'accent1' },
  '{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}': { family: 'medium1', accent: 'accent2' },
  '{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}': { family: 'medium1', accent: 'accent3' },
  '{1E171933-4619-4E11-9A3F-F7608DF75F80}': { family: 'medium1', accent: 'accent4' },
  '{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}': { family: 'medium1', accent: 'accent5' },
  '{10A1B5D5-9B99-4C35-A422-299274C87663}': { family: 'medium1', accent: 'accent6' },
  '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}': { family: 'medium2' },
  '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}': { family: 'medium2', accent: 'accent1' },
  '{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}': { family: 'medium2', accent: 'accent2' },
  '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}': { family: 'medium2', accent: 'accent3' },
  '{00A15C55-8517-42AA-B614-E9B94910E393}': { family: 'medium2', accent: 'accent4' },
  '{7DF18680-E054-41AD-8BC1-D1AEF772440D}': { family: 'medium2', accent: 'accent5' },
  '{93296810-A885-4BE3-A3E7-6D5BEEA58F35}': { family: 'medium2', accent: 'accent6' },
  '{8EC20E35-A176-4012-BC5E-935CFFF8708E}': { family: 'medium3' },
  '{6E25E649-3F16-4E02-A733-19D2CDBF48F0}': { family: 'medium3', accent: 'accent1' },
  '{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}': { family: 'medium3', accent: 'accent2' },
  '{EB344D84-9AFB-497E-A393-DC336BA19D2E}': { family: 'medium3', accent: 'accent3' },
  '{EB9631B5-78F2-41C9-869B-9F39066F8104}': { family: 'medium3', accent: 'accent4' },
  '{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}': { family: 'medium3', accent: 'accent5' },
  '{2A488322-F2BA-4B5B-9748-0D474271808F}': { family: 'medium3', accent: 'accent6' },
  '{D7AC3CCA-C797-4891-BE02-D94E43425B78}': { family: 'medium4' },
  '{69CF1AB2-1976-4502-BF36-3FF5EA218861}': { family: 'medium4', accent: 'accent1' },
  '{8A107856-5554-42FB-B03E-39F5DBC370BA}': { family: 'medium4', accent: 'accent2' },
  '{0505E3EF-67EA-436B-97B2-0124C06EBD24}': { family: 'medium4', accent: 'accent3' },
  '{C4B1156A-380E-4F78-BDF5-A606A8083BF9}': { family: 'medium4', accent: 'accent4' },
  '{22838BEF-8BB2-4498-84A7-C5851F593DF1}': { family: 'medium4', accent: 'accent5' },
  '{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}': { family: 'medium4', accent: 'accent6' },
  '{E8034E78-7F5D-4C2E-B375-FC64B27BC917}': { family: 'dark1' },
  '{125E5076-3810-47DD-B79F-674D7AD40C01}': { family: 'dark1', accent: 'accent1' },
  '{37CE84F3-28C3-443E-9E96-99CF82512B78}': { family: 'dark1', accent: 'accent2' },
  '{D03447BB-5D67-496B-8E87-E561075AD55C}': { family: 'dark1', accent: 'accent3' },
  '{E929F9F4-4A8F-4326-A1B4-22849713DDAB}': { family: 'dark1', accent: 'accent4' },
  '{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}': { family: 'dark1', accent: 'accent5' },
  '{AF606853-7671-496A-8E4F-DF71F8EC918B}': { family: 'dark1', accent: 'accent6' },
  '{5202B0CA-FC54-4496-8BCA-5EF66A818D29}': { family: 'dark2' },
  '{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}': { family: 'dark2', accent: 'accent1' },
  '{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}': { family: 'dark2', accent: 'accent3' },
  '{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}': { family: 'dark2', accent: 'accent5' },
}

/** Compatibility: the made-up "no style" GUID this app used to write (not in the official table). */
export const LEGACY_NO_STYLE = '{2D5ABB26-0587-4C30-8999-92F81FD0307D}'

/** tint: keep pct of the color and blend the rest toward white (OOXML <a:tint val="20000"> = 20%). */
function tint(hex: string, pct: number): string {
  return mixToward(hex, pct, 255)
}

/** shade: keep pct of the color and blend the rest toward black (OOXML <a:shade>). */
function shade(hex: string, pct: number): string {
  return mixToward(hex, pct, 0)
}

function mixToward(hex: string, pct: number, target: number): string {
  const h = hex.replace('#', '')
  const ch = (i: number) => parseInt(h.slice(i, i + 2), 16)
  const mix = (c: number) => Math.round(target * (1 - pct) + c * pct)
  const to = (c: number) => c.toString(16).toUpperCase().padStart(2, '0')
  return `#${to(mix(ch(0)))}${to(mix(ch(2)))}${to(mix(ch(4)))}`
}

/** Pre-blended solid color of fg composited over bg at alpha opacity (banded alpha fill -> solid color). */
function over(fg: string, alpha: number, bg: string): string {
  const f = fg.replace('#', '')
  const b = bg.replace('#', '')
  const ch = (s: string, i: number) => parseInt(s.slice(i, i + 2), 16)
  const mix = (i: number) => Math.round(ch(f, i) * alpha + ch(b, i) * (1 - alpha))
  const to = (c: number) => c.toString(16).toUpperCase().padStart(2, '0')
  return `#${to(mix(0))}${to(mix(2))}${to(mix(4))}`
}

const solid = (color: string): Fill => ({ type: 'solid', color })
const line = (color: string, width = 12700): Stroke => ({ fill: solid(color), width })

/**
 * Built-in style family -> region styles (color rules follow the LibreOffice
 * predefined-table-styles implementation; colors are resolved to final values via the
 * theme; banded alpha is pre-blended against the underlying background).
 */
function builtinStyle(
  family: BuiltinFamily,
  accentName: string | undefined,
  theme: Theme | undefined,
): TableStyleDef {
  const lt1 = resolveSchemeColor('lt1', theme) ?? '#FFFFFF'
  const dk1 = resolveSchemeColor('dk1', theme) ?? '#000000'
  const accent = accentName ? (resolveSchemeColor(accentName, theme) ?? '#4472C4') : undefined

  switch (family) {
    case 'themed1': {
      if (!accent) return { wholeTbl: { textColor: dk1 } } // No Style, No Grid
      const band = solid(over(accent, 0.4, lt1))
      return {
        wholeTbl: { textColor: dk1 },
        firstRow: { fill: solid(accent), textColor: lt1 },
        band1H: { fill: band },
        band1V: { fill: band },
        insideH: line(accent),
        insideV: line(accent),
        outer: { l: line(accent), r: line(accent), t: line(accent), b: line(accent) },
        firstRowBottom: line(lt1),
      }
    }
    case 'themed2': {
      if (!accent) {
        // No Style, Table Grid: all-dk1 grid
        return {
          wholeTbl: { textColor: dk1 },
          insideH: line(dk1),
          insideV: line(dk1),
          outer: { l: line(dk1), r: line(dk1), t: line(dk1), b: line(dk1) },
        }
      }
      const outer = line(tint(accent, 0.5))
      const band = solid(over(lt1, 0.2, accent))
      return {
        wholeTbl: { fill: solid(accent), textColor: lt1 },
        band1H: { fill: band },
        band1V: { fill: band },
        outer: { l: outer, r: outer, t: outer, b: outer },
        firstRowBottom: line(lt1),
        lastRowTop: line(lt1),
      }
    }
    case 'light1': {
      const a = accent ?? dk1
      const band = solid(over(a, 0.2, lt1))
      return {
        wholeTbl: { textColor: dk1 },
        firstRow: { bold: true, textColor: dk1 },
        firstCol: { bold: true },
        lastCol: { bold: true, textColor: dk1 },
        band1H: { fill: band },
        band1V: { fill: band },
        outer: { t: line(a), b: line(a) },
        firstRowBottom: line(a),
        lastRowTop: line(a),
      }
    }
    case 'light2': {
      const a = accent ?? dk1
      return {
        wholeTbl: { textColor: dk1 },
        firstRow: { fill: solid(a), textColor: lt1, bold: true },
        outer: { l: line(a), r: line(a), t: line(a), b: line(a) },
        lastRowTop: line(a),
      }
    }
    case 'light3': {
      const a = accent ?? dk1
      const band = solid(over(a, 0.2, lt1))
      return {
        wholeTbl: { textColor: dk1 },
        firstRow: { textColor: a, bold: true },
        band1H: { fill: band },
        band1V: { fill: band },
        insideH: line(a),
        insideV: line(a),
        outer: { l: line(a), r: line(a), t: line(a), b: line(a) },
        firstRowBottom: line(a),
        lastRowTop: line(a),
      }
    }
    case 'medium1': {
      const a = accent ?? dk1
      const band = solid(tint(a, 0.2))
      return {
        wholeTbl: { fill: solid(lt1), textColor: dk1 },
        firstRow: { fill: solid(a), textColor: lt1, bold: true },
        lastRow: { fill: solid(lt1), bold: true },
        band1H: { fill: band },
        band1V: { fill: band },
        insideH: line(a),
        outer: { l: line(a), r: line(a), t: line(a), b: line(a) },
        lastRowTop: line(a),
      }
    }
    case 'medium2': {
      const a = accent ?? dk1
      return {
        wholeTbl: { fill: solid(tint(a, 0.2)), textColor: dk1 },
        band1H: { fill: solid(tint(a, 0.4)) },
        band1V: { fill: solid(tint(a, 0.4)) },
        firstRow: { fill: solid(a), textColor: lt1, bold: true },
        lastRow: { fill: solid(a), textColor: lt1, bold: true },
        firstCol: { fill: solid(a), textColor: lt1, bold: true },
        lastCol: { fill: solid(a), textColor: lt1, bold: true },
        insideH: line(lt1),
        insideV: line(lt1),
        outer: { l: line(lt1), r: line(lt1), t: line(lt1), b: line(lt1) },
        firstRowBottom: line(lt1),
        lastRowTop: line(lt1),
      }
    }
    case 'medium3': {
      const a = accent ?? dk1
      const band = solid(tint(dk1, 0.2))
      return {
        wholeTbl: { fill: solid(lt1), textColor: dk1 },
        firstRow: { fill: solid(a), textColor: lt1, bold: true },
        lastRow: { fill: solid(lt1), bold: true },
        firstCol: { fill: solid(a), textColor: lt1 },
        lastCol: { fill: solid(a), textColor: lt1 },
        band1H: { fill: band },
        band1V: { fill: band },
        outer: { t: line(dk1), b: line(dk1) },
        firstRowBottom: line(dk1),
        lastRowTop: line(dk1),
      }
    }
    case 'medium4': {
      const a = accent ?? dk1
      const band = solid(tint(a, 0.4))
      return {
        wholeTbl: { fill: solid(tint(a, 0.2)), textColor: dk1 },
        firstRow: { fill: solid(tint(a, 0.2)), textColor: a, bold: true },
        lastRow: { fill: solid(tint(dk1, 0.2)) },
        band1H: { fill: band },
        band1V: { fill: band },
        insideH: line(a),
        insideV: line(a),
        outer: { l: line(a), r: line(a), t: line(a), b: line(a) },
        lastRowTop: line(dk1),
      }
    }
    case 'dark1': {
      // accent variants use shade (dark background, light text); the no-accent variant uses dk1 tint (light gray background, dark text)
      const a = accent ?? dk1
      const tf = accent ? shade : tint
      const body = accent ? lt1 : dk1
      return {
        wholeTbl: { fill: solid(tf(a, 0.2)), textColor: body },
        firstRow: { fill: solid(dk1), textColor: lt1, bold: true },
        lastRow: { fill: solid(tf(a, 0.2)), textColor: body, bold: true },
        firstCol: { fill: solid(tf(a, 0.6)) },
        lastCol: { fill: solid(tf(a, 0.6)) },
        band1H: { fill: solid(tf(a, 0.4)) },
        band1V: { fill: solid(tf(a, 0.4)) },
        firstRowBottom: line(lt1),
        lastRowTop: line(lt1),
      }
    }
    case 'dark2': {
      const a = accent ?? dk1
      // Header paired colors: base=dk1, Accent1->accent2 / Accent3->accent4 / Accent5->accent6
      const headerPair: Record<string, string> = {
        accent1: 'accent2',
        accent3: 'accent4',
        accent5: 'accent6',
      }
      const header = accentName
        ? (resolveSchemeColor(headerPair[accentName] ?? accentName, theme) ?? dk1)
        : dk1
      const band = solid(tint(a, 0.4))
      return {
        wholeTbl: { fill: solid(tint(a, 0.2)), textColor: dk1 },
        firstRow: { fill: solid(header), textColor: lt1, bold: true },
        lastRow: { fill: solid(tint(a, 0.2)), bold: true },
        band1H: { fill: band },
        band1V: { fill: band },
        lastRowTop: line(dk1),
      }
    }
  }
}

/**
 * styleId → style definition. Prefers explicit definitions in tableStyles.xml, falling
 * back to the built-in table.
 */
export function resolveTableStyle(
  styleId: string | undefined,
  tableStylesXml: string | undefined,
  theme: Theme | undefined,
): TableStyleDef | undefined {
  if (!styleId) return undefined
  if (tableStylesXml && tableStylesXml.includes(styleId)) {
    const def = parseTableStylesXml(tableStylesXml, styleId, theme)
    if (def) return def
  }
  const builtin = BUILTIN[styleId]
  if (builtin) return builtinStyle(builtin.family, builtin.accent, theme)
  if (styleId === LEGACY_NO_STYLE) return {}
  return undefined
}

/**
 * Style borders for cell (r,c) (composed from region borders; explicit tcPr borders
 * take precedence): inner separator lines + whole-table outer border + header
 * separator lines (when the firstRow/lastRow flags are on).
 */
export function cellStyleBorders(
  def: TableStyleDef,
  flags: TableStyleFlags,
  r: number,
  c: number,
  nRows: number,
  nCols: number,
): { l?: Stroke; r?: Stroke; t?: Stroke; b?: Stroke } {
  const out: { l?: Stroke; r?: Stroke; t?: Stroke; b?: Stroke } = {}
  if (r < nRows - 1 && def.insideH) out.b = def.insideH
  if (c < nCols - 1 && def.insideV) out.r = def.insideV
  if (def.outer) {
    if (r === 0 && def.outer.t) out.t = def.outer.t
    if (r === nRows - 1 && def.outer.b) out.b = def.outer.b
    if (c === 0 && def.outer.l) out.l = def.outer.l
    if (c === nCols - 1 && def.outer.r) out.r = def.outer.r
  }
  if (flags.firstRow && r === 0 && nRows > 1 && def.firstRowBottom) out.b = def.firstRowBottom
  if (flags.lastRow && r === nRows - 1 && nRows > 1 && def.lastRowTop) out.t = def.lastRowTop
  return out
}

// ── tableStyles.xml parsing ─────────────────────────────────────────

const tsParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  isArray: (name) => name === 'a:tblStyle',
})

function readColor(node: unknown, theme: Theme | undefined): string | undefined {
  if (!node || typeof node !== 'object') return undefined
  const n = asXmlNode(node)
  // tint/shade/alpha child modifiers; alpha lands as an #RRGGBBAA suffix so
  // banded fills composite over <a:tblBg> instead of collapsing to opaque
  const withMods = (base: string, clr: XmlNode): string => {
    let c = base
    const t = asXmlNode(clr['a:tint'])['@_val']
    if (t) c = tint(c, parseInt(String(t), 10) / 100000)
    const sh = asXmlNode(clr['a:shade'])['@_val']
    if (sh) c = shade(c, parseInt(String(sh), 10) / 100000)
    const al = asXmlNode(clr['a:alpha'])['@_val']
    if (al) {
      const a = Math.max(0, Math.min(255, Math.round((parseInt(String(al), 10) / 100000) * 255)))
      c = c + a.toString(16).padStart(2, '0').toUpperCase()
    }
    return c
  }
  if (n['a:srgbClr']) {
    const srgb = asXmlNode(n['a:srgbClr'])
    return withMods('#' + String(srgb['@_val']).toUpperCase(), srgb)
  }
  if (n['a:schemeClr']) {
    const scheme = asXmlNode(n['a:schemeClr'])
    const base = resolveSchemeColor(String(scheme['@_val']), theme)
    if (!base) return undefined
    return withMods(base, scheme)
  }
  const prst = asXmlNode(n['a:prstClr'])['@_val']
  if (prst === 'black') return '#000000'
  if (prst === 'white') return '#FFFFFF'
  return undefined
}

function readPart(part: unknown, theme: Theme | undefined): TablePartStyle | undefined {
  if (!part || typeof part !== 'object') return undefined
  const p = asXmlNode(part)
  const out: TablePartStyle = {}
  const fillColor = readColor(asXmlNode(asXmlNode(p['a:tcStyle'])['a:fill'])['a:solidFill'], theme)
  if (fillColor) out.fill = solid(fillColor)
  if (p['a:tcTxStyle']) {
    const tx = asXmlNode(p['a:tcTxStyle'])
    if (tx['@_b'] === 'on') out.bold = true
    const c = readColor(tx, theme)
    if (c) out.textColor = c
  }
  return Object.keys(out).length ? out : undefined
}

function readInside(
  part: unknown,
  tag: 'a:insideH' | 'a:insideV' | 'a:left' | 'a:right' | 'a:top' | 'a:bottom',
  theme: Theme | undefined,
): Stroke | undefined {
  const bdr = asXmlNode(asXmlNode(asXmlNode(part)['a:tcStyle'])['a:tcBdr'])
  const edge = asXmlNode(bdr[tag])
  const lnRaw = edge['a:ln']
  if (lnRaw) {
    const ln = asXmlNode(lnRaw)
    if ('a:noFill' in ln) return undefined
    const c = readColor(ln['a:solidFill'], theme)
    if (!c) return undefined
    return line(c, parseInt(String(ln['@_w']), 10) || 12700)
  }
  // <a:lnRef idx>: theme lnStyleLst template; the ref's child color substitutes phClr
  const refRaw = edge['a:lnRef']
  if (refRaw) {
    const ref = asXmlNode(refRaw)
    const c = readColor(ref, theme)
    if (!c) return undefined
    const idx = parseInt(String(ref['@_idx'] ?? '0'), 10) || 0
    const tplLn = asXmlNode(asXmlNode(theme?.lnStyles?.[idx - 1])['a:ln'])
    return line(c, parseInt(String(tplLn['@_w']), 10) || 12700)
  }
  return undefined
}

function parseTableStylesXml(
  xml: string,
  styleId: string,
  theme: Theme | undefined,
): TableStyleDef | undefined {
  let doc: XmlNode
  try {
    doc = asXmlNode(tsParser.parse(xml))
  } catch {
    return undefined
  }
  const list = asXmlNode(doc['a:tblStyleLst'])['a:tblStyle'] ?? []
  const style = xmlArray(list).find((s) => s['@_styleId'] === styleId)
  if (!style) return undefined
  const def: TableStyleDef = {}
  const tblBgRaw = style['a:tblBg']
  if (tblBgRaw) {
    const bg = asXmlNode(tblBgRaw)
    const direct = readColor(asXmlNode(bg['a:fill'])['a:solidFill'], theme)
    if (direct) def.tblBg = solid(direct)
    const refRaw = bg['a:fillRef']
    if (refRaw) {
      const ref = asXmlNode(refRaw)
      const idx = parseInt(String(ref['@_idx'] ?? '0'), 10) || 0
      if (idx > 0) def.tblBgRef = { idx, phClr: readColor(ref, theme) }
    }
  }
  for (const key of [
    'wholeTbl',
    'band1H',
    'band2H',
    'band1V',
    'band2V',
    'firstRow',
    'lastRow',
    'firstCol',
    'lastCol',
  ] as const) {
    const p = readPart(style['a:' + key], theme)
    if (p) def[key] = p
  }
  def.insideH = readInside(style['a:wholeTbl'], 'a:insideH', theme)
  def.insideV = readInside(style['a:wholeTbl'], 'a:insideV', theme)
  if (!def.insideH) delete def.insideH
  if (!def.insideV) delete def.insideV
  const outer = {
    l: readInside(style['a:wholeTbl'], 'a:left', theme),
    r: readInside(style['a:wholeTbl'], 'a:right', theme),
    t: readInside(style['a:wholeTbl'], 'a:top', theme),
    b: readInside(style['a:wholeTbl'], 'a:bottom', theme),
  }
  if (outer.l || outer.r || outer.t || outer.b) def.outer = outer
  return Object.keys(def).length ? def : undefined
}

/**
 * Compute the final region style for cell (r,c) (OOXML precedence:
 * wholeTbl < banding < lastRow/firstCol/lastCol < firstRow).
 */
export function cellPartStyle(
  def: TableStyleDef,
  flags: TableStyleFlags,
  r: number,
  c: number,
  nRows: number,
  nCols: number,
): TablePartStyle {
  const merge = (base: TablePartStyle, over?: TablePartStyle): TablePartStyle =>
    over
      ? { ...base, ...Object.fromEntries(Object.entries(over).filter(([, v]) => v !== undefined)) }
      : base

  let out: TablePartStyle = { ...(def.wholeTbl ?? {}) }
  const isFirstRow = flags.firstRow && r === 0
  const isLastRow = flags.lastRow && r === nRows - 1
  const isFirstCol = flags.firstCol && c === 0
  const isLastCol = flags.lastCol && c === nCols - 1

  if (flags.bandRow && !isFirstRow && !isLastRow) {
    const dataR = r - (flags.firstRow ? 1 : 0)
    out = merge(out, dataR % 2 === 0 ? def.band1H : def.band2H)
  }
  if (flags.bandCol && !isFirstCol && !isLastCol) {
    const dataC = c - (flags.firstCol ? 1 : 0)
    out = merge(out, dataC % 2 === 0 ? def.band1V : def.band2V)
  }
  if (isLastCol) out = merge(out, def.lastCol)
  if (isFirstCol) out = merge(out, def.firstCol)
  if (isLastRow) out = merge(out, def.lastRow)
  if (isFirstRow) out = merge(out, def.firstRow)
  return out
}
