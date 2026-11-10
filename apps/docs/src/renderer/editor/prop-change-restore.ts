/**
 * Rejecting a tracked table / row / cell property change puts the recorded
 * previous properties back in the saved file (the engine restores the raw
 * bytes). The editor restores the modeled subset here so the page shows the
 * same result without a reload; unmodeled children (borders, exceptions) only
 * catch up on reopen.
 */

const attr = (xml: string, tag: string, name: string): string | undefined =>
  new RegExp(`<${tag}\\b[^>]*\\s${name}="([^"]*)"`).exec(xml)?.[1]
const has = (xml: string, tag: string): boolean => new RegExp(`<${tag}[\\s/>]`).test(xml)
const flagOn = (xml: string, tag: string): boolean => {
  const m = new RegExp(`<${tag}\\b([^>]*)/?>`).exec(xml)
  if (!m) return false
  const val = /\sw:val="([^"]*)"/.exec(m[1])?.[1]
  return val === undefined || !/^(?:0|false|off)$/.test(val)
}

const TBL_ALIGN: Record<string, 'left' | 'center' | 'right'> = {
  left: 'left',
  start: 'left',
  center: 'center',
  right: 'right',
  end: 'right',
}

function cellMarginsOf(xml: string, holder: string): Record<string, number> | null {
  const inner = new RegExp(`<${holder}\\b[^>]*>([\\s\\S]*?)</${holder}>`).exec(xml)?.[1]
  if (inner === undefined) return null
  const out: Record<string, number> = {}
  for (const side of ['top', 'left', 'bottom', 'right', 'start', 'end'] as const) {
    const w = Number(attr(inner, `w:${side}`, 'w:w'))
    if (!Number.isFinite(w)) continue
    out[side === 'start' ? 'left' : side === 'end' ? 'right' : side] = w
  }
  return out
}

/** table node attrs restored from the previous tblPr children */
export function tblPrRestoreAttrs(oldXml: string): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  const jc = attr(oldXml, 'w:jc', 'w:val')
  out.tblAlign = (jc && TBL_ALIGN[jc]) || null
  out.tblStyleId = attr(oldXml, 'w:tblStyle', 'w:val') ?? null
  const ind = Number(attr(oldXml, 'w:tblInd', 'w:w'))
  out.indentTwips = Number.isFinite(ind) && ind !== 0 ? ind : null
  const fill = attr(oldXml, 'w:shd', 'w:fill')
  out.tblFill = fill && fill !== 'auto' ? fill.toUpperCase() : null
  out.cellMar = cellMarginsOf(oldXml, 'w:tblCellMar')
  out.cellMarEdited = true
  if (has(oldXml, 'w:tblLook')) {
    const look = /<w:tblLook\b([^>]*)\/?>/.exec(oldXml)![1]
    const val = parseInt(/\sw:val="([0-9A-Fa-f]+)"/.exec(look)?.[1] ?? '', 16)
    const flag = (name: string, bit: number) => {
      const explicit = /\s/.test(look) && new RegExp(`\\sw:${name}="([^"]*)"`).exec(look)?.[1]
      if (explicit !== undefined && explicit !== false)
        return explicit === '1' || explicit === 'true'
      return Number.isFinite(val) ? (val & bit) !== 0 : false
    }
    out.tblLook = {
      firstRow: flag('firstRow', 0x20),
      lastRow: flag('lastRow', 0x40),
      firstColumn: flag('firstColumn', 0x80),
      lastColumn: flag('lastColumn', 0x100),
      bandedRows: !flag('noHBand', 0x200),
      bandedColumns: !flag('noVBand', 0x400),
    }
  } else out.tblLook = null
  out.tblLookEdited = true
  if (has(oldXml, 'w:tblW') || has(oldXml, 'w:tblLayout')) {
    const type = attr(oldXml, 'w:tblW', 'w:type')
    const w = Number(attr(oldXml, 'w:tblW', 'w:w'))
    out.tblAutoFit =
      attr(oldXml, 'w:tblLayout', 'w:type') === 'fixed'
        ? 'fixed'
        : type === 'pct' && w >= 5000
          ? 'window'
          : type === 'auto' || !(w > 0)
            ? 'contents'
            : 'fixed'
    out.tblAutoFitEdited = true
  }
  out.tblCaption = attr(oldXml, 'w:tblCaption', 'w:val') ?? null
  out.tblDescription = attr(oldXml, 'w:tblDescription', 'w:val') ?? null
  out.tblAltEdited = true
  return out
}

/** row node attrs restored from the previous trPr children */
export function trPrRestoreAttrs(oldXml: string): Record<string, unknown> {
  const h = Number(attr(oldXml, 'w:trHeight', 'w:val'))
  return {
    heightTwips: Number.isFinite(h) && h > 0 ? Math.min(h, 31680) : null,
    heightRule:
      Number.isFinite(h) && h > 0
        ? attr(oldXml, 'w:trHeight', 'w:hRule') === 'exact'
          ? 'exact'
          : 'atLeast'
        : null,
    repeatHeader: flagOn(oldXml, 'w:tblHeader'),
    repeatHeaderEdited: true,
    cantSplit: flagOn(oldXml, 'w:cantSplit'),
    cantSplitEdited: true,
  }
}

/** cell node attrs restored from the previous tcPr children */
export function tcPrRestoreAttrs(oldXml: string): Record<string, unknown> {
  const fill = attr(oldXml, 'w:shd', 'w:fill')
  const vAlign = attr(oldXml, 'w:vAlign', 'w:val')
  const dir = attr(oldXml, 'w:textDirection', 'w:val')
  return {
    fill: fill && fill !== 'auto' ? fill.toUpperCase() : null,
    vAlign: vAlign === 'center' || vAlign === 'bottom' ? vAlign : null,
    textDirection: dir?.startsWith('tbRl') ? 'tbRl' : dir?.startsWith('btLr') ? 'btLr' : null,
    cellMar: cellMarginsOf(oldXml, 'w:tcMar'),
    noWrap: flagOn(oldXml, 'w:noWrap'),
    tcPrEdited: true,
  }
}

/** gridCol widths (twips) of a previous tblGrid */
export function gridColTwips(oldXml: string): number[] {
  return [...oldXml.matchAll(/<w:gridCol\b[^>]*\sw:w="(\d+)"/g)].map((m) => Number(m[1]))
}
