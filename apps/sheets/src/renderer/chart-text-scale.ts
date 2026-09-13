import type { CSSProperties } from 'react'

/// Rendered size of a chart SVG (CSS px) and the sheet zoom it renders at.
export interface ChartTextBox {
  readonly width: number
  readonly height: number
  readonly zoom: number
}

/// Excel's defaults: 9pt tick labels, 10pt axis titles.
export const AXIS_LABEL_PT = 9
export const AXIS_TITLE_PT = 10

/// Average Calibri glyph advance as a fraction of the font size.
export const CHAR_EM = 0.54

const PT_TO_PX = 4 / 3

/// User units of one point for text inside an SVG drawn with a `vbW×vbH`
/// viewBox (xMidYMid meet). The viewBox scales with the frame while Excel
/// keeps chart text at its point size, so text must be counter-scaled:
/// unit = pt→px · zoom / viewBox scale. Unmeasured boxes keep 1 unit per
/// point, which is the historic look.
export function chartTextUnit(
  box: ChartTextBox | undefined,
  viewBoxWidth: number,
  viewBoxHeight: number,
): number {
  if (!box || !(box.width > 0) || !(box.height > 0)) return 1
  const scale = Math.min(box.width / viewBoxWidth, box.height / viewBoxHeight)
  return round((PT_TO_PX * (box.zoom || 1)) / scale)
}

/// Text unit plus the left viewBox extension that keeps value tick labels
/// (widest `labelEm` em at `pt`, plus `reservePt` of axis title / unit label
/// columns) inside the picture when their 48-unit gutter is too narrow (the
/// labels end at x=54). Widening the viewBox lowers the scale,
/// which enlarges the unit again; the width-bound case solves that fixed
/// point in closed form: unit = a·(base + extra), extra = c·unit − available.
export function valueAxisLayout(
  box: ChartTextBox | undefined,
  viewBoxHeight: number,
  labelEm: number,
  pt: number,
  reservePt = 0,
): { unit: number; extra: number } {
  const base = 600
  const available = 48
  const c = labelEm * pt + reservePt
  const extraFor = (unit: number): number => Math.min(300, Math.max(0, round(c * unit - available)))
  if (!box || !(box.width > 0) || !(box.height > 0)) return { unit: 1, extra: extraFor(1) }
  const zoomedPt = PT_TO_PX * (box.zoom || 1)
  const unitH = zoomedPt / (box.height / viewBoxHeight)
  const extraH = extraFor(unitH)
  if (box.width / (base + extraH) >= box.height / viewBoxHeight) {
    return { unit: round(unitH), extra: extraH }
  }
  const a = zoomedPt / box.width
  let unit = a * base
  if (c * unit > available && a * c < 1) unit = (a * (base - available)) / (1 - a * c)
  return { unit: round(unit), extra: extraFor(unit) }
}

export interface ChartAxisText {
  readonly labelSize?: number | undefined
  readonly labelColor?: string | undefined
  readonly titleSize?: number | undefined
  readonly titleColor?: string | undefined
}

/// Points of gutter the rotated axis title and unit label columns take.
export function axisSideReservePt(
  title: string | null | undefined,
  units: string | undefined,
  titlePt: number,
  labelPt: number,
): number {
  return (title ? 1.15 * titlePt : 0) + (units ? 1.1 * labelPt : 0)
}

/// CSS custom properties the chart stylesheet sizes SVG text from.
export function chartTextStyle(
  unit: number,
  category: ChartAxisText | undefined,
  value: ChartAxisText | undefined,
): CSSProperties {
  const style: Record<string, string | number> = { '--chart-pt': `${unit}px` }
  if (category?.labelSize) style['--chart-cat-pt'] = category.labelSize
  if (category?.labelColor) style['--chart-cat-color'] = category.labelColor
  if (value?.labelSize) style['--chart-val-pt'] = value.labelSize
  if (value?.labelColor) style['--chart-val-color'] = value.labelColor
  const titlePt = value?.titleSize ?? category?.titleSize
  if (titlePt) style['--chart-title-pt'] = titlePt
  const titleColor = value?.titleColor ?? category?.titleColor
  if (titleColor) style['--chart-title-color'] = titleColor
  return style as CSSProperties
}

/// Ink for chart text over the chart-area fill: the file's own text color
/// when it has one, otherwise light ink on a dark area (WCAG relative
/// luminance) and the theme default (undefined) on a light one.
export function chartAreaInk(
  fill: string | undefined,
  explicit: string | undefined,
): string | undefined {
  if (explicit) return explicit
  if (!fill) return undefined
  const luminance = relativeLuminance(fill)
  return luminance !== undefined && luminance < 0.35 ? '#F2F2F2' : undefined
}

function relativeLuminance(color: string): number | undefined {
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(color.trim())?.[1]
  if (!hex) return undefined
  const digits =
    hex.length === 3 ? [...hex].map((d) => d + d) : [hex.slice(0, 2), hex.slice(2, 4), hex.slice(4)]
  const [r, g, b] = digits.map((d) => {
    const channel = Number.parseInt(d, 16) / 255
    return channel <= 0.03928 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4
  }) as [number, number, number]
  return 0.2126 * r + 0.7152 * g + 0.0722 * b
}

function round(value: number): number {
  return Number(value.toFixed(3))
}

/// Bottom band under the plot: tick line(s) from y=292/298, then the axis
/// title at 317 (333 past a group band). Two point-sized tick lines can
/// run into the title; the title then moves down and the viewBox grows.
export function bottomAxisLayout(
  wrapped: boolean,
  hasTitle: boolean,
  groupBand: boolean,
  tickFont: number,
  titleFont: number,
): { bottomY: number; viewBoxHeight: number } {
  const bottomY = groupBand ? 333 : 317
  const viewBoxHeight = groupBand ? 336 : 320
  if (!wrapped || !hasTitle) return { bottomY, viewBoxHeight }
  const secondLine = 292 + 1.1 * tickFont
  const shifted = Math.max(bottomY, round(secondLine + 0.3 * tickFont + 0.85 * titleFont))
  if (shifted <= bottomY) return { bottomY, viewBoxHeight }
  return { bottomY: shifted, viewBoxHeight: Math.ceil(shifted + 4) }
}
