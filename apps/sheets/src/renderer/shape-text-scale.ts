import type { CSSProperties } from 'react'

const PT_TO_PX = 4 / 3

/// Shape text is authored in logical px (zoom 1) while the float box the
/// shape lives in is `frame × zoom` wide, so one zoomed px is 100cqw over
/// the logical frame width. Without a known frame the text stays logical.
export function shapeZoomedPx(frameWidth: number | undefined): string {
  return frameWidth && frameWidth > 0 ? `calc(100cqw / ${frameWidth})` : '1px'
}

export function shapeTextScaleStyle(frameWidth: number | undefined): CSSProperties {
  return { '--shape-px': shapeZoomedPx(frameWidth) } as CSSProperties
}

export function shapeRunFontSize(sizePt: number): string {
  return `calc(var(--shape-px, 1px) * ${Number((sizePt * PT_TO_PX).toFixed(3))})`
}

/// bodyPr vertOverflow / horzOverflow: `clip` cuts the text at the frame
/// minus insets; `ellipsis` has no CSS twin in a fixed frame, so it clips
/// too. The ECMA default `overflow` keeps the frame-edge clipping.
export function shapeTextOverflowClass(
  vertOverflow: string | undefined,
  horzOverflow: string | undefined,
): string {
  const clips = (value: string | undefined) => value === 'clip' || value === 'ellipsis'
  return clips(vertOverflow) || clips(horzOverflow) ? ' shape-text-clip' : ''
}
