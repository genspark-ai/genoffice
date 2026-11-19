/**
 * DrawingML picture transform (pic a:xfrm rot / flipH / flipV) as CSS.
 * Flips mirror the source and rot turns the result; CSS applies functions
 * right-to-left, so the scale sits last.
 */
export function pictureTransformFns(
  rotDeg: number | null | undefined,
  flipH?: boolean | null,
  flipV?: boolean | null,
): string[] {
  const fns: string[] = []
  if (rotDeg) fns.push(`rotate(${Number(rotDeg)}deg)`)
  if (flipH) fns.push('scaleX(-1)')
  if (flipV) fns.push('scaleY(-1)')
  return fns
}

/**
 * Word turns a picture about the centre of its unrotated wp:extent and lets
 * the flow reserve the rotated bounding box. For a quarter turn that box is
 * the extent with its sides swapped: the visual sticks out (w-h)/2 above and
 * below the extent box and tucks in by the same amount left and right.
 * Returns that half-difference (0 for anything but 90°/270°).
 */
export function quarterTurnInsetPx(
  widthPx: number,
  heightPx: number,
  rotDeg: number | null | undefined,
): number {
  const rot = ((Math.round(Number(rotDeg ?? 0)) % 360) + 360) % 360
  if (rot !== 90 && rot !== 270) return 0
  if (!(widthPx > 0) || !(heightPx > 0)) return 0
  return (widthPx - heightPx) / 2
}

/** margins that grow an in-flow picture's layout box to its quarter-turned bounding box */
export function quarterTurnMarginCss(inset: number): string {
  if (!inset) return ''
  return `margin:${inset.toFixed(1)}px ${(-inset).toFixed(1)}px`
}

const SIDES = ['top', 'right', 'bottom', 'left'] as const
const SIDE_INSET: Record<(typeof SIDES)[number], 1 | -1> = {
  top: 1,
  bottom: 1,
  left: -1,
  right: -1,
}

/**
 * Same footprint growth for a box that already carries positioning margins
 * (`;`-separated declarations): each side's inset folds into the declared
 * value instead of replacing it, so posOffset lifts and indents survive.
 */
export function foldQuarterTurnMargins(inset: number, declarations: string): string {
  const decls = declarations.split(';').filter(Boolean)
  if (!inset) return decls.join(';')
  const rest: string[] = []
  const margins: Record<string, string> = {}
  for (const decl of decls) {
    const m = /^margin-(top|right|bottom|left):(.+)$/.exec(decl.trim())
    if (m) margins[m[1]] = m[2].trim()
    else rest.push(decl)
  }
  for (const side of SIDES) {
    const add = SIDE_INSET[side] * inset
    const prior = margins[side]
    if (prior === undefined) margins[side] = `${add.toFixed(1)}px`
    else if (/^-?\d+(?:\.\d+)?px$/.test(prior))
      margins[side] = `${(parseFloat(prior) + add).toFixed(1)}px`
    else {
      const inner = /^calc\((.*)\)$/.exec(prior)?.[1] ?? prior
      margins[side] = `calc(${inner} + ${add.toFixed(1)}px)`
    }
  }
  return [...rest, ...SIDES.map((side) => `margin-${side}:${margins[side]}`)].join(';')
}

export interface PictureLayer {
  dataUrl: string
  x: number
  y: number
  w: number
  h: number
}

/**
 * A picture-only group paints its pictures bottom first as stacked background
 * images of one <img>, each at its fraction of the box. The src stays the
 * first picture so the element is a real image for every other path (the
 * engine keeps the group bytes); the pixels come from the backgrounds and the
 * src itself is shifted out of the box.
 */
export function pictureLayerDecls(layers: PictureLayer[] | null | undefined): string[] {
  if (!layers || layers.length === 0) return []
  const pct = (v: number) => `${(v * 100).toFixed(2)}%`
  const top = [...layers].reverse()
  return [
    `background-image:${top.map((l) => `url("${l.dataUrl}")`).join(',')}`,
    `background-size:${top.map((l) => `${pct(l.w)} ${pct(l.h)}`).join(',')}`,
    // percentage positions place against the free space, so solve for the
    // offset that lands the layer's left/top edge at its fraction
    `background-position:${top.map((l) => `${edge(l.x, l.w)} ${edge(l.y, l.h)}`).join(',')}`,
    'background-repeat:no-repeat',
    // the layers are the picture: print must keep them like any bitmap
    '-webkit-print-color-adjust:exact',
    'print-color-adjust:exact',
    // natural size shifted out of the box: a percent offset would resolve to
    // zero under the default fill and leave the src painting over the layers
    'object-fit:none',
    'object-position:-100000px -100000px',
  ]
}

function edge(start: number, size: number): string {
  const free = 1 - size
  return Math.abs(free) < 1e-6 ? '0%' : `${((start / free) * 100).toFixed(2)}%`
}
