/** authored `transform` / margin declarations behind the image toolbar; everything is written as inline style */

export interface ImageTransform {
  /** clockwise degrees, normalised to 0 / 90 / 180 / 270 */
  rotate: number
  flipH: boolean
  flipV: boolean
}

export type ImageAlign = 'left' | 'center' | 'right' | null

const IDENTITY: ImageTransform = { rotate: 0, flipH: false, flipV: false }

const norm = (deg: number) => ((Math.round(deg) % 360) + 360) % 360

/** reads back the `rotate(..) scale(..)` form this module writes; any other authored transform starts from identity */
export function parseImageTransform(inline: string): ImageTransform {
  const rot = /rotate\((-?\d+(?:\.\d+)?)deg\)/.exec(inline)
  const scale = /scale\((-?1),\s*(-?1)\)/.exec(inline)
  return {
    rotate: rot ? norm(Number(rot[1])) : 0,
    flipH: scale ? scale[1] === '-1' : /scaleX\(-1\)/.test(inline),
    flipV: scale ? scale[2] === '-1' : /scaleY\(-1\)/.test(inline),
  }
}

/** null means identity: the declaration is removed rather than written as `none` */
export function serializeImageTransform(t: ImageTransform): string | null {
  const parts: string[] = []
  if (t.rotate) parts.push(`rotate(${t.rotate}deg)`)
  if (t.flipH || t.flipV) parts.push(`scale(${t.flipH ? -1 : 1}, ${t.flipV ? -1 : 1})`)
  return parts.length ? parts.join(' ') : null
}

export function rotateImage(t: ImageTransform, dir: -1 | 1): ImageTransform {
  return { ...t, rotate: norm(t.rotate + dir * 90) }
}

/**
 * Mirror the image as the user sees it on screen. With `rotate(θ) scale(sx, sy)` a screen-space
 * reflection commutes past the rotation as F·R(θ) = R(−θ)·F, so the angle flips sign too.
 */
export function flipImage(t: ImageTransform, axis: 'h' | 'v'): ImageTransform {
  return {
    rotate: norm(-t.rotate),
    flipH: axis === 'h' ? !t.flipH : t.flipH,
    flipV: axis === 'v' ? !t.flipV : t.flipV,
  }
}

export function transformStyles(t: ImageTransform): Record<string, string | null> {
  return { transform: serializeImageTransform(t) }
}

export const identityTransform = (): ImageTransform => ({ ...IDENTITY })

/** current alignment as authored by imageAlignStyles; null when the margins say something else */
export function imageAlignOf(marginLeft: string, marginRight: string): ImageAlign {
  const l = marginLeft.trim()
  const r = marginRight.trim()
  if (l === 'auto' && r === 'auto') return 'center'
  if (l === 'auto' && r && r !== 'auto') return 'right'
  if (r === 'auto' && l && l !== 'auto') return 'left'
  return null
}

/** a replaced element only obeys auto margins as a block */
export function imageAlignStyles(align: Exclude<ImageAlign, null>): Record<string, string> {
  return {
    display: 'block',
    'margin-left': align === 'left' ? '0' : 'auto',
    'margin-right': align === 'right' ? '0' : 'auto',
  }
}
