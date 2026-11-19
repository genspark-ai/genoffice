/**
 * Two side-wrapped run pictures of one paragraph cannot both be CSS floats:
 * float placement compares margin boxes, and a margin-top offset makes every
 * float of the paragraph start at its first line, so the later float is pushed
 * beside the earlier one even when their pictures sit at different heights
 * (Word draws each at its declared offset, allowOverlap). The first picture of
 * a side stays the float and carries the union of the side's wrap rectangles
 * as a stair-shaped shape-outside; the others become overlays at their offsets.
 */
export interface WrapPicture {
  side: 'left' | 'right'
  /** anchor offset from the column start / paragraph top (px) */
  x: number
  y: number
  w: number
  h: number
  /** wrap distances (px): left, right, top, bottom */
  dist: [number, number, number, number]
  /** rotated / mirrored (transform about the box centre) */
  turned?: boolean
}

export interface WrapGroupCss {
  /** inline style of the picture at this index, or null when it floats on its own */
  styles: Array<string | null>
}

const px = (v: number): string => `${(Math.round(v * 10) / 10).toFixed(1)}px`

export function wrapGroupCss(pics: readonly WrapPicture[]): WrapGroupCss {
  const styles: Array<string | null> = pics.map(() => null)
  for (const side of ['left', 'right'] as const) {
    const idx = pics.map((p, i) => (p.side === side ? i : -1)).filter((i) => i >= 0)
    if (idx.length < 2) continue
    const rects = idx.map((i) => {
      const p = pics[i]
      return {
        left: p.x - p.dist[0],
        right: p.x + p.w + p.dist[1],
        top: p.y - p.dist[2],
        bottom: p.y + p.h + p.dist[3],
      }
    })
    const boxTop = Math.min(...rects.map((r) => r.top))
    const boxBottom = Math.max(...rects.map((r) => r.bottom))
    const boxLeft = Math.min(...rects.map((r) => r.left))
    const boxRight = Math.max(...rects.map((r) => r.right))
    // an unturned lead keeps its measured box honest (a transform-origin on a
    // padded box moves the rect getBoundingClientRect reports)
    const leadIdx = idx.find((i) => !pics[i].turned) ?? idx[0]
    const lead = pics[leadIdx]
    // margin-box origin: a right float's margin box starts at the union's
    // text-side edge, a left float's at the column start
    const originX = side === 'right' ? boxLeft : 0
    // the stair profile: per vertical band the text-side edge of the union.
    // The reference box is the margin box, whose top is the float's placement
    // line (margin-top: boxTop lies inside it), so y stays paragraph-relative
    const ys = [...new Set(rects.flatMap((r) => [r.top, r.bottom]))].sort((a, b) => a - b)
    const pts: string[] = []
    for (let i = 0; i + 1 < ys.length; i++) {
      const y0 = ys[i]
      const y1 = ys[i + 1]
      const hit = rects.filter((r) => r.top < y1 && r.bottom > y0)
      const edge =
        side === 'right'
          ? hit.length
            ? Math.min(...hit.map((r) => r.left))
            : boxRight
          : hit.length
            ? Math.max(...hit.map((r) => r.right))
            : boxLeft
      pts.push(`${px(edge - originX)} ${px(y0)}`, `${px(edge - originX)} ${px(y1)}`)
    }
    const far = side === 'right' ? boxRight - originX : 0
    pts.push(`${px(far)} ${px(boxBottom)}`, `${px(far)} ${px(boxTop)}`)
    // the float's margin box spans the union; padding places the picture in it
    const leadStyle = [
      'box-sizing:content-box',
      `margin-top:${px(boxTop)}`,
      `margin-left:${px(lead.x - originX)}`,
      side === 'right'
        ? `margin-right:calc(100% - ${px(lead.x + lead.w)})`
        : `margin-right:${px(boxRight - lead.x - lead.w)}`,
      `padding-top:${px(lead.y - boxTop)}`,
      `padding-bottom:${px(boxBottom - lead.y - lead.h)}`,
      `shape-outside:polygon(${pts.join(', ')})`,
      `--obs-top:${px(boxTop)}`,
    ]
    if (lead.turned) leadStyle.push(`transform-origin:50% ${px(lead.y - boxTop + lead.h / 2)}`)
    styles[leadIdx] = leadStyle.join(';')
    for (const i of idx) {
      if (i === leadIdx) continue
      const p = pics[i]
      styles[i] = `position:absolute;float:none;margin:0;left:${px(p.x)};top:${px(p.y)}`
    }
  }
  return { styles }
}
