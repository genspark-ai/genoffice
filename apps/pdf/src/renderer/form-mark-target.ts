import type { FormWidget } from './form-catalog'

export type Rect = readonly [number, number, number, number]

const MIN_BOX = 5
const MAX_BOX = 26

const norm = (r: Rect): [number, number, number, number] => [
  Math.min(r[0], r[2]),
  Math.min(r[1], r[3]),
  Math.max(r[0], r[2]),
  Math.max(r[1], r[3]),
]

const contains = (r: Rect, x: number, y: number, slack: number): boolean => {
  const n = norm(r)
  return x >= n[0] - slack && x <= n[2] + slack && y >= n[1] - slack && y <= n[3] + slack
}

/** Interactive check box whose widget sits under the point (PDF user space) */
export function checkboxWidgetAt(
  widgets: readonly FormWidget[],
  x: number,
  y: number,
  slack = 1.5,
): FormWidget | null {
  return (
    widgets.find((w) => w.kind === 'checkbox' && !w.readOnly && contains(w.rect, x, y, slack)) ??
    null
  )
}

/** Smallest printed square-ish box (drawn path bbox) under the point, so a mark can fill it */
export function printedBoxAt(
  shapes: readonly Rect[],
  x: number,
  y: number,
  slack = 2,
): [number, number, number, number] | null {
  let best: [number, number, number, number] | null = null
  for (const s of shapes) {
    const n = norm(s)
    const w = n[2] - n[0]
    const h = n[3] - n[1]
    if (w < MIN_BOX || h < MIN_BOX || w > MAX_BOX || h > MAX_BOX) continue
    if (w / h > 1.35 || h / w > 1.35) continue
    if (!contains(n, x, y, slack)) continue
    if (!best || w * h < (best[2] - best[0]) * (best[3] - best[1])) best = n
  }
  return best
}

/** Square mark rect centered on the box, sized to its larger side */
export function markRectForBox(box: Rect): [number, number, number, number] {
  const n = norm(box)
  const side = Math.max(n[2] - n[0], n[3] - n[1])
  const cx = (n[0] + n[2]) / 2
  const cy = (n[1] + n[3]) / 2
  return [cx - side / 2, cy - side / 2, cx + side / 2, cy + side / 2]
}
