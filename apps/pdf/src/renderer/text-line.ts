/** Line grouping over pdf.js text-layer spans. pdf.js splits a visual line into runs
    at font changes and kerning breaks (CJK often one span per glyph), so span-level
    hit-testing selects a few characters out of a sentence. Editing wants the whole
    line: merge the clicked span with its same-baseline neighbors. Works in client-rect
    space so the spans' scaleX transforms are accounted for. */

import { chainLayers } from '../shared/x-layers'

export interface LineGroup {
  spans: HTMLElement[]
  /** Union of the group's client rects */
  rect: { left: number; top: number; right: number; bottom: number }
  /** Line text left-to-right, a space synthesized at visible gaps between runs */
  text: string
  /** Width-weighted dominant span height (client px) — the line's effective font size,
      so one oversized glyph doesn't inflate the whole rebuilt line */
  fontHeight: number
}

/** pdf.js 6 rotates runs via an inline --rotate custom property (styles.css applies it
    in the span transform), not an inline transform; unrotated spans don't carry it */
const rotated = (el: HTMLElement) => {
  const r = el.style.getPropertyValue('--rotate').trim()
  return (r !== '' && r !== '0deg') || /rotate\(/.test(el.style.transform)
}

export function groupLineSpans(anchor: HTMLElement): LineGroup {
  const single = (): LineGroup => {
    const r = anchor.getBoundingClientRect()
    return { spans: [anchor], rect: r, text: anchor.textContent ?? '', fontHeight: r.height }
  }
  const layer = anchor.closest('.textLayer')
  if (!layer || rotated(anchor)) return single()
  const ar = anchor.getBoundingClientRect()
  if (!(ar.height > 0)) return single()

  type Cand = { el: HTMLElement; r: DOMRect }
  const line: Cand[] = []
  for (const el of layer.querySelectorAll<HTMLElement>('span')) {
    if (el !== anchor) {
      if (!(el.textContent ?? '').trim() || el.querySelector('span') || rotated(el)) continue
      // Cheap same-line prefilter before forcing a client rect on every span
      if (Math.abs(el.offsetTop - anchor.offsetTop) > ar.height * 1.5) continue
    }
    const r = el.getBoundingClientRect()
    const vOverlap = Math.min(r.bottom, ar.bottom) - Math.max(r.top, ar.top)
    if (
      el !== anchor &&
      (Math.abs(r.bottom - ar.bottom) > ar.height * 0.4 ||
        vOverlap < Math.min(r.height, ar.height) * 0.5)
    )
      continue
    line.push({ el, r })
  }

  line.sort((a, b) => a.r.left - b.r.left || a.r.top - b.r.top)
  const ai = line.findIndex((c) => c.el === anchor)
  if (ai < 0) return single()
  // Z-stacked doppelgangers (an earlier overflowing edit drew a run over this line)
  // must not interleave into the text: chain the spans into layers and keep only the
  // anchor's own layer.
  const chains = chainLayers(
    line.map((c) => ({ left: c.r.left, right: c.r.right })),
    ar.height * 0.5,
  )
  const mine = line.filter((_, i) => chains[i] === chains[ai])
  const aIdx = mine.findIndex((c) => c.el === anchor)
  // Stop at gaps wider than ~a line-height: word gaps are far smaller, column gutters larger
  const gapMax = ar.height * 1.2
  let lo = aIdx
  while (lo > 0 && mine[lo]!.r.left - mine[lo - 1]!.r.right <= gapMax) lo--
  let hi = aIdx
  while (hi < mine.length - 1 && mine[hi + 1]!.r.left - mine[hi]!.r.right <= gapMax) hi++
  const group = mine.slice(lo, hi + 1)

  let text = ''
  let left = Infinity
  let top = Infinity
  let right = -Infinity
  let bottom = -Infinity
  const widthByHeight = new Map<number, number>()
  let prev: Cand | null = null
  for (const c of group) {
    const t = c.el.textContent ?? ''
    if (prev && c.r.left - prev.r.right > ar.height * 0.15 && !/\s$/.test(text) && !/^\s/.test(t))
      text += ' '
    text += t
    prev = c
    left = Math.min(left, c.r.left)
    top = Math.min(top, c.r.top)
    right = Math.max(right, c.r.right)
    bottom = Math.max(bottom, c.r.bottom)
    const h = Math.round(c.r.height * 2) / 2
    widthByHeight.set(h, (widthByHeight.get(h) ?? 0) + c.r.width)
  }
  let fontHeight = ar.height
  let bestW = -1
  for (const [h, w] of widthByHeight)
    if (w > bestW) {
      bestW = w
      fontHeight = h
    }
  return { spans: group.map((c) => c.el), rect: { left, top, right, bottom }, text, fontHeight }
}
