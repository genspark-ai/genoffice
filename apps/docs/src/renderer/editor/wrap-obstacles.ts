import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import { insideFloatTable } from './pagination-gaps'
import { TopLevelPositions } from './top-level-pos'

const key = new PluginKey<DecorationSet>('wrapObstacles')

/** node decorations carrying the measured gap of each displaced paragraph
 *  (inline styles on plain paragraphs do not survive ProseMirror redraws) */
export const WrapObstaclesExtension = Extension.create({
  name: 'wrapObstacles',
  addProseMirrorPlugins() {
    return [
      new Plugin<DecorationSet>({
        key,
        state: {
          init: () => DecorationSet.empty,
          apply(tr, set) {
            const next = tr.getMeta(key) as DecorationSet | undefined
            return next ?? set.map(tr.mapping, tr.doc)
          },
        },
        props: {
          decorations(state) {
            return key.getState(state)
          },
        },
      }),
    ]
  },
})

/** wrap rectangle in column coordinates (text clearance included) */
export interface ObstacleRect {
  left: number
  right: number
  top: number
  bottom: number
}

export interface LinePlacement {
  top: number
  left: number
  right: number
}

/**
 * Word's per-line wrap rule (probes 2026-10-13): a line that crosses a
 * side-wrapped drawing takes the widest free gap of its column when the gap
 * holds the line's first word (no fixed minimum: a 40pt gap took a 25pt
 * word, a 60pt gap refused a 90pt one), else it moves down to the bottom of
 * the drawing that blocks it (not a whole-line step) and tries again.
 */
export function placeLine(
  rects: readonly ObstacleRect[],
  top: number,
  lineH: number,
  colLeft: number,
  colRight: number,
  needW: number,
): LinePlacement {
  let y = top
  const need = Math.max(needW, 1)
  for (let guard = 0; guard < 64; guard++) {
    const hit = rects.filter(
      (r) => r.top < y + lineH && r.bottom > y && r.right > colLeft && r.left < colRight,
    )
    if (hit.length === 0) return { top: y, left: colLeft, right: colRight }
    let best: [number, number] = [colLeft, colLeft]
    let cursor = colLeft
    for (const r of [...hit].sort((a, b) => a.left - b.left)) {
      if (r.left - cursor > best[1] - best[0]) best = [cursor, r.left]
      cursor = Math.max(cursor, r.right)
    }
    if (colRight - cursor > best[1] - best[0]) best = [cursor, colRight]
    if (best[1] - best[0] >= need) return { top: y, left: best[0], right: best[1] }
    y = Math.min(...hit.map((r) => r.bottom))
  }
  return { top: y, left: colLeft, right: colRight }
}

const SIDE_IMG_SEL = ['square', 'tight', 'through']
  .flatMap((w) => [`img.doc-inline-img--wrap-${w}-left`, `img.doc-inline-img--wrap-${w}-right`])
  .join(', ')

const distOf = (el: HTMLElement): [number, number, number, number] => {
  const d = (el.dataset.wrapDist ?? '').split(' ').map(Number)
  return [d[0] || 0, d[1] || 0, d[2] || 0, d[3] || 0]
}

/** flow blocks whose lines wrap: paragraphs, headings (h1-h6), list items */
const isTextBlock = (el: Element): boolean =>
  el.tagName === 'P' || /^H[1-6]$/.test(el.tagName) || el.classList.contains('doc-li')

const px = (v: number): string => `${(Math.round(v * 10) / 10).toFixed(1)}px`

interface Write {
  el: HTMLElement
  style: Record<string, string>
  data: Record<string, string | undefined>
}

interface ParaSpec {
  el: HTMLElement
  base: [number, number, number]
  pad: [number, number, number]
}

// last values written per element: a follow-up pass is only worth scheduling
// when this pass decided differently, not when it re-asserts a value another
// pre-measure pass (syncAnchorBands' strut drops) reset in between
const lastWritten = new WeakMap<HTMLElement, string>()

/**
 * Side-wrapped drawings of multi-drawing anchor paragraphs are absolutely
 * positioned boxes CSS never wraps around. Word treats every wrap rectangle
 * as an obstacle for every later line of the page: a heading under a photo
 * row flows into the free area beside a lower photo, empty paragraphs fill
 * the rows above a low photo row. This pass measures those rectangles (plus
 * the overlaid run pictures of paragraphs it already moved) and gives each
 * paragraph whose first line crosses one the measured gap as padding
 * (top = displacement, left/right = gap). Uniform per paragraph: later lines
 * keep the first line's gap. A page gap resets the obstacle set. Idempotent:
 * every value derives from undisplaced geometry (box offsets, paragraph
 * tops, first-word widths). Returns whether any write changed the layout.
 */
export function syncWrapObstacles(pm: HTMLElement, factor: number, view: EditorView): boolean {
  const pmRect = pm.getBoundingClientRect()
  const rel = (r: DOMRect): ObstacleRect => ({
    left: (r.left - pmRect.left) / factor,
    right: (r.right - pmRect.left) / factor,
    top: (r.top - pmRect.top) / factor,
    bottom: (r.bottom - pmRect.top) / factor,
  })
  const writes: Write[] = []
  const paras: ParaSpec[] = []
  let rects: ObstacleRect[] = []
  for (const el of Array.from(pm.children) as HTMLElement[]) {
    if (el.classList.contains('page-gap')) {
      if (!insideFloatTable(el)) rects = []
      continue
    }
    if (
      el.classList.contains('page-float-host') ||
      el.classList.contains('page-float-carry') ||
      el.classList.contains('page-repeat-header')
    ) {
      continue
    }
    if (el.classList.contains('doc-protected-floating')) {
      if (el.classList.contains('doc-protected-pagepinned')) continue
      collectWrapper(el, rel, rects, writes, factor)
      continue
    }
    if (!isTextBlock(el)) continue
    if (rects.length === 0 && el.dataset.obsMode !== '1') continue
    placeParagraph(el, rel, rects, paras, factor)
  }
  let changed = false
  for (const w of writes) {
    const sig = JSON.stringify([w.style, w.data])
    if (lastWritten.get(w.el) !== sig) {
      lastWritten.set(w.el, sig)
      changed = true
    }
    for (const [k, v] of Object.entries(w.style)) {
      if (w.el.style.getPropertyValue(k) === v) continue
      if (v) w.el.style.setProperty(k, v)
      else w.el.style.removeProperty(k)
    }
    for (const [k, v] of Object.entries(w.data)) {
      if (w.el.dataset[k] === v) continue
      if (v === undefined) delete w.el.dataset[k]
      else w.el.dataset[k] = v
    }
  }
  const positions = new TopLevelPositions(view)
  const decos: Decoration[] = []
  for (const p of paras) {
    const range = positions.of(p.el)
    if (!range) continue
    const style = `padding-top:${px(p.base[0] + p.pad[0])};padding-left:${px(p.base[1] + p.pad[1])};padding-right:${px(p.base[2] + p.pad[2])}`
    const base = p.base.map((v) => Math.round(v * 10) / 10).join(' ')
    decos.push(
      Decoration.node(
        range.from,
        range.to,
        { style, 'data-obs-mode': '1', 'data-obs-base': base },
        { key: `obs-${style}-${base}` },
      ),
    )
  }
  const prev = key.getState(view.state)
  const prevKeys = (prev ? prev.find() : []).map(specKey).join('|')
  const nextKeys = decos.map(specKey).join('|')
  if (prevKeys !== nextKeys) {
    const next = DecorationSet.create(view.state.doc, decos)
    view.dispatch(view.state.tr.setMeta(key, next).setMeta('addToHistory', false))
    changed = true
  }
  return changed
}

const specKey = (d: Decoration): string => `${d.from}:${(d.spec as { key?: string }).key ?? ''}`

function collectWrapper(
  el: HTMLElement,
  rel: (r: DOMRect) => ObstacleRect,
  rects: ObstacleRect[],
  writes: Write[],
  factor: number,
): void {
  const all = Array.from(el.querySelectorAll<HTMLElement>(':scope > .doc-textbox'))
  const sides = all.filter((b) => b.dataset.wrapSide === '1')
  const data: Write['data'] = {}
  if (sides.length === 0) {
    if (el.dataset.obsBottom !== undefined) {
      writes.push({ el, style: {}, data: { obsBottom: undefined, obsMaxh: undefined } })
    }
    restoreBand(el, writes)
    dropStrut(el, rel(el.getBoundingClientRect()), rects, writes, factor)
    return
  }
  const wr = rel(el.getBoundingClientRect())
  const wrTop = wr.top
  dropStrut(el, wr, rects, writes, factor)
  let bottom = 0
  let maxH = 0
  const own: ObstacleRect[] = []
  for (const box of sides) {
    // a page-positioned box is not anchor-relative; the display shift
    // (syncFloatShifts, recomputed after every slicing) is not geometry
    if (box.dataset.pageRelV === '1') continue
    const r = rel(box.getBoundingClientRect())
    if (r.bottom <= r.top) continue
    const dy = parseFloat(box.dataset.pageFloatDy ?? '0') || 0
    r.top -= dy
    r.bottom -= dy
    const [l, rr, t, b] = distOf(box)
    const o = { left: r.left - l, right: r.right + rr, top: r.top - t, bottom: r.bottom + b }
    rects.push(o)
    own.push(o)
    bottom = Math.max(bottom, r.bottom - wrTop)
    maxH = Math.max(maxH, r.bottom - r.top)
  }
  // the page-turn rule applies when every drawing of the anchor is a side
  // wrap; a band-keep wrapper keeps its own overflow rule
  if (sides.length === all.length && el.dataset.bandKeep !== '1' && bottom > 0) {
    data.obsBottom = px(bottom)
    data.obsMaxh = px(maxH)
  } else {
    data.obsBottom = undefined
    data.obsMaxh = undefined
  }
  writes.push({ el, style: {}, data })
  syncBand(el, sides.length === all.length, rects, wrTop, rel, writes, factor)
}

/**
 * A union band (data-band: the row's gaps are slivers) reserves the whole
 * band as flow height, which also pushes the paragraphs that Word lays in
 * the free rows above the drawings. With the drawings tracked as obstacles
 * the band collapses to the anchor line, but only for a lone wrapper whose
 * band region is followed by plain paragraphs (tables and further anchors
 * keep the band route and the anchor-run layout of syncAnchorBands).
 */
function syncBand(
  el: HTMLElement,
  allSides: boolean,
  // every obstacle of the page so far, the wrapper's own boxes included
  rects: ObstacleRect[],
  wrTop: number,
  rel: (r: DOMRect) => ObstacleRect,
  writes: Write[],
  factor: number,
): void {
  const band = Math.round(parseFloat(el.dataset.band ?? '0') || 0)
  if (
    !allSides ||
    band <= 0 ||
    el.dataset.bandKeep === '1' ||
    el.dataset.bandBeside === '1' ||
    el.dataset.bandAdj !== undefined ||
    isWrapper(el.previousElementSibling) ||
    isWrapper(el.nextElementSibling)
  ) {
    restoreBand(el, writes)
    return
  }
  const strut = strutOf(el)
  // syncAnchorBands pads the strut to drop the anchor line below the band
  const strutPad = strut ? parseFloat(strut.style.paddingTop) || 0 : 0
  const line = strut ? strut.getBoundingClientRect().height / factor - strutPad : 0
  if (!strut || line <= 0) {
    restoreBand(el, writes)
    return
  }
  let acc = line
  for (
    let sib = el.nextElementSibling as HTMLElement | null;
    sib && acc < band;
    sib = sib.nextElementSibling as HTMLElement | null
  ) {
    if (sib.classList.contains('page-gap') || sib.classList.contains('page-float-host')) continue
    const sibStrut = isWrapper(sib) ? strutOf(sib) : null
    if (!isTextBlock(sib) && !sibStrut) {
      restoreBand(el, writes)
      return
    }
    const cs = getComputedStyle(sib)
    acc +=
      (sibStrut ?? sib).getBoundingClientRect().height / factor +
      (parseFloat(cs.marginTop) || 0) +
      (parseFloat(cs.marginBottom) || 0)
  }
  // the anchor's own line sits beside its drawings when any gap is left
  // (probe: a V=0 picture shifts it sideways), else below the blocking one
  const wr = rel(el.getBoundingClientRect())
  const lineTop =
    wrTop + (strut.getBoundingClientRect().top - el.getBoundingClientRect().top) / factor
  const drop = placeLine(rects, lineTop, line, wr.left, wr.right, 0).top - lineTop
  writes.push({ el, style: { 'min-height': px(line + drop) }, data: { obsBand: '1' } })
  writes.push({
    el: strut,
    style: { 'padding-top': drop > 0.5 ? px(drop) : '' },
    data: { obsDrop: drop > 0.5 ? px(drop) : undefined },
  })
}

const strutOf = (el: HTMLElement): HTMLElement | null =>
  el.querySelector<HTMLElement>(':scope > .doc-anchor-strut, :scope > .doc-textbox-stray')

/**
 * The anchor paragraph's own line against EARLIER obstacles: it stays where
 * any gap is left (an empty line fits anywhere) and drops to the blocking
 * drawing's bottom otherwise. Only a drop this pass decided is written (or
 * cleared), so syncAnchorBands' own band drops stay untouched.
 */
function dropStrut(
  el: HTMLElement,
  wr: ObstacleRect,
  rects: ObstacleRect[],
  writes: Write[],
  factor: number,
): void {
  const strut = strutOf(el)
  if (!strut) return
  const strutPad = parseFloat(strut.style.paddingTop) || 0
  const ownPad = strut.dataset.obsDrop !== undefined
  const line = strut.getBoundingClientRect().height / factor - strutPad
  if (line <= 0) return
  const lineTop =
    wr.top + (strut.getBoundingClientRect().top - el.getBoundingClientRect().top) / factor
  const drop = rects.length
    ? placeLine(rects, lineTop, line, wr.left, wr.right, 0).top - lineTop
    : 0
  if (drop > 0.5) {
    writes.push({ el: strut, style: { 'padding-top': px(drop) }, data: { obsDrop: px(drop) } })
  } else if (ownPad) {
    writes.push({ el: strut, style: { 'padding-top': '' }, data: { obsDrop: undefined } })
  }
}

function restoreBand(el: HTMLElement, writes: Write[]): void {
  if (el.dataset.obsBand !== '1') return
  const band = Math.round(parseFloat(el.dataset.band ?? '0') || 0)
  writes.push({
    el,
    style: { 'min-height': band > 0 ? `${band}px` : '' },
    data: { obsBand: undefined },
  })
}

const isWrapper = (el: Element | null): boolean =>
  !!el &&
  el.classList.contains('doc-protected-floating') &&
  !el.classList.contains('doc-protected-pagepinned')

function firstWord(el: HTMLElement): Range | null {
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
    acceptNode: (n) =>
      /\S/.test(n.nodeValue ?? '') && !n.parentElement?.closest('.doc-textbox, .doc-anchor-strut')
        ? NodeFilter.FILTER_ACCEPT
        : NodeFilter.FILTER_SKIP,
  })
  const node = walker.nextNode() as Text | null
  if (!node) return null
  const m = /^(\s*)(\S+)/.exec(node.nodeValue ?? '')
  if (!m) return null
  const range = document.createRange()
  range.setStart(node, m[1].length)
  range.setEnd(node, m[1].length + m[2].length)
  return range
}

function placeParagraph(
  el: HTMLElement,
  rel: (r: DOMRect) => ObstacleRect,
  rects: ObstacleRect[],
  paras: ParaSpec[],
  factor: number,
): void {
  const cs = getComputedStyle(el)
  // base = the paragraph's own padding before this pass added to it
  let base = (el.dataset.obsBase ?? '').split(' ').map(Number)
  if (el.dataset.obsMode !== '1' || base.length !== 3 || base.some((v) => !Number.isFinite(v))) {
    base = [
      parseFloat(cs.paddingTop) || 0,
      parseFloat(cs.paddingLeft) || 0,
      parseFloat(cs.paddingRight) || 0,
    ]
  }
  const applied = el.dataset.obsMode === '1' ? (parseFloat(cs.paddingTop) || 0) - base[0] : 0
  const r = rel(el.getBoundingClientRect())
  const border = [
    parseFloat(cs.borderTopWidth) || 0,
    parseFloat(cs.borderLeftWidth) || 0,
    parseFloat(cs.borderRightWidth) || 0,
  ]
  const top = r.top + border[0] + base[0]
  const colLeft = r.left + border[1] + base[1]
  const colRight = r.right - border[2] - base[2]
  // summed fragment widths: layout-independent even when an earlier pass
  // left the word broken across lines in a too-narrow gap
  const word = firstWord(el)
  const needW = word
    ? Array.from(word.getClientRects()).reduce((w, r) => w + r.width, 0) / factor
    : 0
  const lh = parseFloat(cs.lineHeight)
  const lineH =
    Number.isFinite(lh) && lh > 0
      ? lh
      : Math.max(1, r.bottom - r.top - applied - base[0] - border[0])
  const external = rects.some(
    (o) => o.top < top + lineH && o.bottom > top && o.right > colLeft && o.left < colRight,
  )
  const imgs = Array.from(el.querySelectorAll<HTMLElement>(SIDE_IMG_SEL)).filter(
    (img) => !img.closest('.doc-textbox'),
  )
  if (!external) return
  // the paragraph's own run pictures become overlays at their anchor offsets
  // (data-obs-mode CSS); their rectangles join the obstacle set from here on
  const marginLeft = parseFloat(cs.getPropertyValue('--doc-margin-left')) || 0
  const own: ObstacleRect[] = []
  for (const img of imgs) {
    const b = img.getBoundingClientRect()
    if (b.height <= 0) continue
    const w = b.width / factor
    const h = b.height / factor
    const x = parseFloat(img.dataset.obsX ?? '')
    const left = Number.isFinite(x)
      ? colLeft + x - (img.dataset.obsXFrom === 'page' ? marginLeft : 0)
      : img.className.includes('-right')
        ? colRight - w
        : colLeft
    const y = top + (parseFloat(img.dataset.obsY ?? '') || 0)
    const [l, rr, t, bt] = distOf(img)
    own.push({ left: left - l, right: left + w + rr, top: y - t, bottom: y + h + bt })
  }
  const all = own.length ? [...rects, ...own] : rects
  const p = placeLine(all, top, lineH, colLeft, colRight, needW)
  rects.push(...own)
  paras.push({
    el,
    base: [base[0], base[1], base[2]],
    pad: [p.top - top, p.left - colLeft, colRight - p.right],
  })
}
