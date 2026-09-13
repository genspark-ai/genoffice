import type { LineNumbering, SectionInfo } from '@genoffice/docx-engine'
import { sectionBidi, sectionColGeom } from '../pagination-sections'
import type { BlockBox, BlockMetaOf } from '../pagination-types'
import { pageFramesFromGaps } from './pagination-gaps'

/** Word's gap for an omitted w:distance (measured against Word output: 0.25in). */
export const LINE_NUMBER_DEFAULT_DISTANCE_TWIPS = 360

const twipsToPx = (twips: number) => (twips / 1440) * 96
const NUMBER_BOX_W = 64

/** subtrees inside a block that are not body text: page-gap widgets, float hosts, repeated table headers */
const SKIP_SELECTOR =
  '.page-gap, .page-gap-inline, .page-gap-cut, .page-float-host, .page-float-carry, .page-repeat-header'

export interface LineRect {
  top: number
  bottom: number
  left: number
  right: number
}

/** Merge inline fragment rects into line boxes: fragments sharing most of their vertical extent form one line. */
export function groupLineRects(rects: LineRect[]): LineRect[] {
  const lines: LineRect[] = []
  for (const r of [...rects].sort((a, b) => a.top - b.top)) {
    if (r.bottom - r.top <= 0) continue
    const cur = lines[lines.length - 1]
    if (cur) {
      const overlap = Math.min(cur.bottom, r.bottom) - Math.max(cur.top, r.top)
      if (overlap >= 0.5 * Math.min(cur.bottom - cur.top, r.bottom - r.top)) {
        cur.top = Math.min(cur.top, r.top)
        cur.bottom = Math.max(cur.bottom, r.bottom)
        cur.left = Math.min(cur.left, r.left)
        cur.right = Math.max(cur.right, r.right)
        continue
      }
    }
    lines.push({ ...r })
  }
  return lines
}

/** Line boxes of one paragraph element (viewport coords) from its text, inline images and line breaks. */
export function paragraphLineRects(el: HTMLElement): LineRect[] {
  const rects: LineRect[] = []
  const push = (list: DOMRectList) => {
    for (const r of Array.from(list)) {
      if (r.height > 0) rects.push({ top: r.top, bottom: r.bottom, left: r.left, right: r.right })
    }
  }
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT, {
    acceptNode: (node) => {
      if (node instanceof Element) {
        if (node.matches(SKIP_SELECTOR)) return NodeFilter.FILTER_REJECT
        return node.tagName === 'IMG' || node.tagName === 'BR'
          ? NodeFilter.FILTER_ACCEPT
          : NodeFilter.FILTER_SKIP
      }
      return NodeFilter.FILTER_ACCEPT
    },
  })
  const range = document.createRange()
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node instanceof Element) {
      push(node.getClientRects())
    } else if (node.nodeValue && node.nodeValue.length > 0) {
      range.selectNodeContents(node)
      push(range.getClientRects())
    }
  }
  return groupLineRects(rects)
}

export interface LineRef {
  page: number
  col: number
  top: number
  section: number
}

/**
 * Number lines in reading order (page, column, top). Every line counts; a label
 * is emitted only for multiples of countBy. Restart per page / section per w:restart.
 */
export function numberLines<T extends LineRef>(
  lines: T[],
  lnOf: (section: number) => LineNumbering | undefined,
): Array<{ line: T; label: string }> {
  const sorted = [...lines].sort((a, b) => a.page - b.page || a.col - b.col || a.top - b.top)
  const out: Array<{ line: T; label: string }> = []
  let n = 0
  let started = false
  let prevPage = -1
  let prevSection = -1
  for (const line of sorted) {
    const ln = lnOf(line.section)
    if (!ln) continue
    const restart =
      !started ||
      (ln.restart === 'newPage' && line.page !== prevPage) ||
      (ln.restart === 'newSection' && line.section !== prevSection)
    n = restart ? ln.start : n + 1
    started = true
    prevPage = line.page
    prevSection = line.section
    if (n % ln.countBy === 0) out.push({ line, label: String(n) })
  }
  return out
}

export const hasLineNumbering = (sections: SectionInfo[]): boolean =>
  sections.some((s) => s.settings.lineNumbers !== undefined)

const lnOfSections = (sections: SectionInfo[]) => (i: number) =>
  sections[Math.min(i, sections.length - 1)]?.settings.lineNumbers

const distancePx = (ln: LineNumbering) =>
  twipsToPx(ln.distance ?? LINE_NUMBER_DEFAULT_DISTANCE_TWIPS)

/** left edge of each text column (reading order) relative to the left margin (px);
 *  a w:bidi section fills its columns right-to-left */
export function columnLefts(s: SectionInfo): number[] {
  const geom = sectionColGeom(s)
  const xs: number[] = []
  let x = 0
  for (let c = 0; c < geom.cols; c++) {
    xs.push(x)
    x += geom.widths[c] + (geom.gaps[c] ?? 0)
  }
  return sectionBidi(s) ? xs.map((left, c) => x - left - geom.widths[c]) : xs
}

/** reading-order index of the column whose left edge is the rightmost one at or before x
 *  (the leftmost column when x precedes them all); lefts may be mirrored (w:bidi) */
export function columnAt(lefts: number[], x: number): number {
  let col = -1
  for (let c = 0; c < lefts.length; c++) {
    if (x >= lefts[c] - 0.5 && (col < 0 || lefts[c] > lefts[col])) col = c
  }
  if (col >= 0) return col
  return lefts.reduce((best, left, c) => (left < lefts[best] ? c : best), 0)
}

/** Word's "Line Number" style is the default paragraph font: copy it from the document root. */
function applyDocFont(layer: HTMLElement, docRoot: Element): void {
  const cs = getComputedStyle(docRoot)
  layer.style.fontFamily = cs.fontFamily
  layer.style.fontSize = cs.fontSize
  layer.style.color = cs.color
}

function appendMark(
  layer: HTMLElement,
  cls: string,
  top: number,
  height: number,
  right: number,
  label: string,
): void {
  const el = document.createElement('div')
  el.className = cls
  el.style.top = `${top}px`
  el.style.height = `${height}px`
  el.style.lineHeight = `${height}px`
  el.style.left = `${right - NUMBER_BOX_W}px`
  el.style.width = `${NUMBER_BOX_W}px`
  el.textContent = label
  layer.appendChild(el)
}

const FLOAT_WRAP_RE = /(?:^|\s)img-wrap-(?:square|tight|through)-(?:left|right)(?:\s|$)/

/** blocks whose lines Word does not count: tables, protected objects, floats (same set as the canvas path's `floated`) */
const skipBlock = (el: HTMLElement) =>
  el.tagName === 'TABLE' ||
  el.hasAttribute('data-doc-protected') ||
  el.classList.contains('pv-prune-spacer') ||
  el.classList.contains('page-float-host') ||
  el.classList.contains('page-float-carry') ||
  FLOAT_WRAP_RE.test(el.className)

interface CanvasLine extends LineRef {
  height: number
  right: number
}

/**
 * Canvas line numbers (sectPr w:lnNumType): one absolutely positioned label per
 * counted line on the page wrap, right-aligned `distance` before the text column.
 * Tables, protected objects and w:suppressLineNumbers paragraphs are skipped;
 * header/footer strips live outside the block list.
 */
export function syncLineNumbers(
  wrap: HTMLElement,
  pm: HTMLElement,
  blocks: BlockBox[],
  sections: SectionInfo[],
  zoomFactor: number,
  metaOf?: BlockMetaOf,
  first?: { left: number; width: number },
): void {
  let layer = wrap.querySelector(':scope > .page-line-numbers') as HTMLElement | null
  if (!hasLineNumbering(sections)) {
    layer?.remove()
    return
  }
  if (!layer) {
    layer = document.createElement('div')
    layer.className = 'page-line-numbers'
    wrap.appendChild(layer)
  }
  layer.textContent = ''
  applyDocFont(layer, pm)
  const frames = pageFramesFromGaps(wrap, zoomFactor, first)
  const wr = wrap.getBoundingClientRect()
  const lines: CanvasLine[] = []
  for (const b of blocks) {
    if (!b.el || b.tableRows || b.floated || skipBlock(b.el)) continue
    const sec = Math.min(b.section ?? 0, sections.length - 1)
    const ln = sections[sec]?.settings.lineNumbers
    if (!ln) continue
    if (b.docxIndex !== undefined && metaOf?.(b.docxIndex)?.suppressLineNumbers) continue
    const set = sections[sec].settings
    const xs = columnLefts(sections[sec])
    for (const r of paragraphLineRects(b.el)) {
      const top = (r.top - wr.top) / zoomFactor
      const height = (r.bottom - r.top) / zoomFactor
      const mid = top + height / 2
      const page = frames.findIndex((f) => mid >= f.top && mid < f.bottom)
      if (page < 0) continue
      const textLeft = frames[page].left + twipsToPx(set.marginLeft)
      const rel = (r.left - wr.left) / zoomFactor - textLeft
      const col = columnAt(xs, rel)
      lines.push({
        page,
        col,
        top,
        height,
        section: sec,
        right: textLeft + xs[col] - distancePx(ln),
      })
    }
  }
  for (const { line, label } of numberLines(lines, lnOfSections(sections))) {
    appendMark(layer, 'page-line-number', line.top, line.height, line.right, label)
  }
}

interface PreviewLine extends LineRef {
  host: HTMLElement
  height: number
  right: number
}

/**
 * Pagination preview / print copy: labels per .pv-page sheet, measured on the
 * cloned blocks inside each page window (.pv-clip). Column = the .pv-col index;
 * the label x comes from the section's margin and column geometry (page-relative).
 * The layer lives inside .pv-sheet so it follows the zoom and the markup scale.
 */
export function syncPreviewLineNumbers(
  root: HTMLElement,
  sections: SectionInfo[],
  metaOf?: BlockMetaOf,
): void {
  const pages = Array.from(root.querySelectorAll<HTMLElement>('.pv-page'))
  if (!hasLineNumbering(sections)) {
    for (const p of pages) p.querySelector('.pv-line-numbers')?.remove()
    return
  }
  const sectionOfIdx = (idx: number, fallback: number) => {
    const i = sections.findIndex((s) => idx >= s.firstBlockIndex && idx <= s.lastBlockIndex)
    return i >= 0 ? i : fallback
  }
  const lines: PreviewLine[] = []
  pages.forEach((pageEl, page) => {
    const host = pageEl.querySelector<HTMLElement>(':scope > .pv-sheet') ?? pageEl
    host.querySelector(':scope > .pv-line-numbers')?.remove()
    const hostRect = host.getBoundingClientRect()
    const k =
      hostRect.width / (host.offsetWidth || parseFloat(pageEl.style.width) || hostRect.width) || 1
    const pageSection = Number(pageEl.getAttribute('data-pv-section') ?? -1)
    for (const clip of Array.from(pageEl.querySelectorAll<HTMLElement>('.pv-clip'))) {
      if (clip.closest('.pv-footnotes')) continue
      const colEl = clip.closest<HTMLElement>('.pv-col')
      const col = colEl?.parentElement ? Array.from(colEl.parentElement.children).indexOf(colEl) : 0
      const clipRect = clip.getBoundingClientRect()
      const content = clip.querySelector<HTMLElement>('.pv-content')
      if (!content) continue
      for (const el of Array.from(content.children) as HTMLElement[]) {
        if (skipBlock(el)) continue
        const idx = Number(el.getAttribute('data-idx'))
        const sec = Math.min(
          Number.isFinite(idx) && el.hasAttribute('data-idx')
            ? sectionOfIdx(idx, pageSection)
            : pageSection,
          sections.length - 1,
        )
        const ln = sections[sec]?.settings.lineNumbers
        if (!ln) continue
        if (el.hasAttribute('data-idx') && metaOf?.(idx)?.suppressLineNumbers) continue
        const colLeft =
          twipsToPx(sections[sec].settings.marginLeft) + (columnLefts(sections[sec])[col] ?? 0)
        for (const r of paragraphLineRects(el)) {
          const mid = (r.top + r.bottom) / 2
          if (mid < clipRect.top || mid >= clipRect.bottom) continue
          lines.push({
            host,
            page,
            col,
            section: sec,
            top: (r.top - hostRect.top) / k,
            height: (r.bottom - r.top) / k,
            right: colLeft - distancePx(ln),
          })
        }
      }
    }
  })
  const layers = new Map<HTMLElement, HTMLElement>()
  for (const { line, label } of numberLines(lines, lnOfSections(sections))) {
    let layer = layers.get(line.host)
    if (!layer) {
      layer = document.createElement('div')
      layer.className = 'pv-line-numbers'
      const content = line.host.querySelector('.pv-content')
      if (content) applyDocFont(layer, content)
      line.host.appendChild(layer)
      layers.set(line.host, layer)
    }
    appendMark(layer, 'pv-line-number', line.top, line.height, line.right, label)
  }
}
