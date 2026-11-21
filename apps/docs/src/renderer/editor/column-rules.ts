import type { SectionInfo } from '@genoffice/docx-engine'
import type { PageCut } from '../pagination-cuts'
import { sectionColGeom, sectionColumns } from '../pagination-sections'
import type { BlockBox } from '../pagination-types'
import { columnLefts, paragraphLineRects } from './line-numbers'
import { pageFramesFromGaps } from './pagination-gaps'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** the section draws a line between its columns (w:cols w:sep) and has columns to draw it between */
export const hasColumnRule = (s: SectionInfo | undefined): boolean =>
  !!s && !!s.settings.colSep && sectionColumns(s) > 1

/** x of each line, from the left edge of the text area: the middle of every gap between two columns */
export function columnRuleXs(s: SectionInfo): number[] {
  const { widths } = sectionColGeom(s)
  const edges = columnLefts(s)
    .map((left, c) => ({ left, right: left + widths[c] }))
    .sort((a, b) => a.left - b.left)
  return edges.slice(1).map((col, i) => (edges[i].right + col.left) / 2)
}

interface Extent {
  top: number
  bottom: number
}

function grow(extents: Map<string, Extent>, key: string, top: number, bottom: number): void {
  const cur = extents.get(key)
  if (cur) {
    cur.top = Math.min(cur.top, top)
    cur.bottom = Math.max(cur.bottom, bottom)
  } else {
    extents.set(key, { top, bottom })
  }
}

function appendRule(layer: HTMLElement, cls: string, x: number, extent: Extent): void {
  const el = document.createElement('div')
  el.className = cls
  el.style.left = `${x}px`
  el.style.top = `${extent.top}px`
  el.style.height = `${extent.bottom - extent.top}px`
  layer.appendChild(el)
}

/**
 * The line between columns on the canvas: per page and section, one rule in
 * the middle of each column gap, as tall as the section's text on that page
 * (Word stops it at the longest column). Measured on the laid-out lines, so it
 * follows whichever way the columns were painted.
 */
export function syncColumnRules(
  wrap: HTMLElement,
  pm: HTMLElement,
  blocks: BlockBox[],
  sections: SectionInfo[],
  zoomFactor: number,
  first?: { left: number; width: number },
  cuts?: readonly PageCut[],
): void {
  let layer = wrap.querySelector(':scope > .page-column-rules') as HTMLElement | null
  if (!sections.some(hasColumnRule)) {
    layer?.remove()
    return
  }
  if (!layer) {
    layer = document.createElement('div')
    layer.className = 'page-column-rules'
    wrap.appendChild(layer)
  }
  layer.textContent = ''
  layer.style.color = getComputedStyle(pm).color
  const frames = pageFramesFromGaps(wrap, zoomFactor, first, cuts)
  const wr = wrap.getBoundingClientRect()
  const extents = new Map<string, Extent>()
  for (const b of blocks) {
    if (!b.el || b.floated) continue
    const sec = Math.min(b.section ?? 0, sections.length - 1)
    if (!hasColumnRule(sections[sec])) continue
    for (const r of paragraphLineRects(b.el)) {
      const top = (r.top - wr.top) / zoomFactor
      const bottom = (r.bottom - wr.top) / zoomFactor
      const mid = (top + bottom) / 2
      const page = frames.findIndex((f) => mid >= f.top && mid < f.bottom)
      if (page >= 0) grow(extents, `${page}:${sec}`, top, bottom)
    }
  }
  for (const [key, extent] of extents) {
    const [page, sec] = key.split(':').map(Number)
    const textLeft = frames[page].left + twipsToPx(sections[sec].settings.marginLeft)
    for (const x of columnRuleXs(sections[sec]))
      appendRule(layer, 'page-column-rule', textLeft + x, extent)
  }
}

/**
 * Pagination preview / print copy: the rules per column region of each
 * .pv-page sheet, placed between the region's column boxes and as tall as the
 * lines its windows show.
 */
export function syncPreviewColumnRules(root: HTMLElement, sections: SectionInfo[]): void {
  for (const pageEl of Array.from(root.querySelectorAll<HTMLElement>('.pv-page'))) {
    const host = pageEl.querySelector<HTMLElement>(':scope > .pv-sheet') ?? pageEl
    host.querySelector(':scope > .pv-column-rules')?.remove()
    const hostRect = host.getBoundingClientRect()
    const k =
      hostRect.width / (host.offsetWidth || parseFloat(pageEl.style.width) || hostRect.width) || 1
    let layer: HTMLElement | null = null
    for (const region of Array.from(pageEl.querySelectorAll<HTMLElement>('.pv-region'))) {
      const sec = Number(region.getAttribute('data-pv-region-section') ?? -1)
      if (!hasColumnRule(sections[sec])) continue
      const cols = Array.from(region.querySelectorAll<HTMLElement>(':scope > .pv-col'))
        .map((col) => col.getBoundingClientRect())
        .sort((a, b) => a.left - b.left)
      const extents = new Map<string, Extent>()
      for (const clip of Array.from(region.querySelectorAll<HTMLElement>('.pv-clip'))) {
        const clipRect = clip.getBoundingClientRect()
        const content = clip.querySelector<HTMLElement>('.pv-content')
        for (const el of Array.from(content?.children ?? []) as HTMLElement[]) {
          for (const r of paragraphLineRects(el)) {
            const mid = (r.top + r.bottom) / 2
            if (mid >= clipRect.top && mid < clipRect.bottom)
              grow(extents, '', (r.top - hostRect.top) / k, (r.bottom - hostRect.top) / k)
          }
        }
      }
      const extent = extents.get('')
      if (!extent) continue
      if (!layer) {
        layer = document.createElement('div')
        layer.className = 'pv-column-rules'
        const content = host.querySelector('.pv-content')
        if (content) layer.style.color = getComputedStyle(content).color
        host.appendChild(layer)
      }
      for (let c = 1; c < cols.length; c++) {
        const x = ((cols[c - 1].right + cols[c].left) / 2 - hostRect.left) / k
        appendRule(layer, 'pv-column-rule', x, extent)
      }
    }
  }
}
