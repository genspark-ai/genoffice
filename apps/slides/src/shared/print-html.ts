/**
 * Print page assembly shared by the main process (hidden print window) and the
 * renderer's print-preview pane, so the preview is pixel-identical to the printout.
 * Pure string building — no Electron or DOM imports.
 */

export type PrintLayout =
  'full' | 'handout2' | 'handout3' | 'handout4' | 'handout6' | 'handout9' | 'notes'
export type PrintOrientation = 'portrait' | 'landscape'
export type PrintOrder = 'horizontal' | 'vertical'

export interface PrintDocOptions {
  /** One image URL per slide, in print order (data: or blob: URLs) */
  srcs: string[]
  /** Slide aspect ratio (width / height); sets the page size of the 'full' layout */
  ratio: number
  layout: PrintLayout
  /** Per-slide speaker notes for the 'notes' layout (same order as srcs) */
  notes?: string[]
  /** Paper orientation for handouts/notes pages ('full' pages always follow the slide ratio) */
  orientation?: PrintOrientation
  /** Thumbnail order across the handout grid (row-first vs column-first) */
  order?: PrintOrder
  /** Draw a thin border around full-page slides (handout/notes thumbnails are always framed) */
  frame?: boolean
  /** Preview mode: page gaps/shadows + page-number badges instead of a printable document */
  preview?: boolean
}

/** "1,3,5-8" → 0-based slide indices (1-based input over the whole deck); null = invalid */
export function parsePrintRange(text: string, max: number): number[] | null {
  const out = new Set<number>()
  const tokenRe = /\d+\s*[-–—~～〜]\s*\d+|\d+/g
  const parts = text.match(tokenRe) ?? []
  if (parts.length === 0) return null
  if (text.replace(tokenRe, '').replace(/[,，、;；\s]/g, '') !== '') return null
  for (const part of parts) {
    const m = /^(\d+)\s*[-–—~～〜]\s*(\d+)$|^(\d+)$/.exec(part)
    if (!m) return null
    const a = Number(m[1] ?? m[3])
    const b = Number(m[2] ?? m[3])
    if (a < 1 || b > max || a > b) return null
    for (let i = a; i <= b; i++) out.add(i - 1)
  }
  return [...out].sort((x, y) => x - y)
}

/**
 * Resolve the "current slide" print range. The dialog can outlive deck edits
 * (undo/slide delete), so a stale index yields no selection instead of a
 * crash on slides[current].hidden.
 */
export function currentRangeIndices(current: number, slideCount: number): number[] {
  return Number.isInteger(current) && current >= 0 && current < slideCount ? [current] : []
}

/** Thumbnails per printed sheet for each handout layout (1 for full/notes) */
export function printThumbnailsPerPage(layout: PrintLayout): number {
  switch (layout) {
    case 'handout2':
      return 2
    case 'handout3':
      return 3
    case 'handout4':
      return 4
    case 'handout6':
      return 6
    case 'handout9':
      return 9
    default:
      return 1
  }
}

export function printPageCount(slideCount: number, layout: PrintLayout): number {
  return Math.ceil(slideCount / printThumbnailsPerPage(layout))
}

/// Handout grid geometry (columns × rows). handout3 keeps its single column of
/// slide + ruled lines, so it is not a grid and the order option has nothing
/// to swap there. handout2 follows the order option like PowerPoint's: a
/// landscape sheet lays the pair out side by side only in horizontal order.
function handoutGrid(
  layout: PrintLayout,
  landscape: boolean,
  order: PrintOrder,
): { cols: number; rows: number } | null {
  switch (layout) {
    case 'handout2':
      return order === 'vertical' || !landscape ? { cols: 1, rows: 2 } : { cols: 2, rows: 1 }
    case 'handout4':
      return { cols: 2, rows: 2 }
    case 'handout6':
      return landscape ? { cols: 3, rows: 2 } : { cols: 2, rows: 3 }
    case 'handout9':
      return { cols: 3, rows: 3 }
    default:
      return null
  }
}

/** A4 in inches; slide height for the 'full' layout follows the PDF export (7.5in) */
const A4_W = 8.27
const A4_H = 11.69
const SLIDE_H = 7.5

/**
 * Fallback slide ratio (16:9) when layout state carries a corrupt ratio.
 * Ratios come from slide dimensions that can be zeroed by a failed load or
 * NaN/Infinity from a degenerate transform; emitting those into @page yields
 * NaNin/Infinityin/0in and a broken print. Bounds [0.2, 5] keep real formats
 * (4:3, 16:9, 16:10, portrait variants) exact while clamping absurd values.
 */
export const DEFAULT_PRINT_RATIO = 16 / 9
export const MIN_PRINT_RATIO = 0.2
export const MAX_PRINT_RATIO = 5

export function normalizePrintRatio(ratio: number): number {
  if (!Number.isFinite(ratio) || ratio <= 0) return DEFAULT_PRINT_RATIO
  return Math.min(Math.max(ratio, MIN_PRINT_RATIO), MAX_PRINT_RATIO)
}

/**
 * The print sheet's paper in inches — the same numbers the @page rule below
 * bakes into the document, so a silent CUPS job can size printToPDF's pages
 * to the layout instead of Chromium's Letter default. Full pages follow the
 * slide ratio; the handout/notes sheets are A4 in the chosen orientation.
 */
export function printPageSizeIn(
  layout: PrintLayout,
  orientation: PrintOrientation | undefined,
  ratio: number,
): { width: number; height: number } {
  const isFull = layout === 'full'
  const landscape = !isFull && orientation === 'landscape'
  const slideW = Math.round(normalizePrintRatio(ratio) * SLIDE_H * 1000) / 1000
  return {
    width: isFull ? slideW : landscape ? A4_H : A4_W,
    height: isFull ? SLIDE_H : landscape ? A4_W : A4_H,
  }
}

export function buildPrintDocumentHtml(o: PrintDocOptions): string {
  const layout = o.layout
  const isFull = layout === 'full'
  const landscape = !isFull && o.orientation === 'landscape'
  const { width: pageW, height: pageH } = printPageSizeIn(layout, o.orientation, o.ratio)
  const perPage = printThumbnailsPerPage(layout)
  const esc = (x: string) =>
    x.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]!)
  const escAttr = (x: string) =>
    x.replace(
      /[&<>"']/g,
      (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!,
    )

  let body: string
  if (isFull) {
    body = o.srcs.map((src) => `<div class="page"><img src="${escAttr(src)}"></div>`).join('')
  } else if (layout === 'notes') {
    // Notes page: slide on top + notes text below
    body = o.srcs
      .map(
        (src, i) =>
          `<div class="page notes"><img src="${escAttr(src)}">` +
          `<div class="note">${esc(o.notes?.[i] ?? '').replace(/\n/g, '<br>')}</div></div>`,
      )
      .join('')
  } else {
    // Handouts: perPage thumbnails per page (with 3, ruled lines on the right for handwriting)
    const pages: string[] = []
    for (let i = 0; i < o.srcs.length; i += perPage) {
      const cells = o.srcs
        .slice(i, i + perPage)
        .map(
          (src) =>
            `<div class="cell"><img src="${escAttr(src)}">` +
            (perPage === 3 ? '<div class="rules"></div>' : '') +
            '</div>',
        )
        .join('')
      pages.push(`<div class="page handout h${perPage}">${cells}</div>`)
    }
    body = pages.join('')
  }

  // Thumbnail grids: geometry per layout/orientation/order (handout3 stays flex)
  const grid = handoutGrid(layout, landscape, o.order ?? 'horizontal')
  const tracks = (n: number) => Array.from({ length: n }, () => '1fr').join(' ')
  const handoutGridCss = grid
    ? `.page.handout.h${perPage} { display: grid; grid-auto-flow: ${o.order === 'vertical' ? 'column' : 'row'}; grid-template-columns: ${tracks(grid.cols)}; grid-template-rows: ${tracks(grid.rows)}; }`
    : ''
  const frameCss = o.frame && isFull ? '.page > img { border: 1px solid #bbb; }' : ''
  // Preview chrome (never printed): gray gaps, page shadows, PowerPoint-style page badges
  const total = printPageCount(o.srcs.length, layout)
  const previewCss = o.preview
    ? `
body { counter-reset: pg; background: transparent; padding: 18px 0 6px; }
.page { counter-increment: pg; position: relative; margin: 0 auto 8px; background: #fff; box-shadow: 0 1px 5px rgba(0, 0, 0, 0.35); }
.page::after {
  content: counter(pg) ' / ${total}';
  position: absolute; left: 50%; bottom: 3%; transform: translateX(-50%);
  padding: 3px 12px; border-radius: 999px;
  background: rgba(32, 33, 36, 0.65); color: #fff;
  font-size: 10px; white-space: nowrap;
}
`
    : ''

  return `<!doctype html><html><head><meta charset="utf-8"><style>
@page { size: ${pageW}in ${pageH}in; margin: 0; }
html, body { margin: 0; padding: 0; font-family: -apple-system, 'Segoe UI', sans-serif; }
.page { width: ${pageW}in; height: ${pageH}in; overflow: hidden; page-break-after: always; box-sizing: border-box; }
.page:last-of-type { page-break-after: auto; }
.page > img { display: block; width: 100%; height: 100%; }
.page.handout { padding: 0.4in; gap: 0.24in; }
.page.handout.h3 { display: flex; flex-direction: column; }
${handoutGridCss}
.page.handout .cell { display: flex; gap: 0.2in; align-items: center; flex: 1; min-width: 0; min-height: 0; }
.page.handout .cell img { border: 1px solid #bbb; object-fit: contain; max-height: 100%; }
.page.handout.h2 .cell img, .page.handout.h4 .cell img, .page.handout.h6 .cell img, .page.handout.h9 .cell img { width: 100%; height: auto; max-height: 100%; }
.page.handout.h3 .cell img { width: 55%; height: auto; }
.page.handout.h3 .rules {
  flex: 1; align-self: stretch;
  background: repeating-linear-gradient(#fff 0 0.28in, #ccc 0.28in calc(0.28in + 1px));
}
.page.notes { padding: 0.5in; display: flex; flex-direction: column; }
.page.notes img { width: ${landscape ? 'auto' : '100%'}; ${landscape ? 'max-height: 55%; align-self: center;' : 'height: auto;'} border: 1px solid #bbb; }
.page.notes .note { margin-top: 0.3in; font-size: 11pt; line-height: 1.5; white-space: pre-wrap; }
${frameCss}
${previewCss}
</style></head><body>${body}</body></html>`
}
