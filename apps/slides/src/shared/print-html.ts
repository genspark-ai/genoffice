/**
 * Print page assembly shared by the main process (hidden print window) and the
 * renderer's print-preview pane, so the preview is pixel-identical to the printout.
 * Pure string building — no Electron or DOM imports.
 */

export type PrintLayout = 'full' | 'handout2' | 'handout3' | 'handout6' | 'notes'
export type PrintOrientation = 'portrait' | 'landscape'

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
  /** Draw a thin border around full-page slides (handout/notes thumbnails are always framed) */
  frame?: boolean
  /** Preview mode: page gaps/shadows + page-number badges instead of a printable document */
  preview?: boolean
}

/** "1,3,5-8" → 0-based slide indices (1-based input over the whole deck); null = invalid */
export function parsePrintRange(text: string, max: number): number[] | null {
  const out = new Set<number>()
  const parts = text.split(/[,，、;；\s]+/).filter(Boolean)
  if (parts.length === 0) return null
  for (const part of parts) {
    const m = /^(\d+)\s*[-–]\s*(\d+)$|^(\d+)$/.exec(part)
    if (!m) return null
    const a = Number(m[1] ?? m[3])
    const b = Number(m[2] ?? m[3])
    if (a < 1 || b > max || a > b) return null
    for (let i = a; i <= b; i++) out.add(i - 1)
  }
  return [...out].sort((x, y) => x - y)
}

export function printPageCount(slideCount: number, layout: PrintLayout): number {
  const perPage =
    layout === 'handout2' ? 2 : layout === 'handout3' ? 3 : layout === 'handout6' ? 6 : 1
  return Math.ceil(slideCount / perPage)
}

/** A4 in inches; slide height for the 'full' layout follows the PDF export (7.5in) */
const A4_W = 8.27
const A4_H = 11.69
const SLIDE_H = 7.5

export function buildPrintDocumentHtml(o: PrintDocOptions): string {
  const layout = o.layout
  const isFull = layout === 'full'
  const landscape = !isFull && o.orientation === 'landscape'
  const slideW = Math.round(o.ratio * SLIDE_H * 1000) / 1000
  const pageW = isFull ? slideW : landscape ? A4_H : A4_W
  const pageH = isFull ? SLIDE_H : landscape ? A4_W : A4_H
  const perPage =
    layout === 'handout2' ? 2 : layout === 'handout3' ? 3 : layout === 'handout6' ? 6 : 1
  const esc = (x: string) =>
    x.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]!)

  let body: string
  if (isFull) {
    body = o.srcs.map((src) => `<div class="page"><img src="${src}"></div>`).join('')
  } else if (layout === 'notes') {
    // Notes page: slide on top + notes text below
    body = o.srcs
      .map(
        (src, i) =>
          `<div class="page notes"><img src="${src}">` +
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
            `<div class="cell"><img src="${src}">` +
            (perPage === 3 ? '<div class="rules"></div>' : '') +
            '</div>',
        )
        .join('')
      pages.push(`<div class="page handout h${perPage}">${cells}</div>`)
    }
    body = pages.join('')
  }

  // h6 grid: 2×3 portrait, 3×2 landscape
  const h6Cols = landscape ? '1fr 1fr 1fr' : '1fr 1fr'
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
.page.handout { padding: 0.4in; display: flex; flex-direction: ${landscape && perPage === 2 ? 'row' : 'column'}; gap: 0.24in; }
.page.handout .cell { display: flex; gap: 0.2in; align-items: center; flex: 1; min-width: 0; min-height: 0; }
.page.handout .cell img { border: 1px solid #bbb; object-fit: contain; max-height: 100%; }
.page.handout.h2 .cell img, .page.handout.h6 .cell img { width: 100%; height: auto; max-height: 100%; }
.page.handout.h3 .cell img { width: 55%; height: auto; }
.page.handout.h3 .rules {
  flex: 1; align-self: stretch;
  background: repeating-linear-gradient(#fff 0 0.28in, #ccc 0.28in calc(0.28in + 1px));
}
.page.handout.h6 { display: grid; grid-template-columns: ${h6Cols}; grid-auto-rows: 1fr; }
.page.notes { padding: 0.5in; display: flex; flex-direction: column; }
.page.notes img { width: ${landscape ? 'auto' : '100%'}; ${landscape ? 'max-height: 55%; align-self: center;' : 'height: auto;'} border: 1px solid #bbb; }
.page.notes .note { margin-top: 0.3in; font-size: 11pt; line-height: 1.5; white-space: pre-wrap; }
${frameCss}
${previewCss}
</style></head><body>${body}</body></html>`
}
