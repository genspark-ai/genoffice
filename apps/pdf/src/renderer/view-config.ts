// cmaps/standard fonts/wasm are statically copied by the build into pdfjs/ of the renderer output (same path on the dev server)
export const ASSET_BASE = new URL('pdfjs/', document.baseURI).href

export const ZOOM_STEPS = [0.5, 0.67, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4]
export const MIN_SCALE = ZOOM_STEPS[0]
export const MAX_SCALE = ZOOM_STEPS[ZOOM_STEPS.length - 1]
export const PAGE_GAP = 16
export const SCROLL_PAD = 24
// ── Sidebar (thumbnails / outline) width: drag the divider to resize; persisted ──
export const SIDEBAR_W_KEY = 'genoffice-pdf-sidebar-width'
export const SIDEBAR_W_DEFAULT = 150
export const SIDEBAR_W_MIN = 120
/** pane padding (10px × 2) + thumb box borders (2px × 2) */
export const SIDEBAR_CHROME = 24

export const clampSidebarW = (w: number): number =>
  Math.min(Math.max(w, SIDEBAR_W_MIN), Math.min(320, Math.round(window.innerWidth * 0.4)))

export const loadSidebarW = (): number => {
  const saved = Number(localStorage.getItem(SIDEBAR_W_KEY))
  return Number.isFinite(saved) && saved > 0 ? clampSidebarW(saved) : SIDEBAR_W_DEFAULT
}

export interface PageSize {
  width: number
  height: number
}

export type FitMode = 'width' | 'page' | null

export const DOC_OPTS = {
  cMapUrl: `${ASSET_BASE}cmaps/`,
  cMapPacked: true,
  standardFontDataUrl: `${ASSET_BASE}standard_fonts/`,
  wasmUrl: `${ASSET_BASE}wasm/`,
}

/** Target paper sizes for "page size" (points, portrait) */
export const PAPER_SIZES = [
  { label: 'A3', w: 842, h: 1191 },
  { label: 'A4', w: 595, h: 842 },
  { label: 'A5', w: 420, h: 595 },
  { label: 'Letter', w: 612, h: 792 },
  { label: 'Legal', w: 612, h: 1008 },
] as const

/** Drawing stroke width (PDF pt); thin lines stay crisp under zoom */
export const STROKE_WIDTH = 2

/** Width of the WPS-style comments margin beside the pages */
export const NOTE_MARGIN_W = 300

/** Page ranges like "1-3,5" → list of 1-based page numbers; null if invalid */
export function parsePageRanges(input: string, max: number): number[] | null {
  const out = new Set<number>()
  for (const part of input.split(/[,，]/)) {
    const s = part.trim()
    if (!s) continue
    const m = /^(\d+)\s*[-–]\s*(\d+)$|^(\d+)$/.exec(s)
    if (!m) return null
    const a = Number(m[1] ?? m[3])
    const b = Number(m[2] ?? m[3])
    if (a < 1 || b > max || a > b) return null
    for (let i = a; i <= b; i++) out.add(i)
  }
  return out.size > 0 ? [...out].sort((x, y) => x - y) : null
}
