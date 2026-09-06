/**
 * Word-style dark page: the paper renders dark and authored document colors
 * are remapped for contrast — on screen only. The file model is never touched.
 *
 * Mechanism. Every site that writes an authored color inline keeps the
 * authored declaration and appends the remapped twin as a `--dk-*` custom
 * property (dkColor / dkBackground / dkBorder / dkTableBorders). Under
 * `.workspace.page-dark`, screen-only rules in styles.css switch the painted
 * property to the twin. Because the authored value stays the real inline
 * declaration:
 *   - the light theme / white page render exactly as before,
 *   - the pagination preview (outside `.page-dark`), printToPDF (print media)
 *     and clipboard HTML keep the authored colors,
 *   - toggling the page needs no editor redraw (pure CSS switch).
 *
 * Mapping. Univer's hue-preserving luminance inversion (RGB matrix: black ↔
 * white, mid-gray fixed, colored text keeps its hue), then compressed onto the
 * dark paper tone so authored white lands exactly on the paper color.
 */

/** dark paper tone; `--docs-paper` under `.page-dark` must equal darkPageColor('#ffffff') */
export const DARK_PAPER_HEX = '#1e1e1e'
const PAPER_FLOOR = 0x1e / 255

// Univer's invert-rgb matrix (hue-preserving luminance inversion)
const M = [
  [0.333, -0.667, -0.667],
  [-0.667, 0.333, -0.667],
  [-0.667, -0.667, 0.333],
] as const

const cache = new Map<string, string>()

function clamp01(v: number): number {
  return v > 1 ? 1 : v < 0 ? 0 : v
}

/** normalized rgb → remapped normalized rgb */
function remapRgb(r: number, g: number, b: number): [number, number, number] {
  const inv = M.map((row) => clamp01(row[0] * r + row[1] * g + row[2] * b + 1))
  // compress onto the paper tone: authored white → paper, authored black → white
  return inv.map((c) => PAPER_FLOOR + c * (1 - PAPER_FLOOR)) as [number, number, number]
}

function hex2(v: number): string {
  return Math.round(clamp01(v) * 255)
    .toString(16)
    .padStart(2, '0')
}

interface Parsed {
  r: number
  g: number
  b: number
  /** alpha as authored (CSS text), undefined when opaque */
  alpha?: string
}

const HEX_RE = /^#?([0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i
const RGB_RE =
  /^rgba?\(\s*([\d.]+)\s*[, ]\s*([\d.]+)\s*[, ]\s*([\d.]+)\s*(?:[,/]\s*([\d.]+%?)\s*)?\)$/i

function parse(css: string): Parsed | null {
  const s = css.trim()
  const hex = HEX_RE.exec(s)
  if (hex) {
    let h = hex[1]
    if (h.length <= 4) h = [...h].map((c) => c + c).join('')
    const n = parseInt(h.slice(0, 6), 16)
    const out: Parsed = {
      r: ((n >> 16) & 255) / 255,
      g: ((n >> 8) & 255) / 255,
      b: (n & 255) / 255,
    }
    if (h.length === 8) out.alpha = h.slice(6, 8)
    return out
  }
  const rgb = RGB_RE.exec(s)
  if (rgb) {
    const out: Parsed = {
      r: Number(rgb[1]) / 255,
      g: Number(rgb[2]) / 255,
      b: Number(rgb[3]) / 255,
    }
    if (rgb[4] !== undefined) out.alpha = rgb[4]
    return out
  }
  return null
}

/**
 * Remap one authored color for the dark page. Accepts `#rgb[a]`, `#rrggbb[aa]`,
 * bare OOXML `rrggbb`, and `rgb()/rgba()` (comma or space syntax); alpha is
 * kept as authored. Anything else (keywords, `transparent`, `auto`) comes back
 * unchanged.
 */
export function darkPageColor(css: string): string {
  const hit = cache.get(css)
  if (hit !== undefined) return hit
  const p = parse(css)
  let out = css
  if (p) {
    const [r, g, b] = remapRgb(p.r, p.g, p.b)
    if (p.alpha === undefined) out = `#${hex2(r)}${hex2(g)}${hex2(b)}`
    else if (/^[0-9a-f]{2}$/i.test(p.alpha) && /^#?[0-9a-f]{8}$/i.test(css.trim()))
      out = `#${hex2(r)}${hex2(g)}${hex2(b)}${p.alpha}`
    else
      out = `rgb(${Math.round(r * 255)} ${Math.round(g * 255)} ${Math.round(b * 255)} / ${p.alpha})`
  }
  cache.set(css, out)
  return out
}

/** `1px solid #000` → `1px solid #ffffff` (every color token inside a CSS border value) */
export function darkPageBorderCss(borderCss: string): string {
  return borderCss.replace(/#[0-9a-f]{3,8}\b|rgba?\([^)]*\)/gi, (m) => darkPageColor(m))
}

// ---- inline emission helpers: authored declaration + `--dk-*` twin -----------

/** text color twin (`color:#…` stays the authored declaration) */
export function dkColor(hex: string): string {
  return `--dk-c:${darkPageColor(hex)}`
}

/** background / shading / highlight / cell fill twin */
export function dkBackground(css: string): string {
  return `--dk-bg:${darkPageColor(css)}`
}

export type DkBorderSide = 't' | 'r' | 'b' | 'l'

/** element border twin (paragraph borders, cell borders) */
export function dkBorder(side: DkBorderSide, borderCss: string): string {
  return `--dk-b-${side}:${darkPageBorderCss(borderCss)}`
}

/**
 * Table-level border variables: `--doc-b-*` (authored) get `--dk-tb-*` twins
 * that the `.page-dark` cell rules read instead. Input is the authored
 * `--doc-b-x:value` declaration list from tableBordersCss.
 */
export function dkTableBorders(docBorderDecls: string[]): string[] {
  const out: string[] = []
  for (const decl of docBorderDecls) {
    const m = /^--doc-b-([trblhv]):(.*)$/.exec(decl)
    if (m) out.push(`--dk-tb-${m[1]}:${darkPageBorderCss(m[2])}`)
  }
  return out
}

/** side name → dkBorder side key */
export const DK_SIDE: Record<'top' | 'right' | 'bottom' | 'left', DkBorderSide> = {
  top: 't',
  right: 'r',
  bottom: 'b',
  left: 'l',
}

// ---- imperative-DOM counterparts (header/footer clones, footnote rows) --------

export function setDkColor(el: HTMLElement, hex: string): void {
  el.style.setProperty('--dk-c', darkPageColor(hex))
}

export function setDkBackground(el: HTMLElement, css: string): void {
  el.style.setProperty('--dk-bg', darkPageColor(css))
}

export function setDkBorder(el: HTMLElement, side: DkBorderSide, borderCss: string): void {
  el.style.setProperty(`--dk-b-${side}`, darkPageBorderCss(borderCss))
}

/** React style-object twins (HeaderFooterArea): spread into the element style */
export function dkStyleProps(twins: {
  color?: string
  background?: string
  borders?: Partial<Record<DkBorderSide, string>>
}): Record<string, string> {
  const out: Record<string, string> = {}
  if (twins.color) out['--dk-c'] = darkPageColor(twins.color)
  if (twins.background) out['--dk-bg'] = darkPageColor(twins.background)
  for (const [side, css] of Object.entries(twins.borders ?? {})) {
    if (css) out[`--dk-b-${side}`] = darkPageBorderCss(css)
  }
  return out
}
