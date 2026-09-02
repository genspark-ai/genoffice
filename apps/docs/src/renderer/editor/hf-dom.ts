import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfImage,
  type HfParagraph,
  type Run,
} from '@genoffice/docx-engine'
import { cssRunFontFamily, estimateHfHeight, runLetterSpacingCss } from '../line-metrics'

/**
 * Plain-DOM header/footer rendering for the canvas page gaps (M4 always-on
 * pagination). Mirrors HeaderFooterArea's read-only display: same classes,
 * same PAGE_MARK / NUMPAGES substitution — but built imperatively because page gaps
 * are ProseMirror widget decorations, not React children.
 */

/** w:pBdr line CSS: declared color/width when present, legacy 1px #444 otherwise (document content color) */
export function paraBorderCss(line?: { color?: string; szPt?: number }): string {
  const widthPx = line?.szPt ? Math.max(1, Math.round((line.szPt * 96) / 72)) : 1
  return `${widthPx}px solid #${line?.color ?? '444'}`
}

function applyRunStyle(span: HTMLElement, run: Run): void {
  if (run.bold) span.style.fontWeight = '600'
  else if (run.bold === false) span.style.fontWeight = 'normal'
  if (run.italic) span.style.fontStyle = 'italic'
  else if (run.italic === false) span.style.fontStyle = 'normal'
  const deco = [run.underline && 'underline', run.strike && 'line-through'].filter(Boolean)
  if (deco.length > 0) span.style.textDecoration = deco.join(' ')
  if (run.color) span.style.color = `#${run.color}`
  if (run.sizeHalfPoints) span.style.fontSize = `${run.sizeHalfPoints / 2}pt`
  const letterSpacing = runLetterSpacingCss(run)
  if (letterSpacing) span.style.letterSpacing = letterSpacing
  if (run.font || run.fontAscii) span.style.fontFamily = cssRunFontFamily(run.fontAscii, run.font)
  if (run.caps === 'all') span.style.textTransform = 'uppercase'
  else if (run.caps === 'small') span.style.fontVariantCaps = 'small-caps'
  else if (run.caps === 'none') {
    span.style.textTransform = 'none'
    span.style.fontVariantCaps = 'normal'
  }
}

/** one tab-delimited chunk of a paragraph: the runs after the k-th tab, laid out at its stop */
export interface HfTabSegment {
  runs: Run[]
  /** offset from the text-column left edge; pct = margin-relative (w:ptab / implicit stops) */
  left: { px: number } | { pct: number }
  anchor: 'left' | 'center' | 'right'
}

export interface HfTabLayout {
  lead: Run[]
  segments: HfTabSegment[]
  minHeightPt?: number
  /** w:jc of the tab-laid line: Word lays tabs out in left-aligned space, then
   *  shifts the whole line (right: line end at the column edge; center: line
   *  centered). lineEndPx = laid-out line end in column space. */
  shift?: { align: 'center' | 'right'; lineEndPx: number }
}

const TWIPS_PER_PX = 15
/** default tab grid past the last explicit stop (720 twips, matches Word) */
const HF_DEFAULT_GRID_PX = 48
const HF_DEFAULT_FONT_PT = 10.5

/** effective strip base size (pt): the mounted --hf-default-fs (docStyleCss,
 *  shrink-only) when present, so tab measures match what the strip paints */
function hfBaseFontPt(): number {
  if (typeof document === 'undefined') return HF_DEFAULT_FONT_PT
  const host = document.querySelector('.page-wrap, .doc-page, .pv-page')
  if (!host) return HF_DEFAULT_FONT_PT
  const pt = parseFloat(getComputedStyle(host).getPropertyValue('--hf-default-fs'))
  return Number.isFinite(pt) && pt > 0 ? pt : HF_DEFAULT_FONT_PT
}

let hfMeasureCtx: CanvasRenderingContext2D | null | undefined
/** approximate rendered width of hf runs (px): canvas mirror of applyRunStyle's font mapping.
 *  `display` is the caller's field substitution (PAGE_MARK -> the page digits): the raw
 *  private-use marks measure as notdef boxes and would skew stop placement. */
function hfRunsWidthPx(runs: Run[], display: (text: string) => string): number {
  if (hfMeasureCtx === undefined) {
    hfMeasureCtx =
      typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null
  }
  let w = 0
  const basePt = hfBaseFontPt()
  for (const run of runs) {
    if (run.image?.widthPx) w += run.image.widthPx
    if (!run.text) continue
    const shown = display(run.text)
    const text = run.caps === 'all' ? shown.toUpperCase() : shown
    const sizePt = run.sizeHalfPoints ? run.sizeHalfPoints / 2 : basePt
    if (hfMeasureCtx) {
      const family =
        run.font || run.fontAscii ? cssRunFontFamily(run.fontAscii, run.font) : 'sans-serif'
      // Detached-canvas font parsing cannot resolve var() (no element style context —
      // the whole assignment would be ignored): inline each var's fallback chain
      const canvasFamily = family.replace(/var\(--[\w-]+,([^)]*)\)/g, '$1')
      hfMeasureCtx.font = `${run.italic ? 'italic ' : ''}${run.bold ? '600 ' : ''}${(sizePt * 4) / 3}px ${canvasFamily}`
      w += hfMeasureCtx.measureText(text).width
    } else {
      w += text.length * sizePt * 0.5 * (4 / 3)
    }
    if (run.charSpacingTwips) w += (((run.charSpacingTwips / 20) * 4) / 3) * text.length
  }
  return w
}

/**
 * Tab layout of a header/footer paragraph, Word semantics: style-chain and
 * direct stops are pre-merged by the parser; each tab advances from the
 * current position to the next stop past it (center/right stops align their
 * segment at the stop), the default grid takes over past the last stop, and
 * w:jc then shifts the whole laid-out line. Returns null when the paragraph
 * has no tab. Without any stops the implicit header stops apply (center at
 * half the column, then its right edge).
 */
export function hfTabSegments(
  para: HfParagraph,
  display: (text: string) => string = (t) => t,
): HfTabLayout | null {
  if (!para.runs.some((r) => r.text.includes('\t'))) return null
  const chunks: Run[][] = [[]]
  for (const run of para.runs) {
    const pieces = run.text.split('\t')
    chunks[chunks.length - 1].push({ ...run, text: pieces[0] })
    for (const piece of pieces.slice(1)) chunks.push([{ ...run, text: piece }])
  }
  const stops = (para.tabStops ?? [])
    .filter(
      (s) =>
        (s.val === 'left' || s.val === 'center' || s.val === 'right' || s.val === 'decimal') &&
        Number.isFinite(s.pos),
    )
    .map((s) => ({ x: s.pos / TWIPS_PER_PX, val: s.val }))
    .sort((a, b) => a.x - b.x)

  const segments: HfTabSegment[] = []
  let usedPct = false
  let x = hfRunsWidthPx(chunks[0], display)
  for (let k = 1; k < chunks.length; k++) {
    const runs = chunks[k].filter((r) => r.text !== '')
    // w:ptab carries its own margin-relative alignment and ignores tab stops
    const ptab = para.ptabAligns?.[k - 1]
    if (ptab) {
      usedPct = true
      segments.push({
        runs,
        left: { pct: ptab === 'center' ? 50 : ptab === 'right' ? 100 : 0 },
        anchor: ptab,
      })
      continue
    }
    if (stops.length === 0) {
      // implicit Word header/footer stops: first tab to the column center, next to its right edge
      usedPct = true
      segments.push(
        k === 1
          ? { runs, left: { pct: 50 }, anchor: 'center' }
          : { runs, left: { pct: 100 }, anchor: 'right' },
      )
      continue
    }
    const segW = hfRunsWidthPx(runs, display)
    const stop = stops.find((s) => s.x > x + 0.5)
    const target = stop
      ? stop.x
      : (Math.floor((x + 0.5) / HF_DEFAULT_GRID_PX) + 1) * HF_DEFAULT_GRID_PX
    const val = stop?.val ?? 'left'
    const placed = Math.max(
      x,
      val === 'center' ? target - segW / 2 : val === 'left' ? target : target - segW,
    )
    if (runs.length > 0) segments.push({ runs, left: { px: placed }, anchor: 'left' })
    x = placed + segW
  }
  // absolutely positioned segments add no flow height: an oversized run after a
  // tab (or an empty lead) would collapse to the strip's min-height and clip
  const maxHalfPoints = Math.max(0, ...para.runs.map((r) => r.sizeHalfPoints ?? 0))
  const align = para.align
  return {
    lead: chunks[0].filter((r) => r.text !== ''),
    segments,
    ...(maxHalfPoints > 0 ? { minHeightPt: (maxHalfPoints / 2) * 1.3 } : {}),
    ...(!usedPct && (align === 'center' || align === 'right')
      ? { shift: { align, lineEndPx: x } }
      : {}),
  }
}

/** CSS left of a tab segment, including the line's w:jc shift (never left of its laid-out spot) */
export function hfSegLeftCss(seg: HfTabSegment, layout: HfTabLayout): string {
  if ('pct' in seg.left) return `${seg.left.pct}%`
  const px = seg.left.px
  const s = layout.shift
  if (!s) return `${px.toFixed(1)}px`
  return s.align === 'right'
    ? `max(${px.toFixed(1)}px, calc(100% - ${(s.lineEndPx - px).toFixed(1)}px))`
    : `max(${px.toFixed(1)}px, calc(50% + ${(px - s.lineEndPx / 2).toFixed(1)}px))`
}

/** text-indent that applies the line's w:jc shift to the in-flow lead runs */
export function hfLeadIndentCss(layout: HfTabLayout): string | null {
  const s = layout.shift
  if (!s) return null
  return s.align === 'right'
    ? `max(0px, calc(100% - ${s.lineEndPx.toFixed(1)}px))`
    : `max(0px, calc(50% - ${(s.lineEndPx / 2).toFixed(1)}px))`
}

/** effective paragraphs: rich paras when present, else the legacy single line (mirrors HeaderFooterArea) */
function parasOf(value: HeaderFooter): HfParagraph[] {
  if (value.paras?.length) return value.paras
  const runs: Run[] = value.text ? [{ text: value.text }] : []
  if (value.pageNumber && !value.text.includes('#') && !value.text.includes(PAGE_MARK)) {
    runs.push({ text: runs.length > 0 ? ` ${PAGE_MARK}` : PAGE_MARK })
  }
  return [{ align: 'center', runs }]
}

/**
 * Largest declared run size (pt) when every text-bearing run of the strip
 * declares one, else null (some run inherits the strip base). Whitespace-only
 * runs don't size lines in Word and only count when nothing else does (an
 * empty paragraph is sized by its mark).
 */
export function hfDeclaredStrutPt(paras: HfParagraph[]): number | null {
  let text: number | null = null
  let blank: number | null = null
  const scan = (runs: Run[]): boolean => {
    for (const run of runs) {
      if (!run.text && run.image) continue
      const sz = run.sizeHalfPoints
      if (!(run.text ?? '').trim()) {
        if (sz) blank = Math.max(blank ?? 0, sz)
        continue
      }
      if (!sz) return false
      text = Math.max(text ?? 0, sz)
    }
    return true
  }
  for (const p of paras) {
    if (p.boxAnchored) continue
    if (p.cells) {
      for (const c of p.cells) for (const runs of c.paras) if (!scan(runs)) return null
    } else if (!scan(p.runs)) return null
  }
  const half = text ?? blank
  return half ? half / 2 : null
}

/** parsed parts carry PAGE fields as PAGE_MARK; only mark-free values fall back to the user-typed '#' (mirrors HeaderFooterArea) */
function paraHasPageMark(p: HfParagraph): boolean {
  return [p.runs, ...(p.cells?.flatMap((c) => c.paras) ?? [])].some((rs) =>
    rs.some((r) => r.text.includes(PAGE_MARK)),
  )
}

export function hfUsesLegacyHash(value: HeaderFooter): boolean {
  if (!value.pageNumber) return false
  if (value.text.includes(PAGE_MARK)) return false
  return !value.paras?.some(paraHasPageMark)
}

export function hfHasPageField(value: HeaderFooter | null | undefined): boolean {
  return Boolean(
    value &&
    (value.pageNumber || value.text.includes(PAGE_MARK) || value.paras?.some(paraHasPageMark)),
  )
}

/** Remove Page Numbers: strip PAGE_MARK fields (legacy parts: only the first user-typed '#'), keeping literal '#' text and rich formatting.
 *  Layout-table rows (`cells`) pass through untouched: they are display-only and saving keeps
 *  the part's original w:tbl bytes, so a table-held PAGE field cannot be removed. */
export function hfWithoutPageMarks(value: HeaderFooter): HeaderFooter {
  let legacyHash = hfUsesLegacyHash(value)
  const strip = (t: string) => {
    const out = t.replaceAll(PAGE_MARK, '')
    if (legacyHash && out.includes('#')) {
      legacyHash = false
      return out.replace('#', '')
    }
    return out
  }
  if (!value.paras?.length) return { ...value, text: strip(value.text), pageNumber: false }
  const paras = value.paras
    .map((p) =>
      p.cells
        ? p
        : { ...p, runs: p.runs.map((r) => ({ ...r, text: strip(r.text) })).filter((r) => r.text) },
    )
    // dedicated page-number paragraphs go away entirely; table rows and user-typed blank lines stay
    .filter((p, i) => p.cells != null || p.runs.length > 0 || value.paras![i].runs.length === 0)
  const text = paras
    .map((p) =>
      [...p.runs, ...(p.cells?.flatMap((c) => c.paras.flat()) ?? [])].map((r) => r.text).join(''),
    )
    .join('')
  if (!text) return { ...value, text: '', paras: undefined, pageNumber: false }
  return { ...value, text, paras, pageNumber: false }
}

export function hfHasVisibleContent(
  value: HeaderFooter | null | undefined,
  images?: HfImage[],
): boolean {
  if (images?.length) return true
  if (!value) return false
  return Boolean(
    value.text ||
    value.pageNumber ||
    value.paras?.some((p) => p.runs.length > 0 || p.cells?.length),
  )
}

/** rendered strip heights keyed by content signature (pagination recomputes per page) */
const hfHeightCache = new Map<string, number>()

function hfProbeHost(): HTMLElement | null {
  if (typeof document === 'undefined' || !document.body) return null
  let host = document.getElementById('hf-strip-probe')
  if (!host) {
    host = document.createElement('div')
    host.id = 'hf-strip-probe'
    // .doc-page: the probe inherits the same document-default font/line-height
    // CSS the real gap strips get inside the editor root
    host.className = 'doc-page'
    host.style.cssText =
      'position:absolute;left:-99999px;top:0;visibility:hidden;pointer-events:none'
    // editor-chrome floors (dashed separator, clickable-strip min-height) are
    // not Word geometry: with them the measure over-reserves ~1 line per strip
    // paragraph, costing body lines on every page of every footered document.
    // In document.head because the probe host replaces its children per measure
    const neutralize = document.createElement('style')
    neutralize.textContent =
      '#hf-strip-probe .page-hf { padding: 0; border: none; }' +
      '#hf-strip-probe .page-hf-para { min-height: 0; }' +
      // floating-box content draws at its anchor, not in the strip flow
      '#hf-strip-probe .page-hf-box-anchored { display: none; }'
    document.head.appendChild(neutralize)
    document.body.appendChild(host)
  }
  return host
}

/** djb2 over the mounted doc-scoped stylesheets (memoized on the concatenated string) */
let mountedCssMemo: { text: string; hash: string } | null = null
function mountedDocCssSig(): string {
  const text = Array.from(document.querySelectorAll('style[data-doc-css]'))
    .map((s) => s.textContent ?? '')
    .join(' ')
  if (mountedCssMemo?.text === text) return mountedCssMemo.hash
  let h = 5381
  for (let i = 0; i < text.length; i++) h = ((h << 5) + h + text.charCodeAt(i)) | 0
  const hash = `${text.length}:${h}`
  mountedCssMemo = { text, hash }
  return hash
}

/**
 * Reserved height (px) of a header/footer strip for body push-down. The
 * line-box estimate decides alone outside a DOM (tests); in the renderer the
 * strip is additionally laid out in a hidden .doc-page probe — the same
 * makeGapHfEl markup and CSS the canvas draws — so real wraps, borders and
 * line heights count when font substitution renders taller than the estimate
 * (the DOM value only ever raises the estimate: floating-image reservations
 * are estimate-only and must not drop). The probe measures under whatever
 * doc styles are mounted right now and the cache keys on their content, so a
 * value computed before a new document's CSS commits is re-measured on the
 * caller's post-commit pass (App's hfMeasureEpoch) instead of going stale.
 */
export function hfReservedHeightPx(
  kind: 'header' | 'footer',
  value: HeaderFooter | null,
  contentWidthPx: number,
  images?: HfImage[] | null,
  geom?: { marginTopPx: number; headerDistPx: number },
): number {
  const est = estimateHfHeight(value, contentWidthPx, images, geom)
  const inline = (images ?? []).filter((img) => !img.floating)
  if (!hfHasVisibleContent(value, inline) || contentWidthPx <= 0) return est
  if (typeof document === 'undefined' || !document.body) return est
  // cell-run images can carry megabytes of base64: key on the dataUrl length only
  const noDataUrl = (k: string, v: unknown) => (k === 'dataUrl' ? String(v).length : v)
  const key =
    `${kind}|${Math.round(contentWidthPx)}|${mountedDocCssSig()}|` +
    `${JSON.stringify(value, noDataUrl)}|` +
    inline.map((im) => `${im.widthPx ?? ''}x${im.heightPx ?? ''}:${im.dataUrl.length}`).join(',')
  let dom = hfHeightCache.get(key)
  if (dom === undefined) {
    const host = hfProbeHost()
    if (!host) return est
    const el = makeGapHfEl({
      kind,
      value: value ?? { text: '' },
      images: inline,
      pageNo: 88,
      pageTotal: 88,
    })
    el.style.position = 'static'
    el.style.left = 'auto'
    el.style.transform = 'none'
    el.style.width = `${contentWidthPx}px`
    // an empty paragraph still occupies its line in Word; with the chrome
    // min-height neutralized it needs a real line box to measure that
    for (const p of el.querySelectorAll('.page-hf-para')) {
      if (!(p.textContent ?? '').length && p.childElementCount === 0) {
        p.appendChild(document.createTextNode('\u200b'))
      }
    }
    host.replaceChildren(el)
    dom = el.getBoundingClientRect().height
    host.replaceChildren()
    if (hfHeightCache.size > 300) hfHeightCache.clear()
    hfHeightCache.set(key, dom)
  }
  return dom > 0 ? Math.max(est, dom) : est
}

/** page (paper) geometry a floating header image positions against, px */
export interface HfFloatBox {
  pageW: number
  pageH: number
  marginLeft: number
  marginRight: number
  /** effective top margin after header push-down (mirrors the strip layout) */
  marginTop: number
  marginBottom: number
  /** header strip top (w:headerDist, px): origin of paragraph-relative vertical offsets */
  headerDist: number
  /** raw sectPr top margin (px, before push-down): origin wrapped margin-relative images reserve from */
  sectMarginTop: number
}

/**
 * Position of a floating header image in page coordinates (px from the page's
 * top-left corner), with the translate that resolves center/right/bottom
 * anchors without knowing the image's natural size. wp:posOffset offsets win;
 * alignment fields reproduce the legacy VML placement (margin-box corners).
 */
export function hfFloatPagePos(
  img: HfImage,
  box: HfFloatBox,
): { x: number; y: number; translateX: 0 | -50 | -100; translateY: 0 | -50 | -100 } {
  let x: number
  let translateX: 0 | -50 | -100 = 0
  if (img.posXPx != null) {
    x = img.posHRel === 'page' ? img.posXPx : box.marginLeft + img.posXPx
  } else if (img.posH === 'center') {
    x = box.pageW / 2
    translateX = -50
  } else if (img.posH === 'right') {
    x = box.pageW - box.marginRight
    translateX = -100
  } else {
    x = box.marginLeft
  }
  let y: number
  let translateY: 0 | -50 | -100 = 0
  if (img.posYPx != null) {
    // paragraph/margin-rel wrapped images use the same origins the body
    // push-down estimate measures from (header strip top / raw sectPr margin),
    // so the image never chases the pushed-down margin; watermarks keep the
    // effective margin
    const wrapped = img.wrap && img.wrap !== 'none' && !img.behind
    y =
      img.posVRel === 'page'
        ? img.posYPx
        : img.posVRel === 'paragraph'
          ? box.headerDist + img.posYPx
          : (wrapped ? box.sectMarginTop : box.marginTop) + img.posYPx
  } else if (img.posV === 'center') {
    y = box.pageH / 2
    translateY = -50
  } else if (img.posV === 'bottom') {
    y = box.pageH - box.marginBottom
    translateY = -100
  } else {
    y = box.marginTop
  }
  return { x, y, translateX, translateY }
}

/**
 * Floating header image (picture watermark) for the canvas print view, drawn
 * once per page behind the body text (z-index -1; .view-print .doc-page
 * isolates). host 'gap': anchored in a page gap, whose bottom edge sits
 * marginTop above the next page's content. host 'lead': anchored in the
 * zero-height first-page widget at the content-box origin.
 */

/**
 * Word's washout preset (gain 19661f, blacklevel 22938f) is the linear map
 * out = 0.3*in + 0.7 per channel (MS-ODRAW brightness/contrast), so white
 * stays white and the image fades toward white. invert-brightness-invert
 * reproduces it exactly: 1 - 0.3*(1 - in); brightness/contrast chains clamp
 * at white and repaint white pixels as grey.
 */
export const HF_WASHOUT_FILTER = 'invert(1) brightness(0.3) invert(1)'
/** <img> for a header/footer picture; an a:srcRect crop becomes an
 *  overflow-hidden window over a scaled and offset image (body-path technique) */
function hfImgNode(img: {
  dataUrl: string
  widthPx?: number
  heightPx?: number
  crop?: HfImage['crop']
}): HTMLElement {
  const im = document.createElement('img')
  im.src = img.dataUrl
  im.alt = ''
  im.draggable = false
  const c = img.crop
  if (c && img.widthPx && img.heightPx) {
    const span = (a: number, b: number) => Math.max(0.01, 1 - a - b)
    const sw = img.widthPx / span(c.l, c.r)
    const sh = img.heightPx / span(c.t, c.b)
    const win = document.createElement('span')
    win.style.display = 'inline-block'
    win.style.overflow = 'hidden'
    win.style.width = `${img.widthPx}px`
    win.style.height = `${img.heightPx}px`
    im.style.cssText =
      `position:absolute;left:${(-c.l * sw).toFixed(1)}px;top:${(-c.t * sh).toFixed(1)}px;` +
      `width:${sw.toFixed(1)}px;height:${sh.toFixed(1)}px;max-width:none`
    // the window anchors the absolute img even when the caller's class does
    // not position the wrapper itself
    win.style.position = 'relative'
    win.append(im)
    return win
  }
  if (img.widthPx) im.style.width = `${img.widthPx}px`
  if (img.heightPx) im.style.height = `${img.heightPx}px`
  return im
}

export function makeHfFloatImgEl(img: HfImage, box: HfFloatBox, host: 'gap' | 'lead'): HTMLElement {
  const el = hfImgNode(img)
  el.className = 'page-hf-float-img'
  // a crop wrapper carries an inline position:relative for the strip hosts;
  // the float path must stay absolute (inline style beats the class)
  el.style.position = 'absolute'
  const p = hfFloatPagePos(img, box)
  if (host === 'gap') {
    el.style.left = `${p.x}px`
    el.style.top = `calc(100% + ${p.y - box.marginTop}px)`
  } else {
    el.style.left = `${p.x - box.marginLeft}px`
    el.style.top = `${p.y - box.marginTop}px`
  }
  if (p.translateX || p.translateY) {
    el.style.transform = `translate(${p.translateX}%, ${p.translateY}%)`
  }
  if (img.widthPx) el.style.width = `${img.widthPx}px`
  if (img.heightPx) el.style.height = `${img.heightPx}px`
  if (img.washout) el.style.filter = HF_WASHOUT_FILTER
  return el
}

export function makeGapHfEl(opts: {
  kind: 'header' | 'footer'
  value: HeaderFooter
  images?: HfImage[]
  /** page number shown for the PAGE marker (may be a section-formatted string) */
  pageNo: number | string
  /** total page count shown for the NUMPAGES marker */
  pageTotal: number
}): HTMLElement {
  const { kind, value, images, pageNo, pageTotal } = opts
  const legacyHash = hfUsesLegacyHash(value)
  const display = (text: string) => {
    const substituted = text
      .replaceAll(TOTAL_PAGES_MARK, String(pageTotal))
      .replaceAll(PAGE_MARK, String(pageNo))
    return legacyHash ? substituted.replace('#', String(pageNo)) : substituted
  }
  const wrap = document.createElement('div')
  wrap.className = `page-hf page-hf-${kind} page-gap-hf`
  wrap.contentEditable = 'false'
  // Word sizes strip lines by their runs — the strip base is only the fallback
  // for runs without w:sz. When every text run declares a size, the strip font
  // (the em behind every strut/empty-line line-height) takes it, shrink-only
  // (growth would over-reserve push-down, prod_008/091). SAS prod_043: an
  // all-8pt header measured at the 10.5pt strut pushed the body top ~1px and
  // cost every two-column page its 42nd grid row (+1 page).
  const strutPt = hfDeclaredStrutPt(parasOf(value))
  if (strutPt != null) wrap.style.fontSize = `min(${strutPt}pt, var(--hf-default-fs, 10.5pt))`
  if (images && images.length > 0) {
    const imgWrap = document.createElement('div')
    imgWrap.className = 'page-hf-images'
    if (images[0].align === 'right') imgWrap.style.justifyContent = 'flex-end'
    else if (images[0].align === 'center') imgWrap.style.justifyContent = 'center'
    for (const img of images) {
      imgWrap.append(hfImgNode(img))
    }
    wrap.append(imgWrap)
  }
  for (const para of parasOf(value)) {
    const p = document.createElement('div')
    p.className = 'page-hf-para'
    // floating-box content: displayed in the strip, but the push-down probe
    // excludes it (Word draws it at the anchor, off the strip flow)
    if (para.boxAnchored) p.classList.add('page-hf-box-anchored')
    if (para.cells) {
      // layout-table row: flex columns sized by the cell widths
      p.classList.add('page-hf-row')
      for (const cell of para.cells) {
        const cellEl = document.createElement('div')
        cellEl.className = 'page-hf-cell'
        if (cell.widthPct) cellEl.style.width = `${cell.widthPct}%`
        /* document content color (w:shd), theme-independent */
        if (cell.fill) cellEl.style.backgroundColor = `#${cell.fill}`
        if (cell.align) {
          cellEl.style.textAlign =
            cell.align === 'left' || cell.align === 'center' || cell.align === 'right'
              ? cell.align
              : 'justify'
        }
        // one block line per cell paragraph (Word stacks them; a lone empty
        // paragraph still reserves its line inside a shaded cell)
        for (const runs of cell.paras.length > 0 ? cell.paras : [[]]) {
          const paraEl = document.createElement('div')
          paraEl.className = 'page-hf-cell-para'
          if (runs.length === 0) paraEl.textContent = ' '
          for (const run of runs) {
            if (run.image) {
              const im = hfImgNode(run.image)
              im.classList.add('page-hf-cell-img')
              paraEl.append(im)
              if (!run.text) continue
            }
            const span = document.createElement('span')
            span.textContent = display(run.text)
            applyRunStyle(span, run)
            paraEl.append(span)
          }
          cellEl.append(paraEl)
        }
        p.append(cellEl)
      }
      wrap.append(p)
      continue
    }
    if (para.bidi) p.style.direction = 'rtl'
    if (para.align) {
      p.style.textAlign =
        para.align === 'left' || para.align === 'center' || para.align === 'right'
          ? para.align
          : 'justify'
    }
    // frame placement wins over the paragraph's own jc (the frame is narrower
    // than the column; its x position is what the reader sees)
    if (para.frameXAlign) {
      p.classList.add('page-hf-frame')
      p.style.textAlign = para.frameXAlign
    }
    /* document content colors (w:shd / w:pBdr), theme-independent; mirrors the body paragraph path */
    if (para.shadingFill) p.style.backgroundColor = `#${para.shadingFill}`
    if (para.borders) {
      const line = (side: 't' | 'b' | 'l' | 'r') => paraBorderCss(para.borderLines?.[side])
      if (para.borders.includes('t')) p.style.borderTop = line('t')
      if (para.borders.includes('b')) p.style.borderBottom = line('b')
      if (para.borders.includes('l')) p.style.borderLeft = line('l')
      if (para.borders.includes('r')) p.style.borderRight = line('r')
      p.style.padding = '1px 4px'
    }
    const tabbed = hfTabSegments(para, display)
    if (tabbed) {
      p.classList.add('page-hf-tabbed')
      if (tabbed.minHeightPt) p.style.minHeight = `${tabbed.minHeightPt}pt`
      // tab layout happens in left-aligned space; w:jc becomes an explicit shift
      p.style.textAlign = 'left'
      const leadIndent = hfLeadIndentCss(tabbed)
      if (leadIndent) p.style.textIndent = leadIndent
      for (const run of tabbed.lead) {
        const span = document.createElement('span')
        span.textContent = display(run.text)
        applyRunStyle(span, run)
        p.append(span)
      }
      for (const seg of tabbed.segments) {
        const segEl = document.createElement('span')
        segEl.className = `page-hf-tabseg page-hf-tabseg-${seg.anchor}`
        segEl.style.left = hfSegLeftCss(seg, tabbed)
        for (const run of seg.runs) {
          const span = document.createElement('span')
          span.textContent = display(run.text)
          applyRunStyle(span, run)
          segEl.append(span)
        }
        p.append(segEl)
      }
      wrap.append(p)
      continue
    }
    if (para.runs.length === 0) p.textContent = ' '
    for (const run of para.runs) {
      const span = document.createElement('span')
      span.textContent = display(run.text)
      applyRunStyle(span, run)
      p.append(span)
    }
    wrap.append(p)
  }
  return wrap
}
