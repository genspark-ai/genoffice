import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfImage,
  type HfParagraph,
  type Run,
} from '@genoffice/docx-engine'
import { cssDualFontFamily, cssFontFamily } from '../line-metrics'

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
  if (run.italic) span.style.fontStyle = 'italic'
  const deco = [run.underline && 'underline', run.strike && 'line-through'].filter(Boolean)
  if (deco.length > 0) span.style.textDecoration = deco.join(' ')
  if (run.color) span.style.color = `#${run.color}`
  if (run.sizeHalfPoints) span.style.fontSize = `${run.sizeHalfPoints / 2}pt`
  if (run.font && run.fontAscii) span.style.fontFamily = cssDualFontFamily(run.fontAscii, run.font)
  else if (run.font || run.fontAscii)
    span.style.fontFamily = cssFontFamily((run.font ?? run.fontAscii)!)
  if (run.caps === 'all') span.style.textTransform = 'uppercase'
  else if (run.caps === 'small') span.style.fontVariantCaps = 'small-caps'
}

/** one tab-delimited chunk of a paragraph: the runs after the k-th tab, placed at stop k-1 */
export interface HfTabSegment {
  runs: Run[]
  /** stop offset from the text-column left edge; pct = implicit Word stop, px = explicit w:tabs */
  left: { px: number } | { pct: number }
  anchor: 'left' | 'center' | 'right'
}

const TWIPS_PER_PX = 15

/**
 * Tab layout of a header/footer paragraph: Word's three-column headers are
 * "left\tcenter\tright" with stops from w:tabs (usually on the Header/Footer
 * style). Returns null when the paragraph has no tab; otherwise the leading
 * inline runs plus one positioned segment per tab. Without explicit stops the
 * implicit header stops apply (center at half the column, right at its edge).
 */
export function hfTabSegments(
  para: HfParagraph,
): { lead: Run[]; segments: HfTabSegment[]; minHeightPt?: number } | null {
  if (!para.runs.some((r) => r.text.includes('\t'))) return null
  const chunks: Run[][] = [[]]
  for (const run of para.runs) {
    const pieces = run.text.split('\t')
    chunks[chunks.length - 1].push({ ...run, text: pieces[0] })
    for (const piece of pieces.slice(1)) chunks.push([{ ...run, text: piece }])
  }
  const stops = (para.tabStops ?? []).filter(
    (s) => s.val === 'left' || s.val === 'center' || s.val === 'right' || s.val === 'decimal',
  )
  const segments: HfTabSegment[] = []
  for (let k = 1; k < chunks.length; k++) {
    const runs = chunks[k].filter((r) => r.text !== '')
    // w:ptab carries its own margin-relative alignment and ignores tab stops
    const ptab = para.ptabAligns?.[k - 1]
    const stop = ptab ? undefined : stops[k - 1]
    if (ptab) {
      segments.push({
        runs,
        left: { pct: ptab === 'center' ? 50 : ptab === 'right' ? 100 : 0 },
        anchor: ptab,
      })
    } else if (stop) {
      segments.push({
        runs,
        left: { px: stop.pos / TWIPS_PER_PX },
        anchor: stop.val === 'center' ? 'center' : stop.val === 'left' ? 'left' : 'right',
      })
    } else {
      // implicit Word header/footer stops: first tab to the column center, next to its right edge
      segments.push(
        k === 1
          ? { runs, left: { pct: 50 }, anchor: 'center' }
          : { runs, left: { pct: 100 }, anchor: 'right' },
      )
    }
  }
  // absolutely positioned segments add no flow height: an oversized run after a
  // tab (or an empty lead) would collapse to the strip's min-height and clip
  const maxHalfPoints = Math.max(0, ...para.runs.map((r) => r.sizeHalfPoints ?? 0))
  return {
    lead: chunks[0].filter((r) => r.text !== ''),
    segments,
    ...(maxHalfPoints > 0 ? { minHeightPt: (maxHalfPoints / 2) * 1.3 } : {}),
  }
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

/** page (paper) geometry a floating header image positions against, px */
export interface HfFloatBox {
  pageW: number
  pageH: number
  marginLeft: number
  marginRight: number
  /** effective top margin after header push-down (mirrors the strip layout) */
  marginTop: number
  marginBottom: number
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
    y = img.posVRel === 'page' ? img.posYPx : box.marginTop + img.posYPx
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
export function makeHfFloatImgEl(img: HfImage, box: HfFloatBox, host: 'gap' | 'lead'): HTMLElement {
  const el = document.createElement('img')
  el.className = 'page-hf-float-img'
  el.src = img.dataUrl
  el.alt = ''
  el.draggable = false
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
  if (img.washout) el.style.filter = 'brightness(1.6) contrast(0.35)'
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
  if (images && images.length > 0) {
    const imgWrap = document.createElement('div')
    imgWrap.className = 'page-hf-images'
    if (images[0].align === 'right') imgWrap.style.justifyContent = 'flex-end'
    else if (images[0].align === 'center') imgWrap.style.justifyContent = 'center'
    for (const img of images) {
      const el = document.createElement('img')
      el.src = img.dataUrl
      el.alt = ''
      el.draggable = false
      if (img.widthPx) el.style.width = `${img.widthPx}px`
      if (img.heightPx) el.style.height = `${img.heightPx}px`
      imgWrap.append(el)
    }
    wrap.append(imgWrap)
  }
  for (const para of parasOf(value)) {
    const p = document.createElement('div')
    p.className = 'page-hf-para'
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
              const im = document.createElement('img')
              im.className = 'page-hf-cell-img'
              im.src = run.image.dataUrl
              im.alt = ''
              im.draggable = false
              if (run.image.widthPx) im.style.width = `${run.image.widthPx}px`
              if (run.image.heightPx) im.style.height = `${run.image.heightPx}px`
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
    const tabbed = hfTabSegments(para)
    if (tabbed) {
      p.classList.add('page-hf-tabbed')
      if (tabbed.minHeightPt) p.style.minHeight = `${tabbed.minHeightPt}pt`
      for (const run of tabbed.lead) {
        const span = document.createElement('span')
        span.textContent = display(run.text)
        applyRunStyle(span, run)
        p.append(span)
      }
      for (const seg of tabbed.segments) {
        const segEl = document.createElement('span')
        segEl.className = `page-hf-tabseg page-hf-tabseg-${seg.anchor}`
        segEl.style.left = 'px' in seg.left ? `${seg.left.px}px` : `${seg.left.pct}%`
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
