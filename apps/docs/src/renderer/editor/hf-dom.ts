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
  return [p.runs, ...(p.cells?.map((c) => c.runs) ?? [])].some((rs) =>
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
    .map((p) => [...p.runs, ...(p.cells?.flatMap((c) => c.runs) ?? [])].map((r) => r.text).join(''))
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
        if (cell.align) {
          cellEl.style.textAlign =
            cell.align === 'left' || cell.align === 'center' || cell.align === 'right'
              ? cell.align
              : 'justify'
        }
        for (const run of cell.runs) {
          const span = document.createElement('span')
          span.textContent = display(run.text)
          applyRunStyle(span, run)
          cellEl.append(span)
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
