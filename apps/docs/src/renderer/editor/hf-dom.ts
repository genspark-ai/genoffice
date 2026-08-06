import {
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
 * same '#' / NUMPAGES substitution — but built imperatively because page gaps
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
  if (value.pageNumber && !value.text.includes('#')) {
    runs.push({ text: runs.length > 0 ? ' #' : '#' })
  }
  return [{ align: 'center', runs }]
}

export function hfHasVisibleContent(
  value: HeaderFooter | null | undefined,
  images?: HfImage[],
): boolean {
  if (images?.length) return true
  if (!value) return false
  return Boolean(value.text || value.pageNumber || value.paras?.some((p) => p.runs.length > 0))
}

export function makeGapHfEl(opts: {
  kind: 'header' | 'footer'
  value: HeaderFooter
  images?: HfImage[]
  /** page number shown for the '#' marker (may be a section-formatted string) */
  pageNo: number | string
  /** total page count shown for the NUMPAGES marker */
  pageTotal: number
}): HTMLElement {
  const { kind, value, images, pageNo, pageTotal } = opts
  const display = (text: string) => {
    const substituted = text.replaceAll(TOTAL_PAGES_MARK, String(pageTotal))
    return value.pageNumber ? substituted.replace('#', String(pageNo)) : substituted
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
