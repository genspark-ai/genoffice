import type { DocGrid, NoteRun } from '@genoffice/docx-engine'
import type { PageNoteItem } from '../pagination-types'
import { noteMarkText, type NoteKind } from '../note-format'
import { t } from '../i18n/locale'
import { textColorValue } from './text-color'
import { textOutlineCssValue } from './text-outline'
import { setDkColor } from './dark-page'
import { cssFontFamily, noteLineHeightPx, type NoteStyleOpts } from '../line-metrics'
import { fromNoteRuns, toStyleRecord } from './para-layout-spec'

/** classes and inline style of one note row: the first paragraph's line box
 *  (para-layout-spec) under the reservation model's line height, size and face */
export function noteRowBox(
  item: Pick<PageNoteItem, 'richParas' | 'lineHeightPx' | 'fontSizePt' | 'fontFamily'>,
): { classes: string[]; style: Record<string, string> } {
  const box = fromNoteRuns(item.richParas?.[0] ?? [])
  const style = toStyleRecord(box)
  if (item.lineHeightPx) style.lineHeight = `${item.lineHeightPx}px`
  if (item.fontSizePt) style.fontSize = `${item.fontSizePt}pt`
  // the model's face (style chain, then the entry's runs) stands in for a strut face the runs do not settle
  if (item.fontFamily && !style.fontFamily) style.fontFamily = item.fontFamily
  return { classes: box.classes, style }
}

/** row metrics of one note entry: the reservation model's line height, the
 *  resolved base size and face (style chain, then the entry's own runs) */
export function noteRowMetrics(
  style: NoteStyleOpts,
  docGrid: DocGrid | undefined,
): { lineHeightPx: number; fontSizePt: number; fontFamily?: string } {
  const fontFamily = style.fontFamily ? cssFontFamily(style.fontFamily) : undefined
  return {
    lineHeightPx: noteLineHeightPx(docGrid, style),
    fontSizePt: style.sizeHalfPoints ? style.sizeHalfPoints / 2 : 10,
    ...(fontFamily ? { fontFamily } : {}),
  }
}

/** Note entry content: superscript number + styled runs (shared by the canvas gap area and the hidden height measurer) */
/** revision classes of a note run (the app root's rev-display-* class drives the view) */
export function noteRunClass(run: Pick<NoteRun, 'ins' | 'del'>): string | undefined {
  const cls = [run.ins && 'doc-ins', run.del && 'doc-del'].filter(Boolean).join(' ')
  return cls || undefined
}

export function appendNoteRuns(
  row: HTMLElement,
  it: Omit<PageNoteItem, 'height' | 'id'>,
  kind: NoteKind,
): void {
  if (!it.noRefMark) {
    const sup = document.createElement('sup')
    sup.textContent = noteMarkText(kind, it.no)
    row.append(sup)
  }
  if (it.richParas) {
    it.richParas.forEach((para, pi) => {
      if (pi > 0) row.append(document.createElement('br'))
      for (const run of para) {
        const span = document.createElement('span')
        span.textContent = run.text
        const revClass = noteRunClass(run)
        if (revClass) span.className = revClass
        if (run.bold) span.style.fontWeight = '600'
        if (run.italic) span.style.fontStyle = 'italic'
        const deco = [run.underline && 'underline', run.strike && 'line-through'].filter(Boolean)
        if (deco.length > 0) span.style.textDecoration = deco.join(' ')
        if (run.color) {
          span.style.color = textColorValue(run.color)
          // dark-page twin; the authored color stays the declaration
          if (run.color !== 'auto') setDkColor(span, run.color)
        }
        if (run.sizeHalfPoints) span.style.fontSize = `${run.sizeHalfPoints / 2}pt`
        if (run.fontAscii) span.style.fontFamily = cssFontFamily(run.fontAscii)
        if (run.textOutline) span.style.webkitTextStroke = textOutlineCssValue(run.textOutline)
        if (run.caps === 'all') span.style.textTransform = 'uppercase'
        else if (run.caps === 'small') span.style.fontVariantCaps = 'small-caps'
        row.append(span)
      }
    })
  } else {
    const span = document.createElement('span')
    span.textContent = it.text
    row.append(span)
  }
}

/**
 * Hidden DOM measurement of one note entry at the renderers' exact styles: the
 * char-width estimate undercounts complex scripts (a production document's Arabic notes
 * overprinted each other), so reservation and drawing share the DOM wrap truth.
 */
let noteMeasureHost: HTMLDivElement | null = null
export function measureNoteHeightDom(
  entry: Pick<PageNoteItem, 'no' | 'text' | 'richParas' | 'noRefMark'>,
  kind: NoteKind,
  widthPx: number,
  lineHeightPx: number,
  fontSizePt: number,
  fontFamily?: string,
): number | null {
  if (typeof document === 'undefined' || !document.body || widthPx <= 0) return null
  if (!noteMeasureHost || !noteMeasureHost.isConnected) {
    noteMeasureHost = document.createElement('div')
    noteMeasureHost.style.cssText =
      'position:absolute;left:-99999px;top:0;visibility:hidden;pointer-events:none;text-align:left;text-indent:0;'
    document.body.appendChild(noteMeasureHost)
  }
  noteMeasureHost.style.fontFamily =
    fontFamily ?? "Calibri, 'DengXian', 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif"
  noteMeasureHost.style.width = `${widthPx}px`
  noteMeasureHost.style.fontSize = `${fontSizePt}pt`
  noteMeasureHost.style.lineHeight = `${lineHeightPx}px`
  const row = document.createElement('div')
  // the same row box as the renderers (direction and wrapping decide the wrap)
  for (const [k, v] of Object.entries(noteRowBox({ richParas: entry.richParas }).style)) {
    if (k.startsWith('--')) row.style.setProperty(k, v)
    else row.style[k as 'direction'] = v
  }
  appendNoteRuns(row, entry, kind)
  // the renderers' sup style (CSS classes don't reach the detached host)
  const sup = row.querySelector('sup')
  if (sup) {
    sup.style.fontSize = '0.65em'
    sup.style.marginRight = '4px'
  }
  noteMeasureHost.replaceChildren(row)
  const h = row.getBoundingClientRect().height
  noteMeasureHost.replaceChildren()
  return h > 0 ? h : null
}

/** Footnote area at the top of a page gap (previous page's bottom): absolutely positioned in the content area, double-click an entry to edit */
export function makeGapNotesEl(
  items: PageNoteItem[],
  leftPx: number,
  widthPx: number,
  heightPx: number,
  lineHeightPx: number,
  onEdit: (id: string) => void,
): HTMLElement {
  const wrap = document.createElement('div')
  wrap.className = 'page-gap-notes'
  wrap.style.left = `${leftPx}px`
  // paper x: block gaps start at the paper edge, mid-paragraph gaps at their page's
  // edge — alignGapHfStrips re-anchors from the actual gap origin
  wrap.dataset.paperX = leftPx.toFixed(1)
  wrap.style.width = `${widthPx}px`
  wrap.style.height = `${heightPx}px`
  // same line height as the reservation model; min-height keeps rows at the
  // reserved size while letting long entries overflow instead of clipping
  wrap.style.lineHeight = `${lineHeightPx}px`
  for (const it of items) {
    const row = document.createElement('div')
    const box = noteRowBox(it)
    row.className = ['page-gap-note', ...box.classes].join(' ')
    row.title = t('appDblclickEditFootnote')
    row.dataset.noteId = it.id
    row.style.minHeight = `${it.height}px`
    // per-note resolved style: the entry's line boxes must match its reservation
    for (const [k, v] of Object.entries(box.style)) {
      if (k.startsWith('--')) row.style.setProperty(k, v)
      else row.style[k as 'lineHeight'] = v
    }
    appendNoteRuns(row, it, 'footnote')
    row.addEventListener('dblclick', () => onEdit(it.id))
    wrap.appendChild(row)
  }
  return wrap
}
