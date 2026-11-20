import { useRef, useState } from 'react'
import type { MouseEvent as ReactMouseEvent, KeyboardEvent as ReactKeyboardEvent } from 'react'
import type { Editor } from '@tiptap/core'
import type { SectionSettings } from '@genoffice/docx-engine'
import { t } from '../i18n/locale'
import { useMeasurement } from '../use-measurement'
import { GAP_BAND, PAGE_TURN_SEL } from '../editor/pagination-gaps'
import { useRulerFollow } from './ruler-follow'
import {
  INDENT_SNAP_TWIPS,
  MAX_RULER_INCHES,
  MIN_BODY_TWIPS,
  clampMarginDrag,
  rulerTicks,
  snapIndentTwips,
} from './Ruler'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** strip width, CSS px before zoom */
export const VRULER_WIDTH = 20

export interface VerticalRulerDims {
  pageHeight: number
  /** magnitudes in twips; a negative w:top/w:bottom (Word's "exact" margin) keeps its sign on write-back */
  marginTop: number
  marginBottom: number
  topSign: 1 | -1
  bottomSign: 1 | -1
}

/** Section geometry arrives from the file: coerce to finite, in-range dims. */
export function verticalRulerDims(section: SectionSettings): VerticalRulerDims {
  const finite = (v: unknown, fallback: number): number =>
    typeof v === 'number' && Number.isFinite(v) ? v : fallback
  const pageHeight = Math.min(
    Math.max(finite(section.pageHeight, 15840), 0),
    MAX_RULER_INCHES * 1440,
  )
  const top = finite(section.marginTop, 1440)
  const bottom = finite(section.marginBottom, 1440)
  return {
    pageHeight,
    marginTop: Math.min(Math.abs(top), pageHeight),
    marginBottom: Math.min(Math.abs(bottom), pageHeight),
    topSign: top < 0 ? -1 : 1,
    bottomSign: bottom < 0 ? -1 : 1,
  }
}

/** Margin drag target for the top or bottom edge, keeping MIN_BODY_TWIPS of body. */
export function clampVerticalMarginDrag(
  side: 'top' | 'bottom',
  abs: number,
  dims: Pick<VerticalRulerDims, 'pageHeight' | 'marginTop' | 'marginBottom'>,
): number {
  return clampMarginDrag(side === 'top' ? 'left' : 'right', abs, {
    pageWidth: dims.pageHeight,
    marginLeft: dims.marginTop,
    marginRight: dims.marginBottom,
  })
}

/**
 * Paper tops on screen, page by page, from the canvas DOM: the first page
 * starts at the editor's top edge, every later one a gray band below the
 * bottom margin of the page-turn gap that opens it. Screen px.
 */
export function pageTopsOnScreen(
  pmTop: number,
  gaps: readonly { top: number; marginBottomPx: number }[],
  factor: number,
): number[] {
  const tops = [pmTop]
  for (const g of gaps) tops.push(g.top + (g.marginBottomPx + GAP_BAND) * factor)
  return tops.sort((a, b) => a - b)
}

/** index of the page whose paper starts at or above `y` (the caret's line top) */
export function pageIndexAt(tops: readonly number[], y: number): number {
  let idx = 0
  for (let i = 0; i < tops.length; i++) if (tops[i]! <= y + 0.5) idx = i
  return idx
}

/**
 * The page the vertical ruler measures: the caret's page while any of it is in
 * view (Word's ruler follows the insertion point), else the page a third of
 * the way down the view, so the scale never scrolls away with the caret.
 */
export function rulerPageIndex(
  tops: readonly number[],
  caretY: number,
  viewTop: number,
  viewBottom: number,
): number {
  const caret = pageIndexAt(tops, caretY)
  const next = tops[caret + 1] ?? Infinity
  if (tops[caret]! < viewBottom && next > viewTop) return caret
  return pageIndexAt(tops, viewTop + (viewBottom - viewTop) / 3)
}

/**
 * Top of the caret's page relative to the ruler's bar, in the bar's zoomed
 * coordinates (screen px / zoom), or null before the editor is laid out.
 */
function measurePageTop(bar: HTMLElement, editor: Editor | null, zoom: number): number | null {
  if (!editor || editor.isDestroyed) return null
  const pm = editor.view.dom as HTMLElement
  if (!pm.isConnected) return null
  const barRect = bar.getBoundingClientRect()
  const pmRect = pm.getBoundingClientRect()
  const gaps = Array.from(pm.querySelectorAll<HTMLElement>(PAGE_TURN_SEL)).map((g) => ({
    top: g.getBoundingClientRect().top,
    marginBottomPx: Number.parseFloat(g.style.getPropertyValue('--gap-mb')) || 0,
  }))
  const tops = pageTopsOnScreen(pmRect.top, gaps, zoom)
  let caretY = pmRect.top
  try {
    caretY = editor.view.coordsAtPos(editor.state.selection.head).top
  } catch {
    /* position not laid out yet: stay on the first page */
  }
  return (tops[rulerPageIndex(tops, caretY, barRect.top, barRect.bottom)]! - barRect.top) / zoom
}

/** a dotted line across the window at the dragged margin */
function dragGuide(rulerRect: DOMRect) {
  const el = document.createElement('div')
  el.className = 'ruler-drag-guide ruler-drag-guide-h'
  el.style.left = `${rulerRect.right}px`
  el.style.width = `${Math.max(0, window.innerWidth - rulerRect.right)}px`
  document.body.appendChild(el)
  return {
    moveTo(clientY: number) {
      el.style.top = `${clientY}px`
    },
    remove() {
      el.remove()
    },
  }
}

/**
 * Word's vertical ruler: a fixed strip at the left of the document view whose
 * scale lines up with the page the caret is on, numbered from the top margin
 * outward, with gray top/bottom margin zones whose edges drag to set the
 * section's top and bottom margins. The scale follows the zoom; the strip
 * stays put while the pages scroll.
 */
export function VerticalRuler({
  section,
  editor,
  zoom = 1,
  top = 0,
  onMarginsChange,
}: {
  section: SectionSettings
  editor: Editor | null
  /** document zoom factor (1 = 100%) */
  zoom?: number
  /** where the strip starts in the document area (below the horizontal ruler), px */
  top?: number
  onMarginsChange?: (margins: { top: number; bottom: number }) => void
}) {
  const dims = verticalRulerDims(section)
  const measure = useMeasurement()
  const height = twipsToPx(dims.pageHeight)
  const marginTop = twipsToPx(dims.marginTop)
  const marginBottom = twipsToPx(dims.marginBottom)
  const barRef = useRef<HTMLDivElement>(null)
  const [pageTop, setPageTop] = useState<number | null>(null)
  useRulerFollow(editor, () => {
    const bar = barRef.current
    if (!bar) return
    const next = measurePageTop(bar, editor, zoom)
    setPageTop((prev) =>
      prev !== null && next !== null && Math.abs(prev - next) < 0.25 ? prev : next,
    )
  })

  const ticks = rulerTicks(dims.pageHeight, dims.marginTop, measure.unit)

  const zoneRefs = useRef<{ top: HTMLElement | null; bottom: HTMLElement | null }>({
    top: null,
    bottom: null,
  })
  const marginAbs = (side: 'top' | 'bottom', value: number) =>
    side === 'top' ? value : dims.pageHeight - value
  const commit = (side: 'top' | 'bottom', m: number) =>
    onMarginsChange?.({
      top: (side === 'top' ? m : dims.marginTop) * dims.topSign,
      bottom: (side === 'bottom' ? m : dims.marginBottom) * dims.bottomSign,
    })

  const handleMarginMouseDown = (e: ReactMouseEvent<HTMLElement>, side: 'top' | 'bottom') => {
    if (!onMarginsChange) return
    e.stopPropagation()
    e.preventDefault()
    const ruler = e.currentTarget.closest('.vruler') as HTMLElement
    const rect = ruler.getBoundingClientRect()
    if (!(rect.height > 0)) return
    const twipsPerPx = dims.pageHeight / rect.height
    const orig = side === 'top' ? dims.marginTop : dims.marginBottom
    const origAbs = marginAbs(side, orig)
    const startY = e.clientY
    const guide = dragGuide(rect)
    // a plain click leaves an off-grid margin alone: only a moved pointer snaps
    const absAt = (ev: MouseEvent) =>
      ev.clientY === startY
        ? origAbs
        : snapIndentTwips(origAbs + (ev.clientY - startY) * twipsPerPx, !ev.altKey)
    const preview = (m: number) => {
      const zone = zoneRefs.current[side]
      if (!zone) return
      if (side === 'top') zone.style.height = `${twipsToPx(m)}px`
      else {
        zone.style.top = `${height - twipsToPx(m)}px`
        zone.style.height = `${twipsToPx(m)}px`
      }
    }
    const onMouseMove = (ev: MouseEvent) => {
      const m = clampVerticalMarginDrag(side, absAt(ev), dims)
      preview(m)
      guide.moveTo(rect.top + marginAbs(side, m) / twipsPerPx)
    }
    const onMouseUp = (ev: MouseEvent) => {
      document.removeEventListener('mousemove', onMouseMove)
      document.removeEventListener('mouseup', onMouseUp)
      guide.remove()
      const m = clampVerticalMarginDrag(side, absAt(ev), dims)
      if (m !== orig) commit(side, m)
      else preview(orig)
    }
    guide.moveTo(rect.top + origAbs / twipsPerPx)
    document.addEventListener('mousemove', onMouseMove)
    document.addEventListener('mouseup', onMouseUp)
  }

  const handleMarginKeyDown = (e: ReactKeyboardEvent<HTMLElement>, side: 'top' | 'bottom') => {
    if (!onMarginsChange || (e.key !== 'ArrowUp' && e.key !== 'ArrowDown')) return
    e.preventDefault()
    const orig = side === 'top' ? dims.marginTop : dims.marginBottom
    const delta = (e.key === 'ArrowDown' ? 1 : -1) * (e.shiftKey ? 720 : INDENT_SNAP_TWIPS)
    const m = clampVerticalMarginDrag(side, snapIndentTwips(marginAbs(side, orig) + delta), dims)
    if (m !== orig) commit(side, m)
  }

  const marginHandle = (side: 'top' | 'bottom') => (
    <span
      className="vruler-margin-handle"
      data-side={side}
      style={{ top: side === 'top' ? marginTop : height - marginBottom }}
      role="slider"
      aria-orientation="vertical"
      tabIndex={0}
      aria-label={t(side === 'top' ? 'appRulerTopMargin' : 'appRulerBottomMargin')}
      data-tip={`${t(side === 'top' ? 'appRulerTopMargin' : 'appRulerBottomMargin')}: ${measure.format(
        side === 'top' ? dims.marginTop : dims.marginBottom,
      )}`}
      aria-valuemin={0}
      aria-valuemax={dims.pageHeight - MIN_BODY_TWIPS}
      aria-valuenow={side === 'top' ? dims.marginTop : dims.marginBottom}
      aria-valuetext={measure.format(side === 'top' ? dims.marginTop : dims.marginBottom)}
      onMouseDown={(e) => handleMarginMouseDown(e, side)}
      onKeyDown={(e) => handleMarginKeyDown(e, side)}
    />
  )

  return (
    <div
      className="vruler-bar"
      ref={barRef}
      style={{ top, width: Math.ceil((VRULER_WIDTH + 4) * zoom) }}
    >
      <div className="vruler-zoom" style={{ zoom }}>
        {pageTop !== null && (
          <div
            className="vruler"
            role="group"
            aria-label={t('appRulerVertical')}
            style={{ top: pageTop, height }}
          >
            <div
              className="vruler-zone"
              ref={(el) => {
                zoneRefs.current.top = el
              }}
              style={{ top: 0, height: marginTop }}
            />
            <div
              className="vruler-zone"
              ref={(el) => {
                zoneRefs.current.bottom = el
              }}
              style={{ top: height - marginBottom, height: marginBottom }}
            />
            {ticks.map((tk) =>
              tk.kind === 'num' ? (
                <span key={tk.pos} className="vruler-num" style={{ top: twipsToPx(tk.pos) }}>
                  {tk.label}
                </span>
              ) : (
                <span
                  key={tk.pos}
                  className={tk.kind === 'mid' ? 'vruler-tick vruler-tick-mid' : 'vruler-tick'}
                  style={{ top: twipsToPx(tk.pos) }}
                />
              ),
            )}
            {onMarginsChange && marginHandle('top')}
            {onMarginsChange && marginHandle('bottom')}
          </div>
        )}
      </div>
    </div>
  )
}
