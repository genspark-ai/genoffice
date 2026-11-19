import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { MouseEvent as ReactMouseEvent, KeyboardEvent as ReactKeyboardEvent } from 'react'
import type { Editor } from '@tiptap/core'
import type { SectionSettings } from '@genoffice/docx-engine'
import { t } from '../i18n/locale'
import { useMeasurement } from '../use-measurement'
import { GAP_BAND, PAGE_TURN_SEL } from '../editor/pagination-gaps'
import {
  INDENT_SNAP_TWIPS,
  MAX_RULER_INCHES,
  MIN_BODY_TWIPS,
  clampMarginDrag,
  rulerTicks,
  snapIndentTwips,
} from './Ruler'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** strip width and its distance from the paper's left edge, CSS px */
export const VRULER_WIDTH = 20
const VRULER_GAP = 4

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

interface Frame {
  /** paper top and left edge relative to the anchor, unzoomed CSS px */
  top: number
  left: number
}

function measureFrame(anchor: HTMLElement, editor: Editor | null): Frame | null {
  const live = anchor.parentElement
  const pm = live?.querySelector<HTMLElement>('.page-wrap .ProseMirror')
  if (!live || !pm) return null
  const zoomEl = anchor.closest<HTMLElement>('.doc-zoom')
  const factor = Number.parseFloat(zoomEl?.style.zoom ?? '') || 1
  const aRect = anchor.getBoundingClientRect()
  const pmRect = pm.getBoundingClientRect()
  const gaps = Array.from(pm.querySelectorAll<HTMLElement>(PAGE_TURN_SEL)).map((g) => ({
    top: g.getBoundingClientRect().top,
    marginBottomPx: Number.parseFloat(g.style.getPropertyValue('--gap-mb')) || 0,
  }))
  const tops = pageTopsOnScreen(pmRect.top, gaps, factor)
  let caretY = pmRect.top
  if (editor && !editor.isDestroyed) {
    try {
      caretY = editor.view.coordsAtPos(editor.state.selection.head).top
    } catch {
      /* position not laid out yet: stay on the first page */
    }
  }
  const pageTop = tops[pageIndexAt(tops, caretY)] ?? pmRect.top
  return {
    top: (pageTop - aRect.top) / factor,
    left: (pmRect.left - aRect.left) / factor - VRULER_WIDTH - VRULER_GAP,
  }
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
 * Word's vertical ruler: left of the page the caret is on, numbered from the
 * top margin outward, with gray top/bottom margin zones whose edges drag to
 * set the section's top and bottom margins. Zoomed with the page.
 */
export function VerticalRuler({
  section,
  editor,
  onMarginsChange,
}: {
  section: SectionSettings
  editor: Editor | null
  onMarginsChange?: (margins: { top: number; bottom: number }) => void
}) {
  const dims = verticalRulerDims(section)
  const measure = useMeasurement()
  const height = twipsToPx(dims.pageHeight)
  const marginTop = twipsToPx(dims.marginTop)
  const marginBottom = twipsToPx(dims.marginBottom)
  const anchorRef = useRef<HTMLDivElement>(null)
  const [frame, setFrame] = useState<Frame | null>(null)

  // follow the caret's page and the canvas layout: one measure per frame at most
  const rafRef = useRef(0)
  const schedule = useRef(() => {})
  schedule.current = () => {
    if (rafRef.current) return
    rafRef.current = requestAnimationFrame(() => {
      rafRef.current = 0
      const anchor = anchorRef.current
      if (!anchor) return
      const next = measureFrame(anchor, editor)
      setFrame((prev) =>
        prev && next && Math.abs(prev.top - next.top) < 0.5 && Math.abs(prev.left - next.left) < 0.5
          ? prev
          : next,
      )
    })
  }
  useLayoutEffect(() => schedule.current())
  useEffect(() => {
    const run = () => schedule.current()
    editor?.on('transaction', run)
    const pm = anchorRef.current?.parentElement?.querySelector('.page-wrap .ProseMirror')
    const ro = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(run)
    if (pm && ro) ro.observe(pm)
    window.addEventListener('resize', run)
    return () => {
      editor?.off('transaction', run)
      ro?.disconnect()
      window.removeEventListener('resize', run)
      cancelAnimationFrame(rafRef.current)
      rafRef.current = 0
    }
  }, [editor])

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
    <div className="vruler-anchor" ref={anchorRef}>
      {frame && (
        <div
          className="vruler"
          role="group"
          aria-label={t('appRulerVertical')}
          style={{ top: frame.top, left: frame.left, height }}
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
  )
}
