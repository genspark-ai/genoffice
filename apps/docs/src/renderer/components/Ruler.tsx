import { useRef, useState } from 'react'
import type { MouseEvent as ReactMouseEvent, KeyboardEvent as ReactKeyboardEvent } from 'react'
import type { Editor } from '@tiptap/core'
import type { SectionSettings, TabStop } from '@genoffice/docx-engine'
import { t, type StringKey } from '../i18n/locale'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** Largest page width the ruler will lay out: ticks, zones and default-stop
 *  guides all derive from it, so a corrupt/huge section cannot OOM the tab. */
export const MAX_RULER_INCHES = 50
/** Keyboard nudge step (Word snap grid) and Shift-nudge step. */
export const RULER_SNAP_TWIPS = 60

export interface RulerDims {
  /** finite page width in twips actually laid out */
  pageWidth: number
  marginLeft: number
  marginRight: number
  /** whole inches ticked, capped */
  inches: number
}

/** Section geometry arrives from the file: coerce to finite, in-range dims. */
export function rulerDims(section: SectionSettings): RulerDims {
  const finite = (v: unknown, fallback: number): number =>
    typeof v === 'number' && Number.isFinite(v) && v >= 0 ? v : fallback
  const pageWidth = finite(section.pageWidth, 12240)
  const marginLeft = Math.min(finite(section.marginLeft, 1440), pageWidth)
  const marginRight = Math.min(finite(section.marginRight, 1440), pageWidth)
  const inches = Math.min(Math.floor(pageWidth / 1440), MAX_RULER_INCHES)
  return { pageWidth, marginLeft, marginRight, inches }
}

/** Word tab-stop snap grid: nearest 60 twips (~0.04in). */
export function snapTabTwips(posTwips: number): number {
  if (!Number.isFinite(posTwips)) return 0
  return Math.round(posTwips / RULER_SNAP_TWIPS) * RULER_SNAP_TWIPS
}

/** A `clear` stop cancels an inherited stop — it marks no position, so the
    ruler renders nothing for it (write-back still carries it). Exported for tests. */
export function isRenderableTabStop(stop: TabStop): boolean {
  return stop.val !== 'clear'
}

/** A ruler edit makes the whole set direct (Word writes style-inherited stops
    out too). An inherited stop the user removed or moved needs a `clear` at its
    old position, or the style chain puts it back on reopen. Exported for tests. */
export function directTabStops(original: TabStop[], edited: TabStop[]): TabStop[] {
  const clears = original
    .filter((s) => s.inherited && !edited.some((e) => e.pos === s.pos))
    .map((s): TabStop => ({ pos: s.pos, val: 'clear' }))
  return [...edited.map(({ inherited: _inherited, ...s }) => s), ...clears].sort(
    (a, b) => a.pos - b.pos,
  )
}

/**
 * Paragraph indents as the ruler writes them back. `null` clears the direct
 * value so the style chain shows through again — the same contract the
 * paragraph dialog uses (0pt means "no direct indent", not "indent zero").
 */
export interface RulerIndents {
  left: number | null
  right: number | null
  firstLine: number | null
}

/** which indent marker a drag or key press is editing */
export type RulerIndentKind = 'left' | 'first' | 'right'

/** Indents in twips as the ruler draws them: an inherited (`null`) value reads
    as zero, which is where the marker sits for a paragraph with no direct
    indent. Exported for tests. */
export function rulerIndents(attrs: Record<string, unknown>): {
  left: number
  right: number
  firstLine: number
} {
  const twips = (v: unknown): number => (typeof v === 'number' && Number.isFinite(v) ? v : 0)
  return {
    left: Math.max(0, twips(attrs.indentLeft)),
    right: Math.max(0, twips(attrs.indentRight)),
    firstLine: twips(attrs.indentFirstLine),
  }
}

/** Marker positions in twips, measured from the page's left edge. The
    first-line marker is relative to the left indent — negative is a hanging
    indent — and the right marker measures back from the right margin.
    Exported for tests. */
export function rulerIndentPositions(
  dims: RulerDims,
  indents: { left: number; right: number; firstLine: number },
): { left: number; first: number; right: number } {
  const contentLeft = dims.marginLeft
  const contentRight = dims.pageWidth - dims.marginRight
  const clamp = (v: number) => Math.min(Math.max(v, contentLeft), contentRight)
  return {
    left: clamp(contentLeft + indents.left),
    first: clamp(contentLeft + indents.left + indents.firstLine),
    right: clamp(contentRight - indents.right),
  }
}

/** What dropping the `kind` marker at an absolute ruler position writes back.
    Dragging the left-indent block carries the first-line marker with it, as in
    Word: the first-line indent is an offset, so it survives the move. The other
    two indents are echoed unchanged — a write-back replaces the whole set.
    Exported for tests. */
export function indentFromRuler(
  kind: 'left' | 'first' | 'right',
  dims: RulerDims,
  indents: { left: number; right: number; firstLine: number },
  posTwips: number,
): RulerIndents {
  const contentWidth = Math.max(0, dims.pageWidth - dims.marginLeft - dims.marginRight)
  const clamp = (v: number, lo: number, hi: number) => Math.min(Math.max(v, lo), hi)
  // 0 means "no direct value"; a hanging first line is negative and must survive
  const keep = (v: number): number | null => (v !== 0 ? v : null)
  if (kind === 'left') {
    return {
      left: keep(clamp(Math.round(posTwips) - dims.marginLeft, 0, contentWidth)),
      right: keep(indents.right),
      firstLine: keep(indents.firstLine),
    }
  }
  if (kind === 'right') {
    return {
      left: keep(indents.left),
      right: keep(clamp(dims.pageWidth - dims.marginRight - Math.round(posTwips), 0, contentWidth)),
      firstLine: keep(indents.firstLine),
    }
  }
  return {
    left: keep(indents.left),
    right: keep(indents.right),
    firstLine: keep(
      clamp(
        Math.round(posTwips) - dims.marginLeft - indents.left,
        -indents.left,
        Math.max(0, contentWidth - indents.left),
      ),
    ),
  }
}

/** Horizontal ruler above the page: inch numbers, gray margin zones, tab stops. */
export function Ruler({
  section,
  editor,
  onTabStopsChange,
  onIndentsChange,
}: {
  section: SectionSettings
  editor: Editor | null
  onTabStopsChange: (stops: TabStop[] | null) => void
  onIndentsChange: (indents: RulerIndents) => void
}) {
  const dims = rulerDims(section)
  const width = twipsToPx(dims.pageWidth)
  const marginLeft = twipsToPx(dims.marginLeft)
  const marginRight = twipsToPx(dims.marginRight)
  const ticks: number[] = []
  for (let i = 1; i <= dims.inches; i++) ticks.push(i)

  // Default Word tab interval: 0.5in = 720 twips
  const DEFAULT_TAB_TWIPS = 720

  // Tab stop type cycling (Word: click ruler button to cycle L/C/R/Decimal/Bar)
  const [nextTabType, setNextTabType] = useState<TabStop['val']>('left')
  const TAB_TYPE_CYCLE: TabStop['val'][] = ['left', 'center', 'right', 'decimal', 'bar']
  const TAB_TYPE_LABELS: Record<string, string> = {
    left: 'L',
    center: '⊥',
    right: '⌐',
    decimal: '.',
    bar: '|',
  }
  const TAB_TYPE_NAME_KEYS: Record<TabStop['val'], StringKey> = {
    left: 'appTabLeft',
    center: 'appTabCenter',
    right: 'appTabRight',
    decimal: 'appTabDecimal',
    bar: 'appTabBar',
    clear: 'appTabClear',
  }
  const INDENT_LABEL_KEYS: Record<RulerIndentKind, StringKey> = {
    first: 'appIndentFirstLine',
    left: 'appIndentLeft',
    right: 'appIndentRight',
  }

  /** properties of the paragraph the ruler edits — headings and list items
   *  carry the same paragraph properties as docParagraph */
  const paraAttrs = (): Record<string, unknown> => {
    if (!editor) return {}
    const attrs = editor.isActive('docHeading')
      ? editor.getAttributes('docHeading')
      : editor.isActive('docListItem')
        ? editor.getAttributes('docListItem')
        : editor.getAttributes('docParagraph')
    return (attrs ?? {}) as Record<string, unknown>
  }

  // Get current tab stops from focused paragraph. rel stops mirror w:ptab
  // (percent positions): not draggable ruler stops, but every write-back must
  // carry them or a ruler edit silently drops the paragraph's ptab layout.
  const currentTabStops = (): { stops: TabStop[]; relStops: TabStop[] } => {
    const raw = paraAttrs().tabStops as string | null
    if (!raw) return { stops: [], relStops: [] }
    try {
      const parsed = JSON.parse(raw)
      if (!Array.isArray(parsed)) return { stops: [], relStops: [] }
      return { stops: parsed.filter((s) => !s.rel), relStops: parsed.filter((s) => s.rel) }
    } catch {
      return { stops: [], relStops: [] }
    }
  }

  const { stops, relStops } = currentTabStops()

  // Paragraph indent markers: Word puts three on the ruler — the first-line
  // triangle, the left-indent block beneath it, and the right-indent triangle.
  // They were missing entirely, so a drag on the ruler could only ever place a
  // tab stop and the paragraph indent never moved.
  const indents = rulerIndents(paraAttrs())
  const marks = rulerIndentPositions(dims, indents)

  const startIndentDrag = (e: ReactMouseEvent<HTMLSpanElement>, kind: RulerIndentKind) => {
    e.stopPropagation()
    e.preventDefault()
    const rect = (e.currentTarget.closest('.ruler') as HTMLElement).getBoundingClientRect()
    const marker = e.currentTarget
    const contentLeft = marginLeft
    const contentRight = width - marginRight
    const clampX = (clientX: number) =>
      Math.min(Math.max(clientX - rect.left, contentLeft), contentRight)
    const onMouseMove = (ev: MouseEvent) => {
      // visual only, as with the tab stops: no state update per mouse move
      marker.style.left = `${clampX(ev.clientX)}px`
    }
    const onMouseUp = (ev: MouseEvent) => {
      document.removeEventListener('mousemove', onMouseMove)
      document.removeEventListener('mouseup', onMouseUp)
      const posTwips = width > 0 ? (clampX(ev.clientX) / width) * dims.pageWidth : 0
      onIndentsChange(indentFromRuler(kind, dims, indents, posTwips))
    }
    document.addEventListener('mousemove', onMouseMove)
    document.addEventListener('mouseup', onMouseUp)
  }

  const nudgeIndent = (e: ReactKeyboardEvent<HTMLSpanElement>, kind: RulerIndentKind) => {
    if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return
    e.preventDefault()
    const delta = (e.key === 'ArrowRight' ? 1 : -1) * (e.shiftKey ? 720 : RULER_SNAP_TWIPS)
    onIndentsChange(indentFromRuler(kind, dims, indents, snapTabTwips(marks[kind] + delta)))
  }

  const withRel = (edited: TabStop[]): TabStop[] | null => {
    const direct = directTabStops(stops, edited)
    return direct.length > 0 || relStops.length > 0 ? [...direct, ...relStops] : null
  }

  // Drag state
  const dragRef = useRef<{ stopIndex: number; startX: number; origPos: number } | null>(null)

  // Click on ruler: add tab stop at position, skip margin zones
  const handleRulerClick = (e: ReactMouseEvent<HTMLDivElement>) => {
    if (!(width > 0) || !(dims.pageWidth > 0)) return
    const rect = (e.currentTarget as HTMLElement).getBoundingClientRect()
    const x = e.clientX - rect.left
    if (x < marginLeft || x > width - marginRight) return
    const posTwips = Math.round((x / width) * dims.pageWidth)
    // snap to the Word grid
    const snapped = snapTabTwips(posTwips)
    const existing = stops.filter((s) => Math.abs(s.pos - snapped) > 60)
    const newStop: TabStop = { pos: snapped, val: nextTabType }
    const newStops = [...existing, newStop].sort((a, b) => a.pos - b.pos)
    onTabStopsChange(withRel(newStops))
  }

  // Drag tab stop to new position or drop outside to delete
  const handleTabMouseDown = (e: ReactMouseEvent<HTMLSpanElement>, stopIndex: number) => {
    e.stopPropagation()
    e.preventDefault()
    const rect = (e.currentTarget.closest('.ruler') as HTMLElement).getBoundingClientRect()
    dragRef.current = { stopIndex, startX: e.clientX, origPos: stops[stopIndex].pos }

    const onMouseMove = (ev: MouseEvent) => {
      if (!dragRef.current) return
      const x = ev.clientX - rect.left
      const posTwips = Math.round((x / width) * dims.pageWidth)
      const snapped = snapTabTwips(posTwips)
      // visual only update via CSS custom property (no state update for perf)
      const marker = document.querySelector(
        `[data-ruler-stop="${stopIndex}"]`,
      ) as HTMLElement | null
      if (marker) marker.style.left = `${twipsToPx(snapped)}px`
    }

    const onMouseUp = (ev: MouseEvent) => {
      document.removeEventListener('mousemove', onMouseMove)
      document.removeEventListener('mouseup', onMouseUp)
      if (!dragRef.current) return
      const x = ev.clientX - rect.left
      // drop outside the content area: delete the stop
      if (x < marginLeft || x > width - marginRight) {
        const newStops = stops.filter((_, i) => i !== dragRef.current!.stopIndex)
        onTabStopsChange(withRel(newStops))
      } else {
        const posTwips = Math.round((x / width) * dims.pageWidth)
        const snapped = snapTabTwips(posTwips)
        const newStops = stops
          .map((s, i) => (i === dragRef.current!.stopIndex ? { ...s, pos: snapped } : s))
          .sort((a, b) => a.pos - b.pos)
        onTabStopsChange(withRel(newStops))
      }
      dragRef.current = null
    }

    document.addEventListener('mousemove', onMouseMove)
    document.addEventListener('mouseup', onMouseUp)
  }

  // Default tab stop markers (light gray) when no custom stops mark a position.
  // The content width is capped: with corrupt geometry the loop below would
  // otherwise push millions of guides.
  const defaultStops: number[] = []
  if (!stops.some(isRenderableTabStop)) {
    const contentWidth = Math.min(
      dims.pageWidth - dims.marginLeft - dims.marginRight,
      MAX_RULER_INCHES * 1440,
    )
    for (let pos = DEFAULT_TAB_TWIPS; pos < contentWidth; pos += DEFAULT_TAB_TWIPS) {
      defaultStops.push(dims.marginLeft + pos)
    }
  }

  // Keyboard: arrows nudge the focused stop on the snap grid, Delete removes it.
  const handleStopKeyDown = (e: ReactKeyboardEvent<HTMLSpanElement>, stopIndex: number) => {
    const stop = stops[stopIndex]
    if (!stop) return
    if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
      e.preventDefault()
      const delta = (e.key === 'ArrowRight' ? 1 : -1) * (e.shiftKey ? 720 : 60)
      const pos = Math.min(Math.max(snapTabTwips(stop.pos + delta), 0), dims.pageWidth)
      const newStops = stops
        .map((s, i) => (i === stopIndex ? { ...s, pos } : s))
        .sort((a, b) => a.pos - b.pos)
      onTabStopsChange(withRel(newStops))
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault()
      onTabStopsChange(withRel(stops.filter((_, i) => i !== stopIndex)))
    }
  }

  const stopLabel = (stop: TabStop): string =>
    t('appTabStopTitle', {
      type: t(TAB_TYPE_NAME_KEYS[stop.val]),
      pos: Math.round((stop.pos / 144) * 10) / 10,
    }) + (stop.leader ? t('appTabLeader', { leader: stop.leader }) : '')

  return (
    <div className="ruler" role="group" style={{ width }} onClick={handleRulerClick}>
      {/* Tab type selector button at far left */}
      <button
        className="ruler-tab-type"
        data-tip={t('appTabTypeTip', { type: t(TAB_TYPE_NAME_KEYS[nextTabType]) })}
        onClick={(e) => {
          e.stopPropagation()
          const idx = TAB_TYPE_CYCLE.indexOf(nextTabType)
          setNextTabType(TAB_TYPE_CYCLE[(idx + 1) % TAB_TYPE_CYCLE.length])
        }}
      >
        {TAB_TYPE_LABELS[nextTabType]}
      </button>

      <div className="ruler-zone" style={{ left: 0, width: marginLeft }} />
      <div className="ruler-zone" style={{ left: width - marginRight, width: marginRight }} />

      {ticks.map((i) => (
        <span key={i} className="ruler-num" style={{ left: twipsToPx(i * 1440) }}>
          {i}
        </span>
      ))}

      {/* Default tab stop guides (light, no interaction) */}
      {defaultStops.map((posTwips) => (
        <span
          key={`def-${posTwips}`}
          className="ruler-tab-default"
          style={{ left: twipsToPx(posTwips) }}
        />
      ))}

      {/* Custom tab stops (interactive). A `clear` stop cancels inherited
          stops at its position — it places no mark, so it renders nothing
          (returning null keeps data-ruler-stop indexes aligned with `stops`
          for drag handling) while write-back still preserves it. */}
      {stops.map((stop, i) =>
        !isRenderableTabStop(stop) ? null : (
          <span
            key={`${stop.pos}-${i}`}
            data-ruler-stop={i}
            className={`ruler-tab ruler-tab-${stop.val}`}
            style={{ left: Math.min(Math.max(twipsToPx(stop.pos), 0), width) }}
            data-tip={stopLabel(stop)}
            role="slider"
            tabIndex={0}
            aria-label={stopLabel(stop)}
            aria-valuemin={0}
            aria-valuemax={dims.pageWidth}
            aria-valuenow={Number.isFinite(stop.pos) ? stop.pos : 0}
            onMouseDown={(e) => handleTabMouseDown(e, i)}
            onKeyDown={(e) => handleStopKeyDown(e, i)}
          >
            {TAB_TYPE_LABELS[stop.val]}
          </span>
        ),
      )}

      {/* Paragraph indent markers, laid out like Word's: the first-line
          triangle above the left-indent block, and a triangle at the right
          margin. Each is a slider like the tab stops — drag to move, arrows to
          nudge on the snap grid — and each swallows its own click so releasing
          a marker never also places a tab stop. Rendered after the tab stops
          so a tab stop stays the first slider in the ruler. */}
      {(['first', 'left', 'right'] as RulerIndentKind[]).map((kind) => {
        const label = t(INDENT_LABEL_KEYS[kind])
        const contentWidth = Math.max(0, dims.pageWidth - dims.marginLeft - dims.marginRight)
        const value = kind === 'first' ? indents.firstLine : indents[kind]
        return (
          <span
            key={`indent-${kind}`}
            data-ruler-indent={kind}
            className={`ruler-indent ruler-indent-${kind}`}
            style={{ left: twipsToPx(marks[kind]) }}
            data-tip={label}
            role="slider"
            tabIndex={0}
            aria-label={label}
            aria-valuemin={kind === 'first' ? -indents.left : 0}
            aria-valuemax={contentWidth}
            aria-valuenow={Number.isFinite(value) ? value : 0}
            onMouseDown={(e) => startIndentDrag(e, kind)}
            onKeyDown={(e) => nudgeIndent(e, kind)}
            onClick={(e) => e.stopPropagation()}
          />
        )
      })}
    </div>
  )
}
