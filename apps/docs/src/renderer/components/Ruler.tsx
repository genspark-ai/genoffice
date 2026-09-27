import { useRef, useState } from 'react'
import type { MouseEvent as ReactMouseEvent, KeyboardEvent as ReactKeyboardEvent } from 'react'
import type { Editor } from '@tiptap/core'
import type { SectionSettings, TabStop } from '@genoffice/docx-engine'
import { t, type StringKey } from '../i18n/locale'
import { useMeasurement } from '../use-measurement'
import type { MeasurementUnit } from '../units'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** Largest page width the ruler will lay out: ticks, zones and default-stop
 *  guides all derive from it, so a corrupt/huge section cannot OOM the tab. */
export const MAX_RULER_INCHES = 50
const MAX_RULER_TWIPS = MAX_RULER_INCHES * 1440
/** Keyboard nudge step (Word snap grid) and Shift-nudge step. */
export const RULER_SNAP_TWIPS = 60
/** Word snaps indent and margin markers to 1/16 in (Option disables it). */
export const INDENT_SNAP_TWIPS = 90
/** Word refuses a margin drag that leaves less than half an inch of body. */
export const MIN_BODY_TWIPS = 720

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
  const pageWidth = Math.min(finite(section.pageWidth, 12240), MAX_RULER_TWIPS)
  const marginLeft = Math.min(finite(section.marginLeft, 1440), pageWidth)
  const marginRight = Math.min(finite(section.marginRight, 1440), pageWidth)
  const inches = Math.floor(pageWidth / 1440)
  return { pageWidth, marginLeft, marginRight, inches }
}

/** Ruler graduation per unit (twips): numbered majors, a taller mid tick and
 *  minor ticks; `perMajor` is the printed count per major (36 pt, 10 mm, 6 pi). */
const RULER_SCALES: Record<
  MeasurementUnit,
  { major: number; mid: number; minor: number; perMajor: number }
> = {
  in: { major: 1440, mid: 720, minor: 180, perMajor: 1 },
  cm: { major: 1440 / 2.54, mid: 720 / 2.54, minor: 360 / 2.54, perMajor: 1 },
  mm: { major: 1440 / 2.54, mid: 720 / 2.54, minor: 360 / 2.54, perMajor: 10 },
  pt: { major: 720, mid: 360, minor: 180, perMajor: 36 },
  pi: { major: 1440, mid: 720, minor: 240, perMajor: 6 },
}

export interface RulerTick {
  /** twips from the page's left edge */
  pos: number
  kind: 'num' | 'mid' | 'minor'
  label?: string
}

/** Word numbers the ruler from the text origin (left margin, or the cell's
 *  left edge) outward on both sides; the origin itself carries no number. */
export function rulerTicks(pageWidth: number, origin: number, unit: MeasurementUnit): RulerTick[] {
  const sc = RULER_SCALES[unit]
  const perMajor = Math.round(sc.major / sc.minor)
  const perMid = Math.round(sc.mid / sc.minor)
  const out: RulerTick[] = []
  const push = (k: number, pos: number) => {
    if (k % perMajor === 0)
      out.push({ pos, kind: 'num', label: String((k / perMajor) * sc.perMajor) })
    else out.push({ pos, kind: k % perMid === 0 ? 'mid' : 'minor' })
  }
  for (let k = 1; origin + k * sc.minor <= pageWidth + 0.5; k++) push(k, origin + k * sc.minor)
  for (let k = 1; origin - k * sc.minor >= -0.5; k++) push(k, origin - k * sc.minor)
  return out
}

/** Word tab-stop snap grid: nearest 60 twips (~0.04in). */
export function snapTabTwips(posTwips: number): number {
  if (!Number.isFinite(posTwips)) return 0
  return Math.round(posTwips / RULER_SNAP_TWIPS) * RULER_SNAP_TWIPS
}

export function snapIndentTwips(posTwips: number, snap = true): number {
  if (!Number.isFinite(posTwips)) return 0
  return snap ? Math.round(posTwips / INDENT_SNAP_TWIPS) * INDENT_SNAP_TWIPS : Math.round(posTwips)
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

/** Paragraph indents in twips, w:ind semantics: negative firstLine = hanging. */
export interface ParagraphIndents {
  left: number
  right: number
  firstLine: number
}

/** The column the indents are measured from: the page body, or the cell the caret is in. */
export interface BodyBounds {
  left: number
  right: number
  inCell: boolean
}

export type IndentMarker = 'firstLine' | 'hanging' | 'left' | 'right'

export interface IndentAttrs {
  indentLeft?: number
  indentRight?: number
  indentFirstLine?: number
}

/** Absolute marker positions (twips from the page's left edge). */
export function indentMarkerPositions(
  ind: ParagraphIndents,
  body: BodyBounds,
): Record<IndentMarker, number> {
  return {
    firstLine: body.left + ind.left + ind.firstLine,
    hanging: body.left + ind.left,
    left: body.left + ind.left,
    right: body.right - ind.right,
  }
}

/**
 * Word's marker semantics: first-line moves only the first line, hanging moves
 * the wrapped lines while the first line stays put, the box moves both, right
 * moves the right edge. `abs` is the snapped target position of the dragged
 * marker; markers never cross each other or leave the page.
 */
export function applyIndentDrag(
  marker: IndentMarker,
  ind: ParagraphIndents,
  abs: number,
  body: BodyBounds,
  pageWidth: number,
): { indents: ParagraphIndents; attrs: IndentAttrs } {
  const pos = indentMarkerPositions(ind, body)
  const clamp = (v: number, lo: number, hi: number) => Math.min(Math.max(v, lo), hi)
  if (marker === 'right') {
    const right = clamp(abs, Math.max(pos.firstLine, pos.hanging) + INDENT_SNAP_TWIPS, pageWidth)
    const indentRight = body.right - right
    return { indents: { ...ind, right: indentRight }, attrs: { indentRight } }
  }
  const hi = pos.right - INDENT_SNAP_TWIPS
  if (marker === 'firstLine') {
    const indentFirstLine = clamp(abs, 0, hi) - body.left - ind.left
    return { indents: { ...ind, firstLine: indentFirstLine }, attrs: { indentFirstLine } }
  }
  if (marker === 'hanging') {
    const indentLeft = clamp(abs, 0, hi) - body.left
    const indentFirstLine = pos.firstLine - body.left - indentLeft
    return {
      indents: { ...ind, left: indentLeft, firstLine: indentFirstLine },
      attrs: { indentLeft, indentFirstLine },
    }
  }
  // the first-line marker rides along with the box, so it bounds the box too
  const lo = Math.max(0, -ind.firstLine)
  const indentLeft = clamp(abs, lo, Math.max(lo, Math.min(hi, hi - ind.firstLine))) - body.left
  return { indents: { ...ind, left: indentLeft }, attrs: { indentLeft } }
}

/** Margin drag target for one side, keeping at least MIN_BODY_TWIPS of body. */
export function clampMarginDrag(
  side: 'left' | 'right',
  abs: number,
  dims: Pick<RulerDims, 'pageWidth' | 'marginLeft' | 'marginRight'>,
): number {
  if (side === 'left') {
    const max = dims.pageWidth - dims.marginRight - MIN_BODY_TWIPS
    return Math.min(Math.max(abs, 0), Math.max(max, 0))
  }
  const min = dims.marginLeft + MIN_BODY_TWIPS
  return dims.pageWidth - Math.max(Math.min(abs, dims.pageWidth), Math.min(min, dims.pageWidth))
}

/** CSS length (computed px, or a custom property in pt/in/px) to twips. */
export function cssLengthToTwips(value: string | null | undefined): number | null {
  const m = /^\s*(-?[\d.]+)\s*(px|pt|in|cm|mm)?\s*$/.exec(value ?? '')
  if (!m) return null
  const n = Number(m[1])
  if (!Number.isFinite(n)) return null
  switch (m[2]) {
    case 'pt':
      return Math.round(n * 20)
    case 'in':
      return Math.round(n * 1440)
    case 'cm':
      return Math.round((n / 2.54) * 1440)
    case 'mm':
      return Math.round((n / 25.4) * 1440)
    default:
      return Math.round(n * 15)
  }
}

const paraAttrsOf = (editor: Editor): Record<string, unknown> =>
  editor.isActive('docHeading')
    ? editor.getAttributes('docHeading')
    : editor.isActive('docListItem')
      ? editor.getAttributes('docListItem')
      : editor.getAttributes('docParagraph')

const caretParagraphDom = (editor: Editor): HTMLElement | null => {
  const $from = editor.state?.selection?.$from
  if (!$from || $from.depth < 1) return null
  const dom = editor.view?.nodeDOM($from.before())
  return dom instanceof HTMLElement ? dom : null
}

/**
 * Word's ruler shows the effective indent: a direct w:ind value when the
 * paragraph has one, else what the style or numbering level lays out (read
 * back from the rendered paragraph so list geometry and style CSS agree).
 */
function paragraphIndents(editor: Editor): ParagraphIndents {
  const attrs = paraAttrsOf(editor)
  const num = (v: unknown): number | null =>
    typeof v === 'number' && Number.isFinite(v) ? v : null
  let left = num(attrs.indentLeft)
  let right = num(attrs.indentRight)
  let firstLine = num(attrs.indentFirstLine)
  if (left === null || right === null || firstLine === null) {
    const el = caretParagraphDom(editor)
    const cs = el ? getComputedStyle(el) : null
    const isList = editor.isActive('docListItem')
    if (cs) {
      if (left === null)
        left =
          (cssLengthToTwips(cs.marginInlineStart) ?? 0) +
          (isList ? (cssLengthToTwips(cs.paddingInlineStart) ?? 0) : 0)
      if (right === null) right = cssLengthToTwips(cs.marginInlineEnd) ?? 0
      if (firstLine === null) {
        const textIndent = cssLengthToTwips(cs.textIndent) ?? 0
        if (isList && textIndent <= 0) {
          const hang =
            cssLengthToTwips(cs.getPropertyValue('--li-hang')) ??
            cssLengthToTwips(cs.getPropertyValue('--style-li-hang')) ??
            360
          firstLine = -hang
        } else firstLine = textIndent
      }
    }
  }
  return { left: left ?? 0, right: right ?? 0, firstLine: firstLine ?? 0 }
}

/** In a table Word measures indents from the cell; the cell's content box is
 *  read from the page so the markers line up whatever the zoom. */
function bodyBounds(editor: Editor, dims: RulerDims): BodyBounds {
  const margins: BodyBounds = {
    left: dims.marginLeft,
    right: dims.pageWidth - dims.marginRight,
    inCell: false,
  }
  const $from = editor.state?.selection?.$from
  if (!$from) return margins
  for (let d = $from.depth; d > 0; d--) {
    const name = $from.node(d).type.name
    if (name !== 'docTableCell' && name !== 'docTableHeader') continue
    const cell = editor.view?.nodeDOM($from.before(d))
    const page = cell instanceof HTMLElement ? cell.closest('.doc-page') : null
    if (!(cell instanceof HTMLElement) || !(page instanceof HTMLElement))
      return { ...margins, inCell: true }
    const pageRect = page.getBoundingClientRect()
    if (!(pageRect.width > 0)) return { ...margins, inCell: true }
    const cellRect = cell.getBoundingClientRect()
    const cs = getComputedStyle(cell)
    const scale = dims.pageWidth / pageRect.width
    // computed lengths are unzoomed CSS px, the rects are zoomed viewport px
    const zoom = page.offsetWidth > 0 ? pageRect.width / page.offsetWidth : 1
    const edge = (pad: string, border: string) =>
      ((parseFloat(pad) || 0) + (parseFloat(border) || 0)) * zoom
    const padL = edge(cs.paddingLeft, cs.borderLeftWidth)
    const padR = edge(cs.paddingRight, cs.borderRightWidth)
    return {
      left: Math.round((cellRect.left + padL - pageRect.left) * scale),
      right: Math.round((cellRect.right - padR - pageRect.left) * scale),
      inCell: true,
    }
  }
  return margins
}

/** Dotted guide down the page while a marker is dragged (Word). */
function dragGuide(rulerRect: DOMRect) {
  const el = document.createElement('div')
  el.className = 'ruler-drag-guide'
  el.style.top = `${rulerRect.bottom}px`
  el.style.height = `${Math.max(0, window.innerHeight - rulerRect.bottom)}px`
  document.body.appendChild(el)
  return {
    moveTo(clientX: number) {
      el.style.left = `${clientX}px`
    },
    remove() {
      el.remove()
    },
  }
}

const INDENT_MARKER_KEYS: Record<IndentMarker, StringKey> = {
  firstLine: 'appRulerFirstLineIndent',
  hanging: 'appRulerHangingIndent',
  left: 'appRulerLeftIndent',
  right: 'appRulerRightIndent',
}

/** Horizontal ruler above the page: unit graduation from the text origin,
 *  gray margin zones, tab stops, indent and margin markers. */
export function Ruler({
  section,
  editor,
  onTabStopsChange,
  onIndentsChange,
  onMarginsChange,
}: {
  section: SectionSettings
  editor: Editor | null
  onTabStopsChange: (stops: TabStop[] | null) => void
  onIndentsChange?: (attrs: IndentAttrs) => void
  onMarginsChange?: (margins: { left: number; right: number }) => void
}) {
  const dims = rulerDims(section)
  const measure = useMeasurement()
  const width = twipsToPx(dims.pageWidth)
  const marginLeft = twipsToPx(dims.marginLeft)
  const marginRight = twipsToPx(dims.marginRight)

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

  // Get current tab stops from focused paragraph. rel stops mirror w:ptab
  // (percent positions): not draggable ruler stops, but every write-back must
  // carry them or a ruler edit silently drops the paragraph's ptab layout.
  const currentTabStops = (): { stops: TabStop[]; relStops: TabStop[] } => {
    if (!editor) return { stops: [], relStops: [] }
    const raw = paraAttrsOf(editor)?.tabStops as string | null
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
  const withRel = (edited: TabStop[]): TabStop[] | null => {
    const direct = directTabStops(stops, edited)
    return direct.length > 0 || relStops.length > 0 ? [...direct, ...relStops] : null
  }

  const indents = editor ? paragraphIndents(editor) : { left: 0, right: 0, firstLine: 0 }
  const body = editor
    ? bodyBounds(editor, dims)
    : { left: dims.marginLeft, right: dims.pageWidth - dims.marginRight, inCell: false }
  const markerPos = indentMarkerPositions(indents, body)
  const ticks = rulerTicks(dims.pageWidth, body.left, measure.unit)

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

  // Indent and margin markers: the ruler is zoomed with the page, so pointer
  // deltas map through its on-screen width. Live feedback is a guide line plus
  // the marker itself; the document changes once, on release.
  const startMarkerDrag = (
    e: ReactMouseEvent<HTMLElement>,
    origAbs: number,
    onMove: (abs: number) => void,
    onDrop: (abs: number) => void,
  ) => {
    e.stopPropagation()
    e.preventDefault()
    const ruler = e.currentTarget.closest('.ruler') as HTMLElement
    const rect = ruler.getBoundingClientRect()
    if (!(rect.width > 0)) return
    const twipsPerPx = dims.pageWidth / rect.width
    const startX = e.clientX
    const guide = dragGuide(rect)
    // a plain click leaves an off-grid marker alone: only a moved pointer snaps
    const absAt = (ev: MouseEvent) =>
      ev.clientX === startX
        ? origAbs
        : snapIndentTwips(origAbs + (ev.clientX - startX) * twipsPerPx, !ev.altKey)
    const onMouseMove = (ev: MouseEvent) => {
      const abs = absAt(ev)
      onMove(abs)
      guide.moveTo(rect.left + abs / twipsPerPx)
    }
    const onMouseUp = (ev: MouseEvent) => {
      document.removeEventListener('mousemove', onMouseMove)
      document.removeEventListener('mouseup', onMouseUp)
      guide.remove()
      const abs = absAt(ev)
      if (abs !== origAbs) onDrop(abs)
      else onMove(origAbs)
    }
    guide.moveTo(rect.left + origAbs / twipsPerPx)
    document.addEventListener('mousemove', onMouseMove)
    document.addEventListener('mouseup', onMouseUp)
  }

  const markerRefs = useRef<Partial<Record<IndentMarker, HTMLElement | null>>>({})
  const placeMarkers = (pos: Record<IndentMarker, number>) => {
    for (const k of Object.keys(pos) as IndentMarker[]) {
      const el = markerRefs.current[k]
      if (el) el.style.left = `${twipsToPx(pos[k])}px`
    }
  }

  const handleIndentMouseDown = (e: ReactMouseEvent<HTMLElement>, marker: IndentMarker) => {
    if (!onIndentsChange) return
    const preview = (abs: number) =>
      applyIndentDrag(marker, indents, abs, body, dims.pageWidth).indents
    startMarkerDrag(
      e,
      markerPos[marker],
      (abs) => placeMarkers(indentMarkerPositions(preview(abs), body)),
      (abs) => onIndentsChange(applyIndentDrag(marker, indents, abs, body, dims.pageWidth).attrs),
    )
  }

  const handleIndentKeyDown = (e: ReactKeyboardEvent<HTMLElement>, marker: IndentMarker) => {
    if (!onIndentsChange || (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight')) return
    e.preventDefault()
    const delta = (e.key === 'ArrowRight' ? 1 : -1) * (e.shiftKey ? 720 : INDENT_SNAP_TWIPS)
    const abs = snapIndentTwips(markerPos[marker] + delta)
    onIndentsChange(applyIndentDrag(marker, indents, abs, body, dims.pageWidth).attrs)
  }

  const zoneRefs = useRef<{ left: HTMLElement | null; right: HTMLElement | null }>({
    left: null,
    right: null,
  })
  const marginAbs = (side: 'left' | 'right', value: number) =>
    side === 'left' ? value : dims.pageWidth - value
  const handleMarginMouseDown = (e: ReactMouseEvent<HTMLElement>, side: 'left' | 'right') => {
    if (!onMarginsChange) return
    const orig = side === 'left' ? dims.marginLeft : dims.marginRight
    const marginAt = (abs: number) => clampMarginDrag(side, abs, dims)
    startMarkerDrag(
      e,
      marginAbs(side, orig),
      (abs) => {
        const m = marginAt(abs)
        const zone = zoneRefs.current[side]
        if (!zone) return
        if (side === 'left') zone.style.width = `${twipsToPx(m)}px`
        else {
          zone.style.left = `${width - twipsToPx(m)}px`
          zone.style.width = `${twipsToPx(m)}px`
        }
      },
      (abs) => {
        const m = marginAt(abs)
        if (m === orig) return
        onMarginsChange({
          left: side === 'left' ? m : dims.marginLeft,
          right: side === 'right' ? m : dims.marginRight,
        })
      },
    )
  }

  const handleMarginKeyDown = (e: ReactKeyboardEvent<HTMLElement>, side: 'left' | 'right') => {
    if (!onMarginsChange || (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight')) return
    e.preventDefault()
    const orig = side === 'left' ? dims.marginLeft : dims.marginRight
    const delta = (e.key === 'ArrowRight' ? 1 : -1) * (e.shiftKey ? 720 : INDENT_SNAP_TWIPS)
    const m = clampMarginDrag(side, snapIndentTwips(marginAbs(side, orig) + delta), dims)
    if (m === orig) return
    onMarginsChange({
      left: side === 'left' ? m : dims.marginLeft,
      right: side === 'right' ? m : dims.marginRight,
    })
  }

  // Default tab stop markers (light gray) when no custom stops mark a position
  const defaultStops: number[] = []
  if (!stops.some(isRenderableTabStop)) {
    const contentWidth = dims.pageWidth - dims.marginLeft - dims.marginRight
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
      pos: measure.format(stop.pos),
    }) + (stop.leader ? t('appTabLeader', { leader: stop.leader }) : '')

  const clampPx = (twips: number) => Math.min(Math.max(twipsToPx(twips), 0), width)
  const indentMarker = (marker: IndentMarker) => (
    <span
      key={marker}
      ref={(el) => {
        markerRefs.current[marker] = el
      }}
      className={`ruler-ind ruler-ind-${marker}`}
      style={{ left: clampPx(markerPos[marker]) }}
      data-tip={t(INDENT_MARKER_KEYS[marker])}
      role="slider"
      tabIndex={0}
      aria-label={t(INDENT_MARKER_KEYS[marker])}
      aria-valuemin={0}
      aria-valuemax={dims.pageWidth}
      aria-valuenow={Math.round(markerPos[marker])}
      onMouseDown={(e) => handleIndentMouseDown(e, marker)}
      onKeyDown={(e) => handleIndentKeyDown(e, marker)}
      onClick={(e) => e.stopPropagation()}
    />
  )
  const marginHandle = (side: 'left' | 'right') => (
    <span
      className="ruler-margin-handle"
      data-side={side}
      style={{ left: side === 'left' ? marginLeft : width - marginRight }}
      role="slider"
      tabIndex={0}
      aria-label={t(side === 'left' ? 'appRulerLeftMargin' : 'appRulerRightMargin')}
      aria-valuemin={0}
      aria-valuemax={dims.pageWidth}
      aria-valuenow={side === 'left' ? dims.marginLeft : dims.marginRight}
      onMouseDown={(e) => handleMarginMouseDown(e, side)}
      onKeyDown={(e) => handleMarginKeyDown(e, side)}
      onClick={(e) => e.stopPropagation()}
    />
  )

  return (
    <div
      className="ruler"
      role="group"
      aria-label={t('appRulerLabel')}
      style={{ width }}
      onClick={handleRulerClick}
    >
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

      <div
        className="ruler-zone"
        ref={(el) => {
          zoneRefs.current.left = el
        }}
        style={{ left: 0, width: marginLeft }}
      />
      <div
        className="ruler-zone"
        ref={(el) => {
          zoneRefs.current.right = el
        }}
        style={{ left: width - marginRight, width: marginRight }}
      />

      {ticks.map((tk) =>
        tk.kind === 'num' ? (
          <span key={tk.pos} className="ruler-num" style={{ left: twipsToPx(tk.pos) }}>
            {tk.label}
          </span>
        ) : (
          <span
            key={tk.pos}
            className={tk.kind === 'mid' ? 'ruler-tick ruler-tick-mid' : 'ruler-tick'}
            style={{ left: twipsToPx(tk.pos) }}
          />
        ),
      )}

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

      {/* Word: the margin boundary drags only when the pointer is exactly on
          it and the caret is in the body (a table cell shows column markers
          there instead); indent markers sit above and win the hit test. */}
      {onMarginsChange && !body.inCell && marginHandle('left')}
      {onMarginsChange && !body.inCell && marginHandle('right')}
      {onIndentsChange && (['firstLine', 'hanging', 'left', 'right'] as const).map(indentMarker)}
    </div>
  )
}
