import { useEffect, useRef, useState } from 'react'
import type {
  MouseEvent as ReactMouseEvent,
  PointerEvent as ReactPointerEvent,
  ReactElement,
} from 'react'
import { pdfRectToCss, viewToPdf } from './annotations'
import type { PageGeom } from './annotations'
import type { FormFieldInput, FormFieldKind } from '../shared/ipc'
import type { FormWidgetKind } from './form-catalog'
import type { LocalFormField } from './edit-state'

type Rect = FormFieldInput['rect']
/** Display-space box at scale 1 (points, y down) */
type Box = { left: number; top: number; width: number; height: number }

/** Ribbon tools: the field kinds plus a date preset (a text field with date actions) */
export type FormDesignTool = FormFieldKind | 'date'

/** Default widget sizes in PDF points when the user clicks instead of dragging (Acrobat-like). */
export const FORM_FIELD_DEFAULT_SIZE: Record<FormDesignTool, [number, number]> = {
  text: [160, 22],
  date: [110, 22],
  checkbox: [14, 14],
  radio: [14, 14],
  choice: [160, 22],
  signature: [160, 44],
}

export const DATE_FORMATS = ['yyyy-mm-dd', 'mm/dd/yyyy', 'dd/mm/yyyy', 'yyyy/mm/dd'] as const

const MIN_PTS = 6
const CLICK_PX = 4
const SNAP_PX = 5
/** Arrow-key nudge in points (Shift: ×10) */
export const NUDGE_PTS = 1

/** A field the design layer can select, move, resize and delete: either authored this session or already in the file. */
export interface DesignField {
  id: string
  kind: FormWidgetKind
  name: string
  pageIndex: number
  rect: Rect
  required: boolean
  existing: boolean
  pending?: LocalFormField
}

export const widgetDesignId = (widgetId: string): string => `widget:${widgetId}`
export const designWidgetId = (id: string): string | null =>
  id.startsWith('widget:') ? id.slice('widget:'.length) : null

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'
const HANDLES: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']

const boxToRect = (geom: PageGeom, b: Box): Rect => {
  const [ax, ay] = viewToPdf(geom, b.left, b.top)
  const [bx, by] = viewToPdf(geom, b.left + b.width, b.top + b.height)
  return [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)]
}

const clampBox = (b: Box, pageWidth: number, pageHeight: number): Box => {
  const width = Math.min(b.width, pageWidth)
  const height = Math.min(b.height, pageHeight)
  return {
    left: Math.min(Math.max(b.left, 0), pageWidth - width),
    top: Math.min(Math.max(b.top, 0), pageHeight - height),
    width,
    height,
  }
}

/** Keep a rect inside the page (the authored widget must not hang off the page). */
export function clampRectToPage(
  rect: Rect,
  geom: PageGeom,
  pageWidth: number,
  pageHeight: number,
): Rect {
  return boxToRect(geom, clampBox(pdfRectToCss(geom, rect, 1), pageWidth, pageHeight))
}

interface SnapTargets {
  xs: number[]
  ys: number[]
}

const snapValue = (v: number, targets: number[], tol: number): number | null => {
  let best: number | null = null
  for (const t of targets) {
    if (Math.abs(t - v) <= tol && (best === null || Math.abs(t - v) < Math.abs(best - v))) best = t
  }
  return best
}

/** Snap a dragged box's moving edges (and center while moving) to other fields and the page center. */
function snapBox(
  b: Box,
  handle: Handle | null,
  targets: SnapTargets,
  tol: number,
): { box: Box; guides: { x: number[]; y: number[] } } {
  const guides = { x: [] as number[], y: [] as number[] }
  const out = { ...b }
  const west = handle === null || handle.includes('w')
  const east = handle === null || handle.includes('e')
  const north = handle === null || handle.includes('n')
  const south = handle === null || handle.includes('s')
  const moving = handle === null

  const xCands: [number, (d: number) => void][] = []
  if (west)
    xCands.push([b.left, (d) => (moving ? (out.left += d) : ((out.left += d), (out.width -= d)))])
  if (east) xCands.push([b.left + b.width, (d) => (moving ? (out.left += d) : (out.width += d))])
  if (moving) xCands.push([b.left + b.width / 2, (d) => (out.left += d)])
  for (const [v, apply] of xCands) {
    const s = snapValue(v, targets.xs, tol)
    if (s !== null) {
      apply(s - v)
      guides.x.push(s)
      break
    }
  }
  const yCands: [number, (d: number) => void][] = []
  if (north)
    yCands.push([b.top, (d) => (moving ? (out.top += d) : ((out.top += d), (out.height -= d)))])
  if (south) yCands.push([b.top + b.height, (d) => (moving ? (out.top += d) : (out.height += d))])
  if (moving) yCands.push([b.top + b.height / 2, (d) => (out.top += d)])
  for (const [v, apply] of yCands) {
    const s = snapValue(v, targets.ys, tol)
    if (s !== null) {
      apply(s - v)
      guides.y.push(s)
      break
    }
  }
  return { box: out, guides }
}

function resizeBox(b: Box, handle: Handle | null, dx: number, dy: number): Box {
  if (handle === null) return { ...b, left: b.left + dx, top: b.top + dy }
  let { left, top, width, height } = b
  if (handle.includes('w')) {
    left += dx
    width -= dx
  } else if (handle.includes('e')) width += dx
  if (handle.includes('n')) {
    top += dy
    height -= dy
  } else if (handle.includes('s')) height += dy
  if (width < MIN_PTS) {
    if (handle.includes('w')) left = b.left + b.width - MIN_PTS
    width = MIN_PTS
  }
  if (height < MIN_PTS) {
    if (handle.includes('n')) top = b.top + b.height - MIN_PTS
    height = MIN_PTS
  }
  return { left, top, width, height }
}

export function FormDesignLayer({
  designing,
  preview,
  tool,
  geom,
  scale,
  pageWidth,
  pageHeight,
  fields,
  selectedIds,
  onCreate,
  onSelect,
  onRects,
  onContextMenu,
}: {
  /** Form Design tab is active: fields are selectable/movable */
  designing: boolean
  /** Preview inside Form Design: pending fields render as fillable mock controls, existing ones fall through to the fill layer */
  preview: boolean
  /** Field kind to draw; null = select/move only */
  tool: FormDesignTool | null
  geom: PageGeom
  scale: number
  pageWidth: number
  pageHeight: number
  fields: DesignField[]
  selectedIds: ReadonlySet<string>
  /** `keep` = the user held Ctrl/⌘ to keep the tool armed */
  onCreate: (tool: FormDesignTool, rect: Rect, keep: boolean) => void
  /** Plain click replaces the selection; Shift/Ctrl/⌘ toggles; a marquee passes every hit (replace) */
  onSelect: (ids: string[], mode: 'replace' | 'toggle') => void
  /** One drag end, possibly moving every selected field on this page (one undo step) */
  onRects: (changes: { id: string; rect: Rect }[]) => void
  onContextMenu: (e: ReactMouseEvent<HTMLElement>, fieldId: string | null) => void
}): ReactElement | null {
  const [live, setLive] = useState<Box | null>(null)
  const drawStart = useRef<{ view: [number, number]; px: [number, number] } | null>(null)
  const drag = useRef<{
    id: string
    box: Box
    /** Other selected fields on this page that move along with a body drag */
    others: { id: string; box: Box }[]
    handle: Handle | null
    from: [number, number]
    moved: boolean
    toggle: boolean
  } | null>(null)
  const [dragBoxes, setDragBoxes] = useState<Map<string, Box> | null>(null)
  const marquee = useRef<{ view: [number, number]; px: [number, number] } | null>(null)
  const [marqueeBox, setMarqueeBox] = useState<Box | null>(null)
  const [guides, setGuides] = useState<{ x: number[]; y: number[] } | null>(null)

  useEffect(() => {
    if (!designing) {
      drawStart.current = null
      drag.current = null
      marquee.current = null
      setLive(null)
      setDragBoxes(null)
      setMarqueeBox(null)
      setGuides(null)
    }
  }, [designing])

  if (!designing && fields.every((f) => f.existing)) return null

  const layerToView = (e: ReactPointerEvent, layer: Element): [number, number] => {
    const box = layer.getBoundingClientRect()
    return [(e.clientX - box.left) / scale, (e.clientY - box.top) / scale]
  }
  const viewBox = (a: [number, number], b: [number, number]): Box => ({
    left: Math.min(a[0], b[0]),
    top: Math.min(a[1], b[1]),
    width: Math.abs(b[0] - a[0]),
    height: Math.abs(b[1] - a[1]),
  })
  const snapTargets = (exclude: ReadonlySet<string>): SnapTargets => {
    const xs = [0, pageWidth / 2, pageWidth]
    const ys = [0, pageHeight / 2, pageHeight]
    for (const f of fields) {
      if (exclude.has(f.id)) continue
      const b = pdfRectToCss(geom, f.rect, 1)
      xs.push(b.left, b.left + b.width, b.left + b.width / 2)
      ys.push(b.top, b.top + b.height, b.top + b.height / 2)
    }
    return { xs, ys }
  }
  const emit = (b: Box) => boxToRect(geom, clampBox(b, pageWidth, pageHeight))

  const intersects = (a: Box, b: Box) =>
    a.left < b.left + b.width &&
    a.left + a.width > b.left &&
    a.top < b.top + b.height &&
    a.top + a.height > b.top

  const down = (e: ReactPointerEvent<HTMLElement>) => {
    if (e.button !== 0) return
    e.preventDefault()
    e.currentTarget.setPointerCapture(e.pointerId)
    const start = {
      view: layerToView(e, e.currentTarget),
      px: [e.clientX, e.clientY] as [number, number],
    }
    if (!tool) {
      // Empty-space press: a drag rubber-bands a selection, a click clears it
      marquee.current = start
      return
    }
    drawStart.current = start
  }
  const move = (e: ReactPointerEvent<HTMLElement>) => {
    const m = marquee.current
    if (m) {
      if (Math.hypot(e.clientX - m.px[0], e.clientY - m.px[1]) >= CLICK_PX)
        setMarqueeBox(viewBox(m.view, layerToView(e, e.currentTarget)))
      return
    }
    const start = drawStart.current
    if (!start || !tool) return
    setLive(viewBox(start.view, layerToView(e, e.currentTarget)))
  }
  const up = (e: ReactPointerEvent<HTMLElement>) => {
    const m = marquee.current
    if (m) {
      marquee.current = null
      const box = marqueeBox
      setMarqueeBox(null)
      if (e.type !== 'pointerup') return
      if (!box) {
        onSelect([], 'replace')
        return
      }
      const hits = fields
        .filter((f) => intersects(box, pdfRectToCss(geom, f.rect, 1)))
        .map((f) => f.id)
      onSelect(hits, e.shiftKey || e.ctrlKey || e.metaKey ? 'toggle' : 'replace')
      return
    }
    const start = drawStart.current
    drawStart.current = null
    setLive(null)
    if (!start || !tool || e.type !== 'pointerup') return
    const end = layerToView(e, e.currentTarget)
    const dragged = Math.hypot(e.clientX - start.px[0], e.clientY - start.px[1]) >= CLICK_PX
    let box: Box
    if (dragged) {
      box = viewBox(start.view, end)
      if (box.width < MIN_PTS || box.height < MIN_PTS) return
    } else {
      // A click drops a default-size widget with its top-left corner at the pointer
      const [width, height] = FORM_FIELD_DEFAULT_SIZE[tool]
      box = { left: start.view[0], top: start.view[1], width, height }
    }
    onCreate(tool, emit(box), e.ctrlKey || e.metaKey)
  }
  const cancel = () => {
    drawStart.current = null
    marquee.current = null
    setLive(null)
    setMarqueeBox(null)
  }

  const fieldDown = (e: ReactPointerEvent<HTMLElement>, f: DesignField, handle: Handle | null) => {
    if (e.button !== 0 || !designing) return
    e.preventDefault()
    e.stopPropagation()
    const toggle = e.shiftKey || e.ctrlKey || e.metaKey
    // Pressing an already selected field keeps the group so it can be dragged together
    if (toggle) onSelect([f.id], 'toggle')
    else if (!selectedIds.has(f.id)) onSelect([f.id], 'replace')
    e.currentTarget.setPointerCapture(e.pointerId)
    const others =
      handle === null && selectedIds.has(f.id) && !toggle
        ? fields
            .filter((o) => o.id !== f.id && selectedIds.has(o.id))
            .map((o) => ({ id: o.id, box: pdfRectToCss(geom, o.rect, 1) }))
        : []
    drag.current = {
      id: f.id,
      box: pdfRectToCss(geom, f.rect, 1),
      others,
      handle,
      from: [e.clientX, e.clientY],
      moved: false,
      toggle,
    }
  }
  const fieldMove = (e: ReactPointerEvent<HTMLElement>) => {
    const d = drag.current
    if (!d) return
    const dxPx = e.clientX - d.from[0]
    const dyPx = e.clientY - d.from[1]
    if (!d.moved && Math.hypot(dxPx, dyPx) < CLICK_PX) return
    if (d.toggle) return
    d.moved = true
    const raw = resizeBox(d.box, d.handle, dxPx / scale, dyPx / scale)
    const moving = new Set([d.id, ...d.others.map((o) => o.id)])
    const snapped = e.altKey
      ? { box: raw, guides: { x: [], y: [] } }
      : snapBox(raw, d.handle, snapTargets(moving), SNAP_PX / scale)
    const lead = clampBox(snapped.box, pageWidth, pageHeight)
    const dx = lead.left - d.box.left
    const dy = lead.top - d.box.top
    const next = new Map<string, Box>([[d.id, lead]])
    for (const o of d.others) {
      next.set(
        o.id,
        clampBox({ ...o.box, left: o.box.left + dx, top: o.box.top + dy }, pageWidth, pageHeight),
      )
    }
    setDragBoxes(next)
    setGuides(snapped.guides)
  }
  const fieldUp = (e: ReactPointerEvent<HTMLElement>) => {
    const d = drag.current
    drag.current = null
    const result = dragBoxes
    setDragBoxes(null)
    setGuides(null)
    if (!d || !d.moved || !result || e.type !== 'pointerup') return
    onRects([...result].map(([id, box]) => ({ id, rect: emit(box) })))
  }

  const px = (v: number) => v * scale
  const cssBox = (b: Box) => ({
    left: px(b.left),
    top: px(b.top),
    width: px(b.width),
    height: px(b.height),
  })

  if (preview) {
    return (
      <div
        className="pdf-formdesign-layer is-preview"
        style={{ width: pageWidth * scale, height: pageHeight * scale, pointerEvents: 'none' }}
      >
        {fields
          .filter((f) => !f.existing)
          .map((f) => (
            <PreviewControl key={f.id} field={f} style={pdfRectToCss(geom, f.rect, scale)} />
          ))}
      </div>
    )
  }

  return (
    <div
      className={`pdf-formdesign-layer${tool ? ' is-drawing' : ''}`}
      style={{
        width: pageWidth * scale,
        height: pageHeight * scale,
        pointerEvents: designing ? 'auto' : 'none',
      }}
      onPointerDown={down}
      onPointerMove={move}
      onPointerUp={up}
      onPointerCancel={cancel}
      onContextMenu={(e) => {
        e.preventDefault()
        onContextMenu(e, null)
      }}
    >
      {fields.map((f) => {
        const dragged = dragBoxes?.get(f.id)
        const box = dragged ? cssBox(dragged) : pdfRectToCss(geom, f.rect, scale)
        const selected = selectedIds.has(f.id)
        const single = selected && selectedIds.size === 1
        return (
          <div
            key={f.id}
            className={`pdf-formdesign-field is-${f.kind}${selected ? ' is-selected' : ''}${
              f.existing ? ' is-existing' : ''
            }`}
            style={box}
            title={f.name}
            data-field-id={f.id}
            onPointerDown={(e) => fieldDown(e, f, null)}
            onPointerMove={fieldMove}
            onPointerUp={fieldUp}
            onPointerCancel={fieldUp}
            onContextMenu={(e) => {
              e.preventDefault()
              e.stopPropagation()
              if (!selectedIds.has(f.id)) onSelect([f.id], 'replace')
              onContextMenu(e, f.id)
            }}
          >
            <span className="pdf-formdesign-label">
              {f.name}
              {f.required ? ' *' : ''}
            </span>
            {single &&
              HANDLES.map((h) => (
                <span
                  key={h}
                  className={`pdf-formdesign-handle is-${h}`}
                  onPointerDown={(e) => fieldDown(e, f, h)}
                  onPointerMove={fieldMove}
                  onPointerUp={fieldUp}
                  onPointerCancel={fieldUp}
                />
              ))}
          </div>
        )
      })}
      {live && <div className="pdf-formdesign-field is-live" style={cssBox(live)} />}
      {marqueeBox && <div className="pdf-formdesign-marquee" style={cssBox(marqueeBox)} />}
      {guides?.x.map((x) => (
        <div key={`x${x}`} className="pdf-formdesign-guide is-v" style={{ left: px(x) }} />
      ))}
      {guides?.y.map((y) => (
        <div key={`y${y}`} className="pdf-formdesign-guide is-h" style={{ top: px(y) }} />
      ))}
    </div>
  )
}

/** Fillable stand-in for a not-yet-saved field so the author can try the form before saving. */
function PreviewControl({ field, style }: { field: DesignField; style: Box }): ReactElement {
  const input = field.pending?.input
  if (field.kind === 'checkbox') {
    return (
      <input
        type="checkbox"
        className="pdf-form-checkbox pdf-formdesign-preview"
        style={style}
        defaultChecked={!!input?.checked}
      />
    )
  }
  if (field.kind === 'radio') {
    return (
      <input
        type="radio"
        className="pdf-form-radio pdf-formdesign-preview"
        style={style}
        name={`preview:${field.name}`}
        defaultChecked={!!input?.checked}
      />
    )
  }
  if (field.kind === 'choice') {
    return (
      <select
        className="pdf-form-input pdf-formdesign-preview"
        style={style}
        defaultValue={input?.value ?? ''}
      >
        {(input?.options ?? []).map((o) => (
          <option key={o} value={o}>
            {o}
          </option>
        ))}
      </select>
    )
  }
  if (field.kind === 'signature') {
    return (
      <div className="pdf-form-signature pdf-formdesign-preview" style={style}>
        {field.name}
      </div>
    )
  }
  if (input?.dateFormat) {
    return (
      <input
        className="pdf-form-input pdf-formdesign-preview"
        style={style}
        placeholder={input.dateFormat}
        defaultValue={input.value ?? ''}
      />
    )
  }
  if (input?.multiLine) {
    return (
      <textarea
        className="pdf-form-input pdf-formdesign-preview"
        style={style}
        defaultValue={input.value ?? ''}
      />
    )
  }
  return (
    <input
      className="pdf-form-input pdf-formdesign-preview"
      style={style}
      defaultValue={input?.value ?? ''}
    />
  )
}

export interface FormFieldPropsLabels {
  name: string
  required: string
  multiLine: string
  defaultValue: string
  defaultChecked: string
  remove: string
  done: string
  cancel: string
  existingHint: string
  exportValue: string
  radioGroupHint: string
  options: string
  dateFormat: string
}

/** Floating property card under the selected field. The name commits on blur/Enter so a half-typed name never collides. */
export function FormFieldProps({
  field,
  isNew,
  style,
  labels,
  onPatch,
  onDelete,
  onDone,
  onCancel,
}: {
  field: DesignField
  /** Just placed: the name box takes focus and Cancel removes the field again */
  isNew: boolean
  style: { left: number; top: number }
  labels: FormFieldPropsLabels
  onPatch: (input: Partial<FormFieldInput>, coalesceKey?: string) => boolean
  onDelete: () => void
  onDone: () => void
  onCancel: () => void
}): ReactElement {
  const input = field.pending?.input
  const [name, setName] = useState(field.name)
  // Undo/redo can change the name underneath an open card; blur must not re-apply the old draft
  useEffect(() => setName(field.name), [field.name])
  const commitName = () => {
    const next = name.trim()
    if (next === field.name) return
    if (!onPatch({ name: next })) setName(field.name)
  }
  const stop = (e: { stopPropagation(): void }) => e.stopPropagation()
  return (
    <div
      className="pdf-formdesign-props"
      style={style}
      onPointerDown={stop}
      onClick={stop}
      onKeyDown={(e) => {
        if (e.key === 'Escape') (e.target as HTMLElement).blur()
        stop(e)
      }}
    >
      <label>
        <span>{labels.name}</span>
        <input
          value={name}
          spellCheck={false}
          autoFocus={isNew}
          disabled={field.existing}
          onFocus={(e) => isNew && e.target.select()}
          onChange={(e) => setName(e.target.value)}
          onBlur={commitName}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              commitName()
              onDone()
            }
          }}
        />
      </label>
      {field.existing && <span className="pdf-formdesign-hint">{labels.existingHint}</span>}
      {input?.kind === 'radio' && (
        <>
          <label>
            <span>{labels.exportValue}</span>
            <input
              value={input.exportValue ?? ''}
              spellCheck={false}
              onChange={(e) =>
                onPatch({ exportValue: e.target.value }, `formField:${field.id}:exportValue`)
              }
            />
          </label>
          <span className="pdf-formdesign-hint">{labels.radioGroupHint}</span>
        </>
      )}
      {input?.kind === 'choice' && (
        <label>
          <span>{labels.options}</span>
          <textarea
            value={(input.options ?? []).join('\n')}
            rows={4}
            spellCheck={false}
            onChange={(e) =>
              onPatch({ options: e.target.value.split('\n') }, `formField:${field.id}:options`)
            }
          />
        </label>
      )}
      {input?.kind === 'choice' && (
        <label>
          <span>{labels.defaultValue}</span>
          <select value={input.value ?? ''} onChange={(e) => onPatch({ value: e.target.value })}>
            <option value="">—</option>
            {(input.options ?? [])
              .map((o) => o.trim())
              .filter(Boolean)
              .map((o) => (
                <option key={o} value={o}>
                  {o}
                </option>
              ))}
          </select>
        </label>
      )}
      {input?.kind === 'text' && input.dateFormat && (
        <label>
          <span>{labels.dateFormat}</span>
          <select
            value={input.dateFormat}
            onChange={(e) => onPatch({ dateFormat: e.target.value })}
          >
            {DATE_FORMATS.map((f) => (
              <option key={f} value={f}>
                {f}
              </option>
            ))}
          </select>
        </label>
      )}
      {input?.kind === 'text' && !input.dateFormat && (
        <label>
          <span>{labels.defaultValue}</span>
          <input
            value={input.value ?? ''}
            onChange={(e) => onPatch({ value: e.target.value }, `formField:${field.id}:value`)}
          />
        </label>
      )}
      <label className="pdf-formdesign-check">
        <input
          type="checkbox"
          checked={field.required}
          onChange={(e) => onPatch({ required: e.target.checked })}
        />
        <span>{labels.required}</span>
      </label>
      {input?.kind === 'text' && !input.dateFormat && (
        <label className="pdf-formdesign-check">
          <input
            type="checkbox"
            checked={!!input.multiLine}
            onChange={(e) => onPatch({ multiLine: e.target.checked })}
          />
          <span>{labels.multiLine}</span>
        </label>
      )}
      {(input?.kind === 'checkbox' || input?.kind === 'radio') && (
        <label className="pdf-formdesign-check">
          <input
            type="checkbox"
            checked={!!input.checked}
            onChange={(e) => onPatch({ checked: e.target.checked })}
          />
          <span>{labels.defaultChecked}</span>
        </label>
      )}
      <div className="pdf-formdesign-actions">
        <button type="button" className="pdf-formdesign-remove" onClick={onDelete}>
          {labels.remove}
        </button>
        {isNew && (
          <button type="button" onClick={onCancel}>
            {labels.cancel}
          </button>
        )}
        <button type="button" className="pdf-formdesign-done" onClick={onDone}>
          {labels.done}
        </button>
      </div>
    </div>
  )
}
