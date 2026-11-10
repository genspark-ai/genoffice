import { useRef, useState } from 'react'
import type { CSSProperties, PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import { pdfRectToCss, viewToPdf } from './annotations'
import type { PageGeom } from './annotations'
import type { ImageEditInput, PageImageRef, StaticFormFillRecord } from '../shared/ipc'

type Rect = [number, number, number, number]

export interface LocalImageEdit {
  id: string
  input: ImageEditInput
  /** Ghost PNG (base64) of a transformed existing image; null while loading/unavailable */
  png?: string | null
  /** Z-band the existing image had before this op (labels the layer toggle) */
  origAbove?: boolean
  /** Editable source metadata when this image is a static form fill. */
  staticFill?: StaticFormFillRecord
  /** Pre-transparency pixels of the op's current bytes: transparency presets re-bake
      from this so they set an absolute level instead of compounding. Cleared by any
      other pixel edit (flip/crop/cutout/replace/insert-rotate). */
  opacityBase?: string
}

/** Existing images are addressed by their listed rect; stable across renders, not saves */
export const imageRectKey = (r: readonly number[]): string => r.map((v) => v.toFixed(2)).join(',')

type DragTarget = { kind: 'edit'; id: string } | { kind: 'existing'; ref: PageImageRef }

const targetKey = (t: DragTarget): string =>
  t.kind === 'edit' ? `e:${t.id}` : `x:${imageRectKey(t.ref.rect)}`

interface Box {
  left: number
  top: number
  width: number
  height: number
}

/**
 * Image edit layer: previews pending content-stream image ops (insert ghosts, moved
 * images, delete markers) and, in edit-image mode, overlays hit targets on the page's
 * existing images. Drag moves, corner handles resize (aspect kept), click selects.
 * The original pixels of touched images are erased by App's live page preview patch.
 * Below-text previews render in a separate blend layer so page ink shows through,
 * matching how the saved PDF will paint text over the image.
 */
export function ImageEditLayer({
  geom,
  scale,
  edits,
  existing,
  selectedId,
  selectedKey,
  editHint,
  onSelectEdit,
  onSelectExisting,
  onRect,
  onExistingRect,
  existingPng,
  onExistingDragStart,
}: {
  geom: PageGeom
  scale: number
  /** Pending ops on this page */
  edits: LocalImageEdit[]
  /** Existing images on this page not claimed by a pending op (empty unless edit-image mode) */
  existing: PageImageRef[]
  selectedId: string | null
  /** Selected existing image (imageRectKey), when nothing pending targets it yet */
  selectedKey: string | null
  editHint: string
  onSelectEdit: (id: string, x: number, y: number) => void
  onSelectExisting: (ref: PageImageRef, x: number, y: number) => void
  /** Committed move/resize of a pending op's rect; omit to disable (read-only) */
  onRect?: (id: string, rect: Rect) => void
  /** Drag/resize of an untouched existing image (App turns it into a transform op) */
  onExistingRect?: (ref: PageImageRef, rect: Rect) => void
  /** Prefetched pixels of an untouched existing image; shown live while it is dragged/resized */
  existingPng?: (ref: PageImageRef) => string | undefined
  /** Fired when a drag starts on an untouched existing image (App prefetches its pixels) */
  onExistingDragStart?: (ref: PageImageRef) => void
}): ReactElement {
  const [drag, setDrag] = useState<{
    target: DragTarget
    rect: Rect
    from: [number, number]
    to: [number, number]
  } | null>(null)
  const [resize, setResize] = useState<{
    target: DragTarget
    rect: Rect
    corner: number
    start: Box
    box: Box
  } | null>(null)
  const layerRef = useRef<HTMLDivElement>(null)

  const css = (r: Rect): Box => pdfRectToCss(geom, r, scale)

  const commitRect = (target: DragTarget, rect: Rect) => {
    if (target.kind === 'edit') onRect?.(target.id, rect)
    else onExistingRect?.(target.ref, rect)
  }

  const select = (target: DragTarget, x: number, y: number) => {
    if (target.kind === 'edit') onSelectEdit(target.id, x, y)
    else onSelectExisting(target.ref, x, y)
  }

  const canDrag = (target: DragTarget) =>
    target.kind === 'edit' ? onRect !== undefined : onExistingRect !== undefined

  // ── Drag to move (client-space deltas; converted to a PDF-space shift on release) ──

  const dragDown = (e: ReactPointerEvent<HTMLElement>, target: DragTarget, rect: Rect) => {
    if (e.button !== 0) return
    e.preventDefault()
    e.stopPropagation()
    e.currentTarget.setPointerCapture(e.pointerId)
    if (target.kind === 'existing') onExistingDragStart?.(target.ref)
    setDrag({ target, rect, from: [e.clientX, e.clientY], to: [e.clientX, e.clientY] })
  }

  const dragMove = (e: ReactPointerEvent<HTMLElement>) => {
    if (drag && canDrag(drag.target)) setDrag({ ...drag, to: [e.clientX, e.clientY] })
  }

  const dragUp = (e: ReactPointerEvent<HTMLElement>) => {
    const d = drag
    setDrag(null)
    if (!d) return
    if (!canDrag(d.target) || Math.hypot(d.to[0] - d.from[0], d.to[1] - d.from[1]) < 3) {
      select(d.target, e.clientX, e.clientY)
      return
    }
    const [ax, ay] = viewToPdf(geom, 0, 0)
    const [bx, by] = viewToPdf(geom, (d.to[0] - d.from[0]) / scale, (d.to[1] - d.from[1]) / scale)
    const dx = bx - ax
    const dy = by - ay
    commitRect(d.target, [d.rect[0] + dx, d.rect[1] + dy, d.rect[2] + dx, d.rect[3] + dy])
  }

  const dragStyle = (target: DragTarget, rect: Rect): CSSProperties => {
    if (resize && targetKey(resize.target) === targetKey(target)) return { ...resize.box }
    const base: CSSProperties = { ...css(rect) }
    if (drag && targetKey(drag.target) === targetKey(target)) {
      base.transform = `translate(${drag.to[0] - drag.from[0]}px, ${drag.to[1] - drag.from[1]}px)`
    }
    return base
  }

  /** dragStyle plus a pending quarter-turn rotation: odd turns draw the ghost with
      pre-rotation dims centered in the footprint, then CSS-rotate into place */
  const ghostStyle = (target: DragTarget, rect: Rect, turns: number): CSSProperties => {
    const s = dragStyle(target, rect)
    if (!turns) return s
    const rot = `rotate(${turns * 90}deg)`
    const transform = `${(s.transform as string) ?? ''} ${rot}`.trim()
    if (turns % 2 === 0) return { ...s, transform }
    const w = s.width as number
    const h = s.height as number
    return {
      ...s,
      left: (s.left as number) + w / 2 - h / 2,
      top: (s.top as number) + h / 2 - w / 2,
      width: h,
      height: w,
      transform,
    }
  }

  // ── Corner-handle resize (aspect kept, opposite corner fixed — same as DrawLayer) ──

  const handleDown = (
    e: ReactPointerEvent<HTMLElement>,
    target: DragTarget,
    rect: Rect,
    corner: number,
  ) => {
    if (e.button !== 0 || !canDrag(target)) return
    e.preventDefault()
    e.stopPropagation()
    e.currentTarget.setPointerCapture(e.pointerId)
    if (target.kind === 'existing') onExistingDragStart?.(target.ref)
    const start = css(rect)
    setResize({ target, rect, corner, start, box: start })
  }

  const handleMove = (e: ReactPointerEvent<HTMLElement>) => {
    if (!resize || !layerRef.current) return
    const layerBox = layerRef.current.getBoundingClientRect()
    const px = e.clientX - layerBox.left
    const py = e.clientY - layerBox.top
    const { start, corner } = resize
    const fx = corner % 2 === 0 ? start.left + start.width : start.left
    const fy = corner < 2 ? start.top + start.height : start.top
    const k = Math.max(
      Math.abs(px - fx) / start.width,
      Math.abs(py - fy) / start.height,
      16 / Math.max(start.width, start.height),
    )
    const w = start.width * k
    const h = start.height * k
    setResize({
      ...resize,
      box: {
        left: corner % 2 === 0 ? fx - w : fx,
        top: corner < 2 ? fy - h : fy,
        width: w,
        height: h,
      },
    })
  }

  const handleUp = () => {
    const r = resize
    setResize(null)
    if (!r) return
    if (Math.abs(r.box.width - r.start.width) < 2 && Math.abs(r.box.height - r.start.height) < 2)
      return
    const [ax, ay] = viewToPdf(geom, r.box.left / scale, r.box.top / scale)
    const [bx, by] = viewToPdf(
      geom,
      (r.box.left + r.box.width) / scale,
      (r.box.top + r.box.height) / scale,
    )
    commitRect(r.target, [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)])
  }

  // Selected target for the corner handles (must live on this page)
  const selectedTarget: { target: DragTarget; rect: Rect } | null = (() => {
    if (selectedId) {
      const e = edits.find((x) => x.id === selectedId)
      if (e && e.input.kind !== 'deleteImage')
        return { target: { kind: 'edit', id: e.id }, rect: e.input.rect }
    }
    if (selectedKey) {
      const ref = existing.find((x) => imageRectKey(x.rect) === selectedKey)
      if (ref) return { target: { kind: 'existing', ref }, rect: ref.rect }
    }
    return null
  })()

  const pointerProps = (target: DragTarget, rect: Rect) => ({
    onPointerDown: (e: ReactPointerEvent<HTMLElement>) => dragDown(e, target, rect),
    onPointerMove: dragMove,
    onPointerUp: dragUp,
    onPointerCancel: () => setDrag(null),
  })

  /** Will the edit end up painted below the page's text/ink after save? */
  const isBelow = (e: LocalImageEdit): boolean => {
    if (e.input.kind === 'deleteImage') return false
    if (e.input.kind === 'insertImage') return e.input.layer === 'belowText'
    // transform/replace: explicit layer wins, else the image's original z-band
    return (e.input.layer ?? (e.origAbove ? 'aboveText' : 'belowText')) === 'belowText'
  }

  const underImgs: ReactElement[] = []
  // Hit targets of untouched existing images; rendered below the pending-op elements so
  // an invisible page-sized image cannot swallow clicks meant for delete markers/ghosts
  const existingEls: ReactElement[] = []
  const mainEls: ReactElement[] = []
  for (const e of edits) {
    const sel = e.id === selectedId
    const target: DragTarget = { kind: 'edit', id: e.id }
    // Deleted images vanish via the live page preview (Acrobat/WPS behavior):
    // no marker is painted, restore is the undo stack
    if (e.input.kind === 'deleteImage') continue
    if (
      (e.input.kind === 'transformImage' || e.input.kind === 'replaceImage') &&
      imageRectKey(e.input.oldRect) !== imageRectKey(e.input.rect)
    ) {
      mainEls.push(
        <div key={`${e.id}-veil`} className="pdf-imgedit-veil" style={css(e.input.oldRect)} />,
      )
    }
    if (e.input.kind === 'insertImage' || e.input.kind === 'replaceImage' || e.png) {
      const src =
        e.input.kind === 'insertImage' || e.input.kind === 'replaceImage' ? e.input.image : e.png
      const turns =
        e.input.kind === 'transformImage' || e.input.kind === 'replaceImage'
          ? (((e.input.quarterTurns ?? 0) % 4) + 4) % 4
          : 0
      const img = (
        <img
          key={e.id}
          className={`pdf-imgedit-img${sel ? ' pdf-imgedit-selected' : ''}`}
          src={`data:image/png;base64,${src}`}
          style={ghostStyle(target, e.input.rect, turns)}
          data-tip={editHint}
          alt=""
          draggable={false}
          {...pointerProps(target, e.input.rect)}
        />
      )
      if (isBelow(e)) underImgs.push(img)
      else mainEls.push(img)
    } else {
      mainEls.push(
        <div
          key={e.id}
          className={`pdf-imgedit-ghostbox${sel ? ' pdf-imgedit-selected' : ''}`}
          style={dragStyle(target, e.input.rect)}
          data-tip={editHint}
          {...pointerProps(target, e.input.rect)}
        />,
      )
    }
  }

  // Big images (page-sized backgrounds) go first = underneath, so their hit target
  // cannot swallow clicks meant for the smaller pictures sitting on top of them
  const byAreaDesc = [...existing].sort(
    (a, b) =>
      (b.rect[2] - b.rect[0]) * (b.rect[3] - b.rect[1]) -
      (a.rect[2] - a.rect[0]) * (a.rect[3] - a.rect[1]),
  )
  // Same-rect duplicates (repeated watermark strips) need the occurrence index to keep
  // React keys unique, or whole hit targets get dropped
  const seenKeys = new Map<string, number>()
  for (const ref of byAreaDesc) {
    const rectKey = imageRectKey(ref.rect)
    const dup = seenKeys.get(rectKey) ?? 0
    seenKeys.set(rectKey, dup + 1)
    const key = dup === 0 ? rectKey : `${rectKey}#${dup}`
    const target: DragTarget = { kind: 'existing', ref }
    // While an untouched image is being dragged/resized, paint its prefetched pixels
    // in the moving box so the picture follows the hand (the raster copy underneath
    // is only erased by the live page patch after release)
    const active =
      (drag !== null && targetKey(drag.target) === targetKey(target)) ||
      (resize !== null && targetKey(resize.target) === targetKey(target))
    const png = active ? existingPng?.(ref) : undefined
    if (png) {
      const img = (
        <img
          key={`${key}-live`}
          className="pdf-imgedit-liveghost"
          src={`data:image/png;base64,${png}`}
          style={dragStyle(target, ref.rect)}
          alt=""
          draggable={false}
        />
      )
      if (ref.aboveText) mainEls.push(img)
      else underImgs.push(img)
    }
    existingEls.push(
      <div
        key={key}
        className={`pdf-imgedit-hit${rectKey === selectedKey ? ' pdf-imgedit-selected' : ''}`}
        style={dragStyle(target, ref.rect)}
        data-tip={editHint}
        {...pointerProps(target, ref.rect)}
      />,
    )
  }

  return (
    <>
      {underImgs.length > 0 && <div className="pdf-imgedit-under">{underImgs}</div>}
      <div ref={layerRef} className="pdf-imgedit-layer">
        {existingEls}
        {mainEls}
        {selectedTarget &&
          !drag &&
          canDrag(selectedTarget.target) &&
          (() => {
            const box =
              resize && targetKey(resize.target) === targetKey(selectedTarget.target)
                ? resize.box
                : css(selectedTarget.rect)
            const corners: [number, number, string][] = [
              [box.left, box.top, 'nwse-resize'],
              [box.left + box.width, box.top, 'nesw-resize'],
              [box.left, box.top + box.height, 'nesw-resize'],
              [box.left + box.width, box.top + box.height, 'nwse-resize'],
            ]
            return corners.map(([cx, cy, cursor], i) => (
              <div
                key={i}
                className="pdf-imgedit-handle"
                style={{ left: cx - 5, top: cy - 5, cursor }}
                onPointerDown={(e) => handleDown(e, selectedTarget.target, selectedTarget.rect, i)}
                onPointerMove={handleMove}
                onPointerUp={handleUp}
                onPointerCancel={() => setResize(null)}
              />
            ))
          })()}
      </div>
    </>
  )
}
