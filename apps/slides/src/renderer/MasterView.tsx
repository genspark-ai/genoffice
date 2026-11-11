/**
 * Slide master editing view (modeled on PowerPoint View → Slide Master).
 * Full-screen overlay: [master, ...layouts] thumbnails on the left, editable canvas for the
 * current part on the right. All edits go through slides:master-* IPC; App only mounts it and
 * refreshes all slides after exit.
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { RenderFill, RenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import { SlideCanvas } from './SlideCanvas'
import { SlideThumb } from './SlideThumb'
import { TextEditOverlay } from './TextEditOverlay'
import { armColorInput } from './color-input'
import { ContextMenu, type CtxItem } from './components/ContextMenu'
import { ShapeGalleryContent } from './components/ShapeGalleryPopover'
import { IconShapes, IconTextBox } from './components/icons'
import { defaultShapeStyle } from './default-shape'
import { isLineDrawKind, type DrawRect } from './draw-shape'
import { textBoxInsertSpec } from './textbox-insert'
import type { EditParagraph, MasterPartItem } from '../shared/ipc'
import { t } from './i18n/locale'
import { setImageSource } from './image-loader'

interface Props {
  initialItems: MasterPartItem[]
  onClose: () => void
}

export function MasterView({ initialItems, onClose }: Props) {
  const [items, setItems] = useState(initialItems)
  const [sel, setSel] = useState(0)
  const [selectedIds, setSelectedIds] = useState<string[]>([])
  const [editing, setEditing] = useState<{
    sourceId: string
    caret?: { x: number; y: number }
  } | null>(null)
  const [ctxMenu, setCtxMenu] = useState<{ x: number; y: number; targetId: string | null } | null>(
    null,
  )
  const [images, setImages] = useState<Map<string, HTMLImageElement>>(new Map())
  const [drawKind, setDrawKind] = useState<string | null>(null)
  const [shapesOpen, setShapesOpen] = useState(false)
  const colorInputRef = useRef<HTMLInputElement>(null)
  const colorTargetRef = useRef<{ kind: 'fill' | 'stroke'; sourceId: string } | null>(null)
  const drawKindRef = useRef<string | null>(null)
  const toolsRef = useRef<HTMLDivElement>(null)

  const slide = items[sel]?.slide

  const applyCurrent = useCallback(
    (updated: RenderSlide | null) => {
      if (updated)
        setItems((prev) => prev.map((it, i) => (i === sel ? { ...it, slide: updated } : it)))
    },
    [sel],
  )

  // Lazy image loading (picture elements + picture fills + background images; including the decoration layer)
  useEffect(() => {
    const urls = new Set<string>()
    const addFill = (fill: RenderFill | undefined) => {
      if (fill && fill.kind === 'image' && fill.dataUrl) urls.add(fill.dataUrl)
    }
    const walk = (nodes: readonly RenderNode[]) => {
      for (const n of nodes) {
        if (n.type === 'picture' && n.dataUrl) urls.add(n.dataUrl)
        if ((n.type === 'shape' || n.type === 'text') && n.fill) addFill(n.fill)
        if (n.type === 'group' && Array.isArray(n.children)) walk(n.children)
      }
    }
    for (const it of items) {
      addFill(it.slide.background)
      walk(it.slide.nodes)
    }
    urls.forEach((u) => {
      if (images.has(u)) return
      const img = new Image()
      img.onload = () => setImages((prev) => new Map(prev).set(u, img))
      setImageSource(img, u)
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [items])

  const selectPart = useCallback(
    async (i: number) => {
      if (i === sel) return
      const it = items[i]
      if (!it) return
      setEditing(null)
      setSelectedIds([])
      const rs = await window.slidesApi.masterOpen(it.partPath)
      if (rs) {
        setItems((prev) => prev.map((x, k) => (k === i ? { ...x, slide: rs } : x)))
        setSel(i)
      }
    },
    [sel, items],
  )

  const onTransform = useCallback(
    async (
      sourceId: string,
      box: { x: number; y: number; w: number; h: number; rotationDeg: number },
      preview?: boolean,
    ) => {
      applyCurrent(
        await window.slidesApi.masterEditTransform({
          sourceId,
          xPx: box.x,
          yPx: box.y,
          wPx: box.w,
          hPx: box.h,
          rotationDeg: box.rotationDeg,
          fitWidthPx: 1280,
          ...(preview ? { preview: true } : {}),
        }),
      )
    },
    [applyCurrent],
  )

  const commitEdit = useCallback(
    async (paragraphs: EditParagraph[]) => {
      if (!editing) return
      applyCurrent(
        await window.slidesApi.masterEditText({ sourceId: editing.sourceId, paragraphs }),
      )
      setEditing(null)
      setSelectedIds([editing.sourceId])
    },
    [editing, applyCurrent],
  )

  const deleteSelected = useCallback(async () => {
    for (const id of selectedIds) {
      applyCurrent(await window.slidesApi.masterDeleteElement({ sourceId: id }))
    }
    setSelectedIds([])
  }, [selectedIds, applyCurrent])

  const pickColor = useCallback((kind: 'fill' | 'stroke', sourceId: string) => {
    colorTargetRef.current = { kind, sourceId }
    if (colorInputRef.current) {
      armColorInput(colorInputRef.current)
      colorInputRef.current.click()
    }
  }, [])

  const onColorPicked = useCallback(
    async (hex: string) => {
      const target = colorTargetRef.current
      if (!target) return
      colorTargetRef.current = null
      applyCurrent(
        target.kind === 'fill'
          ? await window.slidesApi.masterEditFill({ sourceId: target.sourceId, fill: hex })
          : await window.slidesApi.masterEditStroke({
              sourceId: target.sourceId,
              stroke: { color: hex, widthPt: 1 },
            }),
      )
    },
    [applyCurrent],
  )

  const setDraw = useCallback((kind: string | null) => {
    drawKindRef.current = kind
    setDrawKind(kind)
  }, [])

  // Draw-mode commit: text box / shape / line onto the current part. A text box
  // opens straight into typing like the slide path; shapes take the install's
  // default shape style, lines a plain black stroke.
  const commitDraw = useCallback(
    async (rect: DrawRect) => {
      const kind = drawKindRef.current
      setDraw(null)
      if (!kind || !slide) return
      const box = kind === 'textbox' ? textBoxInsertSpec(rect) : null
      const spec = box ?? rect
      const r = await window.slidesApi.masterAddElement({
        kind,
        xPx: Math.round(spec.x),
        yPx: Math.round(spec.y),
        wPx: Math.round(spec.w),
        hPx: Math.round(spec.h),
        fitWidthPx: 1280,
        ...(isLineDrawKind(kind)
          ? { stroke: { color: '#000000', widthPt: 1 } }
          : kind === 'textbox'
            ? { bodyPr: box!.bodyPr }
            : defaultShapeStyle()),
      })
      if (!r) return
      applyCurrent(r.slide)
      setSelectedIds([r.sourceId])
      if (kind === 'textbox') setEditing({ sourceId: r.sourceId })
    },
    [applyCurrent, setDraw, slide],
  )

  // Close the shapes dropdown on any click outside the toolbar
  useEffect(() => {
    if (!shapesOpen) return
    const onDown = (ev: MouseEvent) => {
      if (toolsRef.current && !toolsRef.current.contains(ev.target as Node)) setShapesOpen(false)
    }
    document.addEventListener('mousedown', onDown)
    return () => document.removeEventListener('mousedown', onDown)
  }, [shapesOpen])

  // Delete/Backspace deletes selected elements (not intercepted while editing)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (editing || !selectedIds.length) return
      const el = e.target as HTMLElement | null
      if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable))
        return
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault()
        void deleteSelected()
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [editing, selectedIds, deleteSelected])

  const editNode = useMemo(() => {
    if (!editing || !slide) return null
    return (
      (slide.nodes.find((n) => n.sourceId === editing.sourceId) as ShapeRenderNode | undefined) ??
      null
    )
  }, [editing, slide])

  const ctxItems = useMemo<Array<CtxItem | null>>(() => {
    if (!ctxMenu?.targetId) return []
    const id = ctxMenu.targetId
    const node = slide?.nodes.find((n) => n.sourceId === id)
    const textual = node && (node.type === 'text' || node.type === 'shape')
    return [
      ...(textual
        ? [{ label: t('appCtxEditText'), onClick: () => setEditing({ sourceId: id }) } as CtxItem]
        : []),
      ...(textual
        ? [
            { label: t('appMasterFill'), onClick: () => pickColor('fill', id) },
            {
              label: t('appMasterFillNone'),
              onClick: () =>
                void window.slidesApi
                  .masterEditFill({ sourceId: id, fill: 'none' })
                  .then(applyCurrent),
            },
            { label: t('appMasterStroke'), onClick: () => pickColor('stroke', id) },
            {
              label: t('appMasterStrokeNone'),
              onClick: () =>
                void window.slidesApi
                  .masterEditStroke({ sourceId: id, stroke: null })
                  .then(applyCurrent),
            },
          ]
        : []),
      null,
      { label: t('appCtxDelete'), hint: '⌫', danger: true, onClick: () => void deleteSelected() },
    ]
  }, [ctxMenu, slide, pickColor, applyCurrent, deleteSelected])

  if (!slide) return null

  return (
    <div className="master-view">
      <div className="master-bar">
        <span className="master-bar-title">{t('appMasterTitle')}</span>
        <div className="master-tools" ref={toolsRef}>
          <button
            className={`master-tool ${drawKind === 'textbox' ? 'active' : ''}`}
            data-tip={t('appMasterInsertTextBox')}
            aria-label={t('appMasterInsertTextBox')}
            onClick={() => setDraw(drawKind === 'textbox' ? null : 'textbox')}
          >
            <IconTextBox size={16} />
          </button>
          <div className="master-tool-anchor">
            <button
              className={`master-tool ${drawKind && drawKind !== 'textbox' ? 'active' : ''}`}
              data-tip={t('appMasterInsertShape')}
              aria-label={t('appMasterInsertShape')}
              onClick={() => setShapesOpen((v) => !v)}
            >
              <IconShapes size={16} />
            </button>
            {shapesOpen && (
              <div className="master-tool-drop">
                <ShapeGalleryContent
                  includeLines
                  onPick={(prst) => {
                    setShapesOpen(false)
                    setDraw(prst)
                  }}
                />
              </div>
            )}
          </div>
        </div>
        <button className="master-close-btn" onClick={onClose}>
          {t('appMasterClose')}
        </button>
      </div>
      <div className="master-body">
        <div className="slide-list master-list">
          {items.map((it, i) => (
            <div
              key={it.partPath}
              className={`thumb master-thumb ${i === sel ? 'active' : ''} ${it.kind === 'layout' ? 'master-thumb-layout' : ''}`}
              onClick={() => void selectPart(i)}
              data-tip={it.kind === 'master' ? t('appMasterThumbMaster') : it.name}
            >
              <SlideThumb
                slide={it.slide}
                images={images}
                width={it.kind === 'layout' ? 108 : 126}
              />
              <span className="master-thumb-label">
                {it.kind === 'master'
                  ? t('appMasterThumbMaster')
                  : it.name || t('appMasterThumbLayout', { n: i })}
              </span>
            </div>
          ))}
        </div>
        <div className="master-stage">
          <div
            className="stage-rel"
            style={{ position: 'relative', width: slide.widthPx, height: slide.heightPx }}
          >
            <SlideCanvas
              slide={slide}
              selectedIds={selectedIds}
              onSelect={(id, additive) =>
                setSelectedIds((prev) => {
                  if (!id) return []
                  if (additive)
                    return prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
                  return [id]
                })
              }
              onEditText={(sourceId, caret) => setEditing({ sourceId, caret })}
              onTransform={onTransform}
              // Master edits have no snapshot undo to pop: settle the gesture on its start box
              onTransformCancel={(id, startBox) => void onTransform(id, startBox)}
              onEditTableCell={() => {}}
              onTableColResize={() => {}}
              onContextMenu={(sourceId, x, y) => {
                if (editing) return
                if (sourceId)
                  setSelectedIds((prev) => (prev.includes(sourceId) ? prev : [sourceId]))
                setCtxMenu({ x, y, targetId: sourceId })
              }}
              onMarqueeSelect={setSelectedIds}
              images={images}
              editingText={editing ? { sourceId: editing.sourceId } : null}
              drawMode={drawKind ? { kind: drawKind } : null}
              onDrawCommit={(rect) => void commitDraw(rect)}
              onDrawCancel={() => setDraw(null)}
            />
            {editing && editNode && (
              <TextEditOverlay
                node={editNode}
                scale={slide.scale}
                caretPoint={editing.caret}
                onCommit={commitEdit}
                onCancel={() => setEditing(null)}
              />
            )}
          </div>
        </div>
      </div>
      <input
        ref={colorInputRef}
        type="color"
        style={{ position: 'fixed', left: -9999, top: -9999 }}
        onChange={(e) => void onColorPicked(e.target.value.toUpperCase())}
      />
      {ctxMenu && (
        <ContextMenu
          x={ctxMenu.x}
          y={ctxMenu.y}
          items={ctxItems}
          onClose={() => setCtxMenu(null)}
        />
      )}
    </div>
  )
}
