/**
 * PowerPoint's 3D rotation handle: a control at the selected model's centre. Dragging it turns
 * the model live over its picture (horizontal drag about y, vertical about x); releasing stores
 * the rotation and the model drawn at it, one undo step.
 */
import { useEffect, useRef, useState } from 'react'
import type { PictureRenderNode } from '@genoffice/pptx-render'
import type { Model3dRotation } from '@genoffice/pptx-engine/model3d'
import { t } from '../i18n/locale'
import { Icon3d } from '../components/icons'
import { loadModel, type ModelView } from './scene'

/** am3d:rot units per pixel dragged: half a degree */
const UNITS_PER_PX = 30000

interface Props {
  node: PictureRenderNode & { model3d: Model3dRotation }
  slideIndex: number
  onTurn: (rot: Model3dRotation, view: ModelView) => Promise<unknown>
}

export function Model3dTurnHandle({ node, slideIndex, onTurn }: Props) {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const viewRef = useRef<Promise<ModelView | null> | null>(null)
  const dragRef = useRef<{
    x: number
    y: number
    from: Model3dRotation
    rot: Model3dRotation
  } | null>(null)
  const [turning, setTurning] = useState(false)

  useEffect(
    () => () => {
      void viewRef.current?.then((view) => view?.dispose())
      viewRef.current = null
    },
    [node.sourceId, slideIndex],
  )

  const view = () =>
    (viewRef.current ??= window.slidesApi
      .getModel3dData(slideIndex, node.sourceId)
      .then((bytes) => (bytes && canvasRef.current ? loadModel(bytes, canvasRef.current) : null))
      .catch(() => null))

  const onPointerDown = (e: React.PointerEvent) => {
    e.stopPropagation()
    e.currentTarget.setPointerCapture(e.pointerId)
    dragRef.current = { x: e.clientX, y: e.clientY, from: node.model3d, rot: node.model3d }
    void view().then((v) => {
      if (!v || !dragRef.current) return
      v.draw(node.box.w, node.box.h, dragRef.current.rot)
      setTurning(true)
    })
  }

  const onPointerMove = (e: React.PointerEvent) => {
    const drag = dragRef.current
    if (!drag) return
    drag.rot = {
      ax: drag.from.ax + (e.clientY - drag.y) * UNITS_PER_PX,
      ay: drag.from.ay + (e.clientX - drag.x) * UNITS_PER_PX,
      az: drag.from.az,
    }
    void view().then((v) => v?.draw(node.box.w, node.box.h, drag.rot))
  }

  const onPointerUp = async () => {
    const drag = dragRef.current
    dragRef.current = null
    const v = await view()
    if (drag && v && (drag.rot.ax !== drag.from.ax || drag.rot.ay !== drag.from.ay)) {
      await onTurn(drag.rot, v)
    }
    setTurning(false)
  }

  return (
    <div
      className={`m3d-turn${turning ? ' turning' : ''}`}
      style={{ left: node.box.x, top: node.box.y, width: node.box.w, height: node.box.h }}
    >
      <canvas ref={canvasRef} />
      <button
        className="m3d-turn-handle"
        data-tip={t('ribbonModel3dTurnTip')}
        aria-label={t('ribbonModel3dTurnTip')}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={() => void onPointerUp()}
        onPointerCancel={() => void onPointerUp()}
      >
        <Icon3d size={20} />
      </button>
    </div>
  )
}
