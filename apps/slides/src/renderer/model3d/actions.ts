/**
 * 3D models in the editor: the renderer draws them (scene.ts), the main process stores the
 * rotation and the drawing as PowerPoint does (pptx-engine/model3d.ts).
 */
import { NO_ROTATION, type Model3dRotation } from '@genoffice/pptx-engine/model3d'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { ActionCtx } from '../action-context'
import { FIT_WIDTH } from '../app-constants'
import { t } from '../i18n/locale'
import { bytesToBase64 } from '../insert-actions'
import { loadModel, renderingSize, type ModelView } from './scene'

/** A new model's frame is half the slide high, as wide as the model seen from the front. */
const INSERT_HEIGHT_RATIO = 0.5
/** A frame no wider or taller than this many times its other side. */
const MAX_FRAME_ASPECT = 3

export async function insertModel3dFile(ctx: ActionCtx): Promise<void> {
  const slide = ctx.slide
  if (!slide) return
  const picked = await window.slidesApi.pickModel3d()
  if (!picked) return
  let view: ModelView | null = null
  let posterBase64: string | undefined
  let aspect = 1
  try {
    view = await loadModel(picked.bytes)
    aspect = Math.min(MAX_FRAME_ASPECT, Math.max(1 / MAX_FRAME_ASPECT, view.frontAspect))
    const height = slide.heightPx * INSERT_HEIGHT_RATIO
    const size = renderingSize(height * aspect, height)
    posterBase64 = bytesToBase64(await view.png(size.width, size.height, NO_ROTATION))
  } catch {
    ctx.setStatus(t('appStatusModel3dNoPreview'))
  } finally {
    view?.dispose()
  }
  const r = await window.slidesApi.addModel3d({
    slideIndex: ctx.current,
    bytes: picked.bytes,
    name: picked.name,
    ...(posterBase64 ? { posterBase64 } : {}),
    aspect,
    fitWidthPx: FIT_WIDTH,
  })
  if (!r) return
  ctx.applySlide(ctx.current, r.slide)
  ctx.setSelectedIds([r.sourceId])
  if (posterBase64) ctx.setStatus(t('appStatusModel3dInserted'))
}

/** Store `rot` and the model drawn at it; `view` is reused when the caller already has the model loaded. */
export async function storeModel3dView(
  ctx: ActionCtx,
  sourceId: string,
  box: { w: number; h: number },
  rot: Model3dRotation,
  view?: ModelView,
): Promise<RenderSlide | null> {
  const slideIndex = ctx.current
  let owned: ModelView | null = null
  try {
    if (!view) {
      const bytes = await window.slidesApi.getModel3dData(slideIndex, sourceId)
      if (!bytes) return null
      view = owned = await loadModel(bytes)
    }
    const size = renderingSize(box.w, box.h)
    const posterBase64 = bytesToBase64(await view.png(size.width, size.height, rot))
    const updated = await window.slidesApi.setModel3dView({
      slideIndex,
      sourceId,
      rot,
      posterBase64,
    })
    if (updated) ctx.applySlide(slideIndex, updated)
    return updated
  } catch {
    ctx.setStatus(t('appStatusModel3dNoPreview'))
    return null
  } finally {
    owned?.dispose()
  }
}
