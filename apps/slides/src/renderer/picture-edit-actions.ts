/**
 * Picture crop and cutout (background removal) actions. Extracted from
 * App.tsx; functions read the latest App state through ActionCtx.
 */
import type { PictureRenderNode } from '@genoffice/pptx-render'
import { FIT_WIDTH } from './app-constants'
import type { ActionCtx } from './action-context'
import { t } from './i18n/locale'

/** Enter picture crop mode: find the selected picture node and read its box and srcRect */
export function startCrop(ctx: ActionCtx): void {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return
  const id = ctx.selectedIds[0]!
  const node = ctx.slide.nodes.find((n) => n.sourceId === id)
  if (!node || node.type !== 'picture') return
  const picNode = node as PictureRenderNode
  // Reconstruct the full original-image extent from the displayed (cropped) box:
  // the crop is non-destructive (srcRect), so the frame may expand back out to it.
  const sr = picNode.srcRect
  const kw = sr ? Math.max(1 - sr.l - sr.r, 0.01) : 1
  const kh = sr ? Math.max(1 - sr.t - sr.b, 0.01) : 1
  const fullW = picNode.box.w / kw
  const fullH = picNode.box.h / kh
  ctx.setCropTarget({
    sourceId: id,
    box: { x: picNode.box.x, y: picNode.box.y, w: picNode.box.w, h: picNode.box.h },
    srcRect: sr,
    fullBox: {
      x: picNode.box.x - (sr?.l ?? 0) * fullW,
      y: picNode.box.y - (sr?.t ?? 0) * fullH,
      w: fullW,
      h: fullH,
    },
    dataUrl: picNode.dataUrl,
  })
}

export async function commitCrop(
  ctx: ActionCtx,
  rect: { l: number; t: number; r: number; b: number } | null,
): Promise<void> {
  if (!ctx.cropTarget) return
  const { sourceId, srcRect: prev, fullBox } = ctx.cropTarget
  ctx.setCropTarget(null)
  // `rect` is relative to the FULL original image (the frame may have been dragged
  // outward past the previous crop); null = frame covers the whole original.
  if (!rect && !prev) return // never cropped and frame left at full — nothing to do
  // The element frame moves/resizes to the on-screen crop frame in the same atomic
  // edit: the kept region stays exactly where it was framed (PowerPoint semantics),
  // dragging outward restores original content, and one undo reverts everything.
  const updated = await window.slidesApi.editPictureSrcRect({
    slideIndex: ctx.current,
    sourceId,
    srcRect: rect,
    boxPx: rect
      ? {
          x: fullBox.x + rect.l * fullBox.w,
          y: fullBox.y + rect.t * fullBox.h,
          w: Math.max(1, fullBox.w * (1 - rect.l - rect.r)),
          h: Math.max(1, fullBox.h * (1 - rect.t - rect.b)),
        }
      : { x: fullBox.x, y: fullBox.y, w: fullBox.w, h: fullBox.h }, // full frame = original restored
    fitWidthPx: FIT_WIDTH,
  })
  if (updated) {
    ctx.applySlide(ctx.current, updated)
    ctx.setDirty(true)
    ctx.setStatus(rect ? t('appStatusCropApplied') : t('appStatusCropRemoved'))
  }
}

export function cancelCrop(ctx: ActionCtx): void {
  ctx.setCropTarget(null)
}

// ── Picture cutout (background removal) ───────────────────────────────────────

/** Enter cutout mode: a single selected picture (except audio/video poster frames) → open the tolerance preview dialog */
export function startCutout(ctx: ActionCtx): void {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === ctx.selectedIds[0])
  if (!node || node.type !== 'picture') return
  const pic = node as PictureRenderNode
  if (pic.media) {
    ctx.setStatus(t('appStatusCutoutMediaUnsupported'))
    return
  }
  if (!pic.dataUrl) {
    ctx.setStatus(t('appStatusCutoutNoData'))
    return
  }
  ctx.setCutoutTarget({ sourceId: pic.sourceId, dataUrl: pic.dataUrl })
}

/**
 * Apply the cutout result: swap the picture's backing image for the
 * background-removed PNG in place. One atomic IPC — frame, rotation, z-order,
 * border and effects all survive; the crop window is kept because the result
 * PNG shares the source image's pixel geometry.
 */
export async function applyCutout(ctx: ActionCtx, pngDataUrl: string): Promise<void> {
  if (!ctx.cutoutTarget || !ctx.slide) return
  const targetId = ctx.cutoutTarget.sourceId
  const node = ctx.slide.nodes.find((n) => n.sourceId === targetId)
  ctx.setCutoutTarget(null)
  if (!node || node.type !== 'picture') return
  const base64 = pngDataUrl.split(',')[1]
  if (!base64) {
    ctx.setStatus(t('appStatusCutoutEncodeFailed'))
    return
  }
  const updated = await window.slidesApi.replacePictureBytes({
    slideIndex: ctx.current,
    sourceId: targetId,
    base64,
    ext: 'png',
    keepSrcRect: true,
  })
  if (!updated || 'error' in updated) {
    ctx.setStatus(t('appStatusCutoutInsertFailed'))
    return
  }
  ctx.applySlide(ctx.current, updated)
  ctx.setSelectedIds([targetId])
  ctx.setDirty(true)
  ctx.setStatus(t('appStatusCutoutDone'))
}
