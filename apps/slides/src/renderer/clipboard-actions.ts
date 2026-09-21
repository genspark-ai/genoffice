/**
 * Clipboard and format-painter actions extracted from App.tsx:
 * element copy/cut/paste/duplicate, slide copy/paste, external clipboard
 * ingestion, and the format painter. Functions take the ActionCtx built fresh per call.
 */
import type { ActionCtx } from './action-context'
import type { PasteSlideMode, PasteSlideResult } from '../shared/ipc'
import { rangeSelection } from '../shared/slide-selection'
import { FIT_WIDTH } from './app-constants'
import { renderSlidesToPngBase64 } from './export-render'
import { renderSelectionToPngBase64 } from './selection-image'
import {
  extractBrushFormat,
  computeBrushApply,
  buildEditParagraphsWithFormat,
} from './format-brush'
import { t } from './i18n/locale'

export async function deleteSelected(ctx: ActionCtx): Promise<void> {
  if (ctx.selectedIds.length) {
    const updated = await window.slidesApi.deleteElements({
      slideIndex: ctx.current,
      sourceIds: ctx.selectedIds,
    })
    if (updated) ctx.applySlide(ctx.current, updated)
  }
  ctx.setSelectedIds([])
}

export async function copySelected(ctx: ActionCtx): Promise<void> {
  if (ctx.selectedIds.length === 0) return
  const clipboardToken = crypto.randomUUID()
  const sourceIds = [...ctx.selectedIds]
  const n = await window.slidesApi.copyElements({
    slideIndex: ctx.current,
    sourceIds,
    clipboardToken,
  })
  if (n > 0) {
    ctx.setHasClipboard(true)
    ctx.setStatus(t('appStatusCopied', { count: n }))
    if (ctx.slide) {
      try {
        const png = await renderSelectionToPngBase64(ctx.slide, sourceIds, ctx.images)
        await window.slidesApi.copyElementsImage(clipboardToken, png)
      } catch {
        // The editable copy is already available if rendering or image decoding fails.
      }
    }
  }
}

export async function cutSelected(ctx: ActionCtx): Promise<void> {
  if (ctx.selectedIds.length === 0) return
  const n = await window.slidesApi.copyElements({
    slideIndex: ctx.current,
    sourceIds: ctx.selectedIds,
    cut: true,
  })
  if (n > 0) {
    ctx.setHasClipboard(true)
    await deleteSelected(ctx)
    ctx.setStatus(t('appStatusCut', { count: n }))
  }
}

/** Insert an external image at PowerPoint's dpi-aware size (main's pictureFrame), centered on the page or `atPx`. */
export async function insertExternalImage(
  ctx: ActionCtx,
  base64: string,
  ext: string,
  atPx?: { x: number; y: number },
): Promise<void> {
  const naturalPx = await new Promise<{ width: number; height: number }>((resolve) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth || 320, height: img.naturalHeight || 240 })
    img.onerror = () => resolve({ width: 320, height: 240 })
    img.src = `data:image/${ext};base64,${base64}`
  })
  const r = await window.slidesApi.addImageBytes({
    slideIndex: ctx.current,
    base64,
    ext,
    naturalPx,
    ...(atPx ? { centerPx: atPx } : {}),
    fitWidthPx: FIT_WIDTH,
  })
  if (!r) return
  if ('error' in r) {
    ctx.setSelectedIds([])
    ctx.setStatus(t('appStatusImageUnsupported', { ext: r.ext.toUpperCase() }))
    return
  }
  ctx.applySlide(ctx.current, r.slide)
  ctx.setSelectedIds([r.sourceId])
}

export async function copySlides(ctx: ActionCtx, indexes: number[]): Promise<void> {
  const sel = [...new Set(indexes)].sort((a, b) => a - b).filter((i) => ctx.slides[i])
  if (!sel.length) return
  // Renderings ride along for 'picture'-mode pastes (the source deck may be
  // gone by paste time). Copy still succeeds if the capture fails.
  let pngs: string[] | undefined
  try {
    pngs = await renderSlidesToPngBase64(
      sel.map((i) => ctx.slides[i]!),
      ctx.images,
    )
  } catch {
    pngs = undefined
  }
  const ok = await window.slidesApi.copySlides({ slideIndexes: sel, ...(pngs ? { pngs } : {}) })
  ctx.setCanPasteSlide(ok)
  ctx.setStatus(ok ? t('appStatusSlideCopied') : t('appStatusSlideCopyFailed'))
}

function applySlidePaste(ctx: ActionCtx, r: PasteSlideResult, mode: PasteSlideMode): void {
  ctx.setSlides(r.slides)
  ctx.setCurrent(r.index)
  ctx.setSelectedSlides(rangeSelection(r.index, r.index + Math.max(r.count - 1, 0)))
  ctx.setSelectedIds(r.sourceIds ?? [])
  ctx.setEditing(null)
  ctx.setDirty(true)
  ctx.setPasteFloater({ index: r.index, mode })
}

export async function pasteSlideAfter(
  ctx: ActionCtx,
  index: number,
  mode: PasteSlideMode = 'theme',
): Promise<void> {
  const r = await window.slidesApi.pasteSlide({ afterIndex: index, fitWidthPx: FIT_WIDTH, mode })
  if (!r) {
    ctx.setStatus(t('appStatusSlidePasteFailed'))
    return
  }
  applySlidePaste(ctx, r, mode)
}

/** Paste-options floater: redo the just-completed paste with another mode. */
export async function repasteSlideAs(ctx: ActionCtx, mode: PasteSlideMode): Promise<void> {
  if (ctx.pasteFloater?.mode === mode) return
  const r = await window.slidesApi.repasteSlide({ mode, fitWidthPx: FIT_WIDTH })
  if (!r) {
    ctx.setPasteFloater(null)
    ctx.setStatus(t('appStatusPasteOptionsExpired'))
    return
  }
  applySlidePaste(ctx, r, mode)
}

/**
 * Paste: whole slide (slide marker newest) → app elements (element marker not
 * overwritten externally) → external images → external text into a text box
 */
export async function pasteClipboard(ctx: ActionCtx): Promise<void> {
  const external = await window.slidesApi.clipboardExternal()
  if (external.kind === 'slide') {
    await pasteSlideAfter(ctx, Math.max(...ctx.selectedSlides, ctx.current))
    return
  }
  if (external.kind === 'image') {
    await insertExternalImage(ctx, external.base64, external.ext)
    return
  }
  if (external.kind === 'text') {
    const w = 400
    const r = await window.slidesApi.addElement({
      slideIndex: ctx.current,
      kind: 'textbox',
      xPx: ((ctx.slide?.widthPx ?? FIT_WIDTH) - w) / 2,
      yPx: (ctx.slide?.heightPx ?? FIT_WIDTH * 0.5625) / 2 - 40,
      wPx: w,
      hPx: 80,
      fitWidthPx: FIT_WIDTH,
      paragraphs: external.text.split(/\r?\n/).map((line) => ({ runs: [{ text: line }] })),
    })
    if (r) {
      ctx.applySlide(ctx.current, r.slide)
      ctx.setSelectedIds([r.sourceId])
    }
    return
  }
  const r = await window.slidesApi.pasteElements({ slideIndex: ctx.current, fitWidthPx: FIT_WIDTH })
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds(r.sourceIds)
    ctx.setEditing(null)
  }
}

/** Duplicate in place (⌘D / Option+drag copy): doesn't touch the app clipboard, the copy becomes selected. */
export async function duplicateSelected(
  ctx: ActionCtx,
  ids: string[],
  dxPx = 16,
  dyPx = 16,
): Promise<void> {
  if (!ids.length) return
  const r = await window.slidesApi.duplicateElements({
    slideIndex: ctx.current,
    sourceIds: ids,
    dxPx,
    dyPx,
    fitWidthPx: FIT_WIDTH,
  })
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds(r.sourceIds)
  }
}

// ── Format painter operations ───────────────────────────────────────────────

/** ⌘⇧C: copy the selected single element's format to the clipboard (without entering paste mode) */
export function copyFormat(ctx: ActionCtx): void {
  if (ctx.selectedIds.length !== 1 || !ctx.slide) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === ctx.selectedIds[0])
  if (!node) return
  const fmt = extractBrushFormat(node)
  if (fmt) {
    ctx.setBrushFormat(fmt)
    ctx.setStatus(t('appStatusFormatCopied'))
  }
}

/** ⌘⇧V: paste format onto selected elements (format must have been copied first) */
export async function pasteFormat(ctx: ActionCtx): Promise<void> {
  if (!ctx.brushFormat || ctx.selectedIds.length === 0 || !ctx.slide) return
  for (const id of ctx.selectedIds) {
    const node = ctx.slide.nodes.find((n) => n.sourceId === id)
    if (!node) continue
    const plan = computeBrushApply(ctx.brushFormat, node)
    // Apply fill
    if (plan.fillColor !== undefined) {
      await window.slidesApi.editFill({
        slideIndex: ctx.current,
        sourceId: id,
        fill: plan.fillColor,
      })
    }
    // Apply stroke
    if (plan.stroke !== undefined) {
      await window.slidesApi.editStroke({
        slideIndex: ctx.current,
        sourceId: id,
        stroke: plan.stroke,
      })
    }
    // Apply text format (runs + alignment)
    if ((plan.runFormat || plan.align) && (node.type === 'shape' || node.type === 'text')) {
      const paras = buildEditParagraphsWithFormat(node, plan.runFormat, plan.align)
      if (paras) {
        const updated = await window.slidesApi.editText({
          slideIndex: ctx.current,
          sourceId: id,
          paragraphs: paras,
        })
        if (updated) ctx.applySlide(ctx.current, updated)
      }
    } else if (plan.fillColor !== undefined || plan.stroke !== undefined) {
      // fill/stroke were applied; re-render to get the latest page
      const updated = await window.slidesApi.editFill({
        slideIndex: ctx.current,
        sourceId: id,
        fill: plan.fillColor ?? 'none',
      })
      if (updated) ctx.applySlide(ctx.current, updated)
    }
  }
  // Fetch the latest page again
  // (each operation above already called applySlide; pull once more at the end to be sure it's fresh)
  ctx.setStatus(t('appStatusFormatPasted'))
}

/** Format painter button single click: copy current format + enter single-paste mode */
export function onFormatBrushClick(ctx: ActionCtx): void {
  if (ctx.selectedIds.length !== 1 || !ctx.slide) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === ctx.selectedIds[0])
  if (!node) return
  const fmt = extractBrushFormat(node)
  if (!fmt) return
  ctx.setBrushFormat(fmt)
  ctx.setBrushMode('once')
  ctx.setStatus(t('appStatusBrushOnce'))
}

/** Format painter button double click: enter continuous mode */
export function onFormatBrushDoubleClick(ctx: ActionCtx): void {
  if (ctx.selectedIds.length !== 1 || !ctx.slide) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === ctx.selectedIds[0])
  if (!node) return
  const fmt = extractBrushFormat(node)
  if (!fmt) return
  ctx.setBrushFormat(fmt)
  ctx.setBrushMode('continuous')
  ctx.setStatus(t('appStatusBrushContinuous'))
}

/** Apply format when an element is clicked in format painter mode */
export async function applyBrushToElement(ctx: ActionCtx, targetId: string): Promise<void> {
  if (!ctx.brushFormat || !ctx.slide) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === targetId)
  if (!node) return
  const plan = computeBrushApply(ctx.brushFormat, node)

  if (plan.fillColor !== undefined) {
    const updated = await window.slidesApi.editFill({
      slideIndex: ctx.current,
      sourceId: targetId,
      fill: plan.fillColor,
    })
    if (updated) ctx.applySlide(ctx.current, updated)
  }
  if (plan.stroke !== undefined) {
    const updated = await window.slidesApi.editStroke({
      slideIndex: ctx.current,
      sourceId: targetId,
      stroke: plan.stroke,
    })
    if (updated) ctx.applySlide(ctx.current, updated)
  }
  if ((plan.runFormat || plan.align) && (node.type === 'shape' || node.type === 'text')) {
    const paras = buildEditParagraphsWithFormat(node, plan.runFormat, plan.align)
    if (paras) {
      const updated = await window.slidesApi.editText({
        slideIndex: ctx.current,
        sourceId: targetId,
        paragraphs: paras,
      })
      if (updated) ctx.applySlide(ctx.current, updated)
    }
  }

  if (ctx.brushMode === 'once') {
    ctx.setBrushMode(null)
    ctx.setStatus(t('appStatusFormatApplied'))
  }
}
