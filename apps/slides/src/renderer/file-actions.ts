/**
 * File actions extracted from App.tsx: save / save-as, image & PDF
 * export. Each function takes the ActionCtx built fresh per call.
 * (Printing lives in components/PrintDialog.tsx — preview + options dialog.)
 */
import type { RenderSlide } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import { renderSlidesToPngBase64 } from './export-render'
import { t } from './i18n/locale'
import { showToast } from './components/toast-bus'

/**
 * If a text box/table is still being edited on ⌘S/close-save, blur first so the
 * overlay commits (blur→commitEdit), and save only after the commit lands —
 * otherwise we save pre-edit content and get a dirty-close prompt again.
 */
export async function flushActiveEdit(ctx: ActionCtx): Promise<void> {
  const active = document.activeElement as HTMLElement | null
  if (!active?.isContentEditable) return
  active.blur()
  for (let i = 0; i < 40 && ctx.editingActiveRef.current; i++) {
    await new Promise((r) => setTimeout(r, 50))
  }
}

/**
 * After save, the main process reopens the file with all-new element ids: swap
 * the render tree, mapping selection/edit state to new ids by per-page node ordinal.
 */
export function adoptSavedSlides(ctx: ActionCtx, next: RenderSlide[]): void {
  const remap = (id: string) => {
    const i = ctx.slides[ctx.current]?.nodes.findIndex((n) => n.sourceId === id) ?? -1
    return next[ctx.current]?.nodes[i]?.sourceId ?? null
  }
  ctx.setSelectedIds((ids) => ids.map(remap).filter((x): x is string => x !== null))
  ctx.setEnteredGroupId(null) // Group children ids can't be mapped by top-level ordinal; exit in-group editing after save
  ctx.setEditing((e) => (e && remap(e.sourceId) ? { sourceId: remap(e.sourceId)! } : e))
  ctx.setEditingCell((c) => (c && remap(c.sourceId) ? { ...c, sourceId: remap(c.sourceId)! } : c))
  ctx.setSlides(next)
}

export async function save(ctx: ActionCtx, quiet = false): Promise<boolean> {
  await flushActiveEdit(ctx)
  await ctx.flushNotes()
  const r = await window.slidesApi.save()
  if (r.ok) {
    if (r.slides) adoptSavedSlides(ctx, r.slides)
    if (r.path) ctx.setPath(r.path)
    ctx.setDirty(false)
    const saved = t('appStatusSaved')
    ctx.setStatus(saved)
    if (!quiet) showToast(saved)
  } else {
    const failed = t('appStatusSaveFailed', { error: r.error ?? t('appErrorCanceled') })
    ctx.setStatus(failed)
    if (!quiet) showToast(failed, 'error')
  }
  return r.ok
}

export async function saveAs(ctx: ActionCtx): Promise<void> {
  await flushActiveEdit(ctx)
  await ctx.flushNotes()
  const name = ctx.path?.split('/').pop() ?? 'presentation.pptx'
  const r = await window.slidesApi.saveAs(name)
  if (r.ok) {
    if (r.slides) adoptSavedSlides(ctx, r.slides)
    ctx.setPath(r.path ?? ctx.path)
    ctx.setDirty(false)
    const saved = t('appStatusSavedAs')
    ctx.setStatus(saved)
    showToast(saved)
  } else if (r.error) {
    // a canceled dialog returns ok:false without error — only real write
    // failures surface, matching the docs/sheets save-as feedback
    const failed = t('appStatusSaveFailed', { error: r.error })
    ctx.setStatus(failed)
    showToast(failed, 'error')
  }
}

/** Export base name: file name without the .pptx extension */
export function exportBaseName(ctx: ActionCtx): string {
  return (ctx.path?.split('/').pop() ?? t('appUntitledPresentation')).replace(/\.pptx$/i, '')
}

/** Export as images: each page (skipping hidden ones) rendered offscreen to 2x PNG, written to disk by the main process */
export async function exportImages(ctx: ActionCtx): Promise<void> {
  const visible = ctx.slides.filter((s) => !s.hidden)
  if (visible.length === 0) {
    ctx.setStatus(t('appExportNoSlides'))
    return
  }
  const dir = await window.slidesApi.pickExportDir()
  if (!dir) return
  ctx.setStatus(t('appExportImagesProgress', { count: visible.length }))
  try {
    const pngs = await renderSlidesToPngBase64(visible, ctx.images)
    const r = await window.slidesApi.exportImages({
      dir,
      baseName: exportBaseName(ctx),
      pngsBase64: pngs,
    })
    ctx.setStatus(
      r.ok
        ? t('appExportImagesDone', { count: r.paths?.length ?? 0, dir })
        : t('appExportImagesFailed', { error: r.error ?? t('appUnknownError') }),
    )
  } catch (err) {
    ctx.setStatus(t('appExportImagesFailed', { error: String(err) }))
  }
}

/** Export as PDF: each page (skipping hidden ones) rendered offscreen to 2x PNG; main process printToPDF in a hidden window */
export async function exportPdf(ctx: ActionCtx): Promise<void> {
  const visible = ctx.slides.filter((s) => !s.hidden)
  if (visible.length === 0) {
    ctx.setStatus(t('appExportNoSlides'))
    return
  }
  const target = await window.slidesApi.pickExportPdfPath(`${exportBaseName(ctx)}.pdf`)
  if (!target) return
  ctx.setStatus(t('appExportPdfProgress'))
  try {
    const pngs = await renderSlidesToPngBase64(visible, ctx.images)
    const r = await window.slidesApi.exportPdf({
      filePath: target,
      pngsBase64: pngs,
      widthPx: visible[0].widthPx,
      heightPx: visible[0].heightPx,
    })
    ctx.setStatus(
      r.ok
        ? t('appExportPdfDone', { path: r.path ?? '' })
        : t('appExportPdfFailed', { error: r.error ?? t('appUnknownError') }),
    )
  } catch (err) {
    ctx.setStatus(t('appExportPdfFailed', { error: String(err) }))
  }
}
