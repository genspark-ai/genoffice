/**
 * File actions extracted from App.tsx: save / save-as, image & PDF
 * export. Each function takes the ActionCtx built fresh per call.
 * (Printing lives in components/PrintDialog.tsx — preview + options dialog.)
 */
import type { RenderSlide } from '@genoffice/pptx-render'
import type { ExportPdfLink } from '../shared/ipc'
import type { ActionCtx } from './action-context'
import { collectExportPdfLinks } from './export-links'
import { renderSlidesToPdfPages } from './export-pages'
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
  const page = ctx.slides[ctx.current]
  const remap = (id: string) => {
    const i = page?.nodes.findIndex((n) => n.sourceId === id) ?? -1
    return i >= 0 ? (next[ctx.current]?.nodes[i]?.sourceId ?? null) : null
  }
  ctx.setSelectedIds((ids) => ids.map(remap).filter((x): x is string => x !== null))
  ctx.setEnteredGroupId(null) // Group children ids can't be mapped by top-level ordinal; exit in-group editing after save
  // A stale edit target (page gone, deck shrunk) cannot be mapped: drop it rather than keep an id no node owns
  ctx.setEditing((e) => {
    const id = e && remap(e.sourceId)
    return id ? { sourceId: id } : null
  })
  ctx.setEditingCell((c) => {
    const id = c && remap(c.sourceId)
    return id && c ? { ...c, sourceId: id } : null
  })
  ctx.setSlides(next)
}

/**
 * Serializes save passes: a call that arrives while a save is in flight waits
 * for it instead of running concurrently. Two overlapping saves write the
 * same file with two `createWriteStream` pipes — interleaved zip streams,
 * truncated pptx, or EPERM/EBUSY on Windows. The queue is a simple promise
 * chain: each caller awaits the previous tail, then runs its own pass.
 */
let saveTail: Promise<unknown> | null = null

/**
 * Runs `pass` after the in-flight save (if any) finishes, and becomes the
 * tail subsequent saves wait on. A pass that throws still releases the
 * queue; the error propagates to its own caller only.
 */
async function runSerialized<T>(pass: () => Promise<T>): Promise<T> {
  const prior = saveTail
  const current = (async (): Promise<T> => {
    if (prior) await prior.catch(() => undefined)
    return pass()
  })()
  const tail = current.then(
    () => undefined,
    () => undefined,
  )
  saveTail = tail
  void current.then(
    () => {
      if (saveTail === tail) saveTail = null
    },
    () => {
      if (saveTail === tail) saveTail = null
    },
  )
  return current
}

export async function save(getCtx: () => ActionCtx, quiet = false): Promise<boolean> {
  return runSerialized(async () => {
    // resolved only now: a queued pass must remap selection against the tree the prior save adopted
    const ctx = getCtx()
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
      // quiet suppresses the success toast only — a failed save (incl. the 30s
      // auto-save) must surface, or edits silently stop reaching disk
      showToast(failed, 'error')
    }
    return r.ok
  })
}

export async function saveAs(getCtx: () => ActionCtx): Promise<void> {
  // Same queue as save(): Save + Save As (or double Save As) write through
  // the same main-process pipe and would interleave without it.
  await runSerialized(async () => {
    const ctx = getCtx()
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
  })
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

/**
 * Element and text-run hyperlinks of every exported page as clickable overlay
 * rects (the pages are rasterized, so links must ride along separately). Link
 * fetch failures degrade to a link-less PDF rather than failing the export.
 */
async function collectPdfLinks(ctx: ActionCtx): Promise<ExportPdfLink[][]> {
  const modelIndexes = ctx.slides.flatMap((s, i) => (s.hidden ? [] : [i]))
  const pageOfModelIndex = new Map(modelIndexes.map((mi, page) => [mi, page] as const))
  try {
    const [linkLists, runLinkLists] = await Promise.all([
      Promise.all(modelIndexes.map((mi) => window.slidesApi.getSlideLinks(mi))),
      Promise.all(modelIndexes.map((mi) => window.slidesApi.getRunLinks(mi))),
    ])
    const visible = modelIndexes.map((mi) => ctx.slides[mi]!)
    return collectExportPdfLinks(visible, linkLists, runLinkLists, pageOfModelIndex)
  } catch {
    return []
  }
}

/**
 * Export as PDF: each page (skipping hidden ones) rendered to vector SVG (text
 * stays text; a page the SVG backend cannot build falls back to its 2x PNG);
 * main process printToPDF in a hidden window.
 *
 * `outPath` skips the save dialog — the headless CLI entry already knows where
 * the file goes. Resolves true only when a PDF was written.
 */
export async function exportPdf(ctx: ActionCtx, outPath?: string): Promise<boolean> {
  const visible = ctx.slides.filter((s) => !s.hidden)
  if (visible.length === 0) {
    ctx.setStatus(t('appExportNoSlides'))
    return false
  }
  const target = outPath ?? (await window.slidesApi.pickExportPdfPath(`${exportBaseName(ctx)}.pdf`))
  if (!target) return false
  ctx.setStatus(t('appExportPdfProgress'))
  try {
    const { pages, fontCss } = await renderSlidesToPdfPages(visible, ctx.images)
    const links = await collectPdfLinks(ctx)
    const r = await window.slidesApi.exportPdf({
      filePath: target,
      pages,
      widthPx: visible[0].widthPx,
      heightPx: visible[0].heightPx,
      links,
      fontCss,
    })
    ctx.setStatus(
      r.ok
        ? t('appExportPdfDone', { path: r.path ?? '' })
        : t('appExportPdfFailed', { error: r.error ?? t('appUnknownError') }),
    )
    return r.ok
  } catch (err) {
    ctx.setStatus(t('appExportPdfFailed', { error: String(err) }))
    return false
  }
}
