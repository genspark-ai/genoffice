/**
 * Slide- and section-level actions extracted from App.tsx:
 * add/duplicate/delete/cut slides, sections, and slide reorder. Element-level
 * clipboard lives in clipboard-actions.ts. Functions read the latest App
 * state through ActionCtx.
 */
import { FIT_WIDTH } from './app-constants'
import type { ActionCtx } from './action-context'
import type { SectionInfo } from '../shared/ipc'
import { renderSlidesToPngBase64 } from './export-render'
import { movedBlockPositions, rangeSelection } from '../shared/slide-selection'
import { t } from './i18n/locale'
import { currentAfterRemoval, groupSections, indexRange } from './section-groups'

const sortUnique = (indexes: number[]): number[] => [...new Set(indexes)].sort((a, b) => a - b)

export async function addSlide(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide) return
  const r = await window.slidesApi.addBlankSlide({
    sourceIndex: ctx.current,
    fitWidthPx: FIT_WIDTH,
  })
  if (r) {
    ctx.setSlides(r.slides)
    ctx.setCurrent(r.index)
    ctx.setSelectedSlides([r.index])
    ctx.setSelectedIds([])
    ctx.setEditing(null)
    ctx.setDirty(true)
  }
}

/** Insert a blank slide at position pos (0..slides.length), PowerPoint's insertion-point semantics */
export async function addSlideAt(ctx: ActionCtx, pos: number): Promise<void> {
  if (!ctx.slides.length) return
  const before = pos <= 0
  const r = await window.slidesApi.addBlankSlide({
    sourceIndex: before ? 0 : Math.min(pos, ctx.slides.length) - 1,
    fitWidthPx: FIT_WIDTH,
    ...(before ? { before: true } : {}),
  })
  if (r) {
    ctx.setSlides(r.slides)
    ctx.setCurrent(r.index)
    ctx.setSelectedSlides([r.index])
    ctx.setSelectedIds([])
    ctx.setEditing(null)
    ctx.setDirty(true)
  }
}

/** Switch (or, without layoutPath, reset) the layout of slide `index`. */
export async function setSlideLayoutAt(
  ctx: ActionCtx,
  index: number,
  layoutPath?: string,
): Promise<void> {
  const r = await window.slidesApi.setSlideLayout({ slideIndex: index, layoutPath })
  if (r) ctx.applySlide(index, r)
}

export async function addSlideWithLayout(ctx: ActionCtx, layoutPath: string): Promise<void> {
  if (!ctx.slide) return
  const r = await window.slidesApi.addSlideWithLayout({
    sourceIndex: ctx.current,
    layoutPath,
    fitWidthPx: FIT_WIDTH,
  })
  if (r) {
    ctx.setSlides(r.slides)
    ctx.setCurrent(r.index)
    ctx.setSelectedSlides([r.index])
    ctx.setSelectedIds([])
    ctx.setEditing(null)
    ctx.setDirty(true)
  }
}

/** The copies land in order after the last selected slide and become the selection (PowerPoint) */
export async function duplicateSlides(ctx: ActionCtx, indexes: number[]): Promise<void> {
  const sel = sortUnique(indexes)
  if (!sel.length) return
  const r = await window.slidesApi.duplicateSlides({ slideIndexes: sel, fitWidthPx: FIT_WIDTH })
  if (r) {
    ctx.setSlides(r.slides)
    ctx.setCurrent(r.index)
    ctx.setSelectedSlides(rangeSelection(r.index, r.index + sel.length - 1))
    ctx.setSelectedIds([])
    ctx.setEditing(null)
    ctx.setDirty(true)
  }
}

/** Deletes the selection in one undo step, keeping at least one slide; the slide after the last deleted one becomes current */
export async function deleteSlides(ctx: ActionCtx, indexes: number[]): Promise<void> {
  let sel = sortUnique(indexes)
  if (sel.length >= ctx.slides.length) {
    ctx.setStatus(t('appStatusKeepOneSlide'))
    sel = sel.slice(1)
  }
  if (!sel.length) return
  const r = await window.slidesApi.deleteSlides({ slideIndexes: sel })
  if (!r) return
  const next = Math.min(sel[sel.length - 1]! + 1 - sel.length, r.length - 1)
  ctx.setSlides(r)
  ctx.setCurrent(next)
  ctx.setSelectedSlides([next])
  ctx.setSelectedIds([])
  ctx.setEditing(null)
  ctx.setDirty(true)
}

export async function cutSlides(ctx: ActionCtx, indexes: number[]): Promise<void> {
  const sel = sortUnique(indexes)
  if (sel.length >= ctx.slides.length) {
    ctx.setStatus(t('appStatusKeepOneSlide'))
    return
  }
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
  if (!ok) {
    ctx.setStatus(t('appStatusSlideCopyFailed'))
    return
  }
  ctx.setCanPasteSlide(true)
  await deleteSlides(ctx, sel)
  ctx.setStatus(t('appStatusSlideCut'))
}

// ── Section management ─────────────────────────────────────────────────

export async function addSectionAt(ctx: ActionCtx, index: number): Promise<void> {
  const r = await window.slidesApi.addSection({
    atSlideIndex: index,
    name: t('appSectionUntitled'),
  })
  if (r) {
    ctx.setSections(r)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusSectionAdded'))
  }
}

export async function renameSectionTo(ctx: ActionCtx, id: string, name: string): Promise<void> {
  const r = await window.slidesApi.renameSection({ id, name })
  if (r) {
    ctx.setSections(r)
    ctx.setDirty(true)
  }
}

/** Remove the header only; the lead group (id null) hands its slides to the first real section. */
export async function removeSectionAt(ctx: ActionCtx, id: string | null): Promise<void> {
  const r =
    id == null
      ? await window.slidesApi.setSections(absorbLead(ctx))
      : await window.slidesApi.removeSection({ id })
  if (r) {
    ctx.setSections(r)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusSectionRemoved'))
  }
}

function absorbLead(ctx: ActionCtx): SectionInfo[] {
  const groups = groupSections(ctx.sections, ctx.slides.length) ?? []
  return groups.flatMap((g, i) =>
    g.id == null
      ? []
      : [{ id: g.id, name: g.name, slideIndices: indexRange(i <= 1 ? 0 : g.start, g.end) }],
  )
}

export async function removeAllSections(ctx: ActionCtx): Promise<void> {
  const r = await window.slidesApi.setSections([])
  if (r) {
    ctx.setSections(r)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusSectionsRemoved'))
  }
}

export async function removeSectionWithSlides(ctx: ActionCtx, id: string | null): Promise<void> {
  const group = groupSections(ctx.sections, ctx.slides.length)?.find((g) => g.id === id)
  if (!group) return
  const r = await window.slidesApi.removeSectionSlides({ id })
  if (!r) {
    if (group.end - group.start >= ctx.slides.length) ctx.setStatus(t('appStatusKeepOneSlide'))
    return
  }
  ctx.setSlides(r.slides)
  ctx.setSections(r.sections)
  ctx.setCurrent((c) => currentAfterRemoval(c, group, r.slides.length))
  ctx.setSelectedIds([])
  ctx.setEditing(null)
  ctx.setDirty(true)
  ctx.setStatus(t('appStatusSectionSlidesRemoved'))
}

export async function moveSectionDir(
  ctx: ActionCtx,
  id: string,
  dir: 'up' | 'down',
): Promise<void> {
  const r = await window.slidesApi.moveSection({ id, dir })
  if (r) {
    ctx.setSlides(r.slides)
    ctx.setSections(r.sections)
    ctx.setSelectedIds([])
    ctx.setEditing(null)
    ctx.setDirty(true)
    ctx.setStatus(dir === 'up' ? t('appStatusSectionMovedUp') : t('appStatusSectionMovedDown'))
  }
}

/** Drag-reorder: the selected slides land as a block at gap insertAt and stay selected; the anchor follows */
export async function moveSlidesTo(
  ctx: ActionCtx,
  indexes: number[],
  insertAt: number,
): Promise<void> {
  const sel = sortUnique(indexes)
  const landed = movedBlockPositions(sel, insertAt)
  if (sel.every((i, k) => i === landed[k])) return
  const r = await window.slidesApi.moveSlides({ slideIndexes: sel, insertAt })
  if (r) {
    const anchor = landed[Math.max(0, sel.indexOf(ctx.current))]!
    ctx.setSlides(r.slides)
    ctx.setSections(r.sections)
    ctx.setCurrent(anchor)
    ctx.setSelectedSlides(landed)
    ctx.setSelectedIds([])
    ctx.setEditing(null)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusSlideMoved', { page: anchor + 1 }))
  }
}

/** Commit section rename (Enter/blur); empty names aren't committed. */
export function commitRenameSection(ctx: ActionCtx): void {
  if (ctx.renamingSec && ctx.renamingSec.value.trim()) {
    void renameSectionTo(ctx, ctx.renamingSec.id, ctx.renamingSec.value.trim())
  }
  ctx.setRenamingSec(null)
}
