/**
 * Insert > Zoom: Slide Zoom / Section Zoom / Summary Zoom. Every tile is a snapshot picture
 * of its target slide (what PowerPoint writes as the fallback image for older viewers) with a
 * jump link to it; the whole insert is one undo step.
 */
import type { RenderSlide } from '@genoffice/pptx-render'
import type { GetLayoutsResult, SectionInfo } from '../shared/ipc'
import type { ActionCtx } from './action-context'
import { FIT_WIDTH } from './app-constants'
import { t } from './i18n/locale'
import {
  graphicFrameInsertFrame,
  summaryZoomLayout,
  zoomCascadeFrames,
  type InsertFrame,
} from './insert-defaults'
import { groupSections, indexRange } from './section-groups'

export type ZoomMode = 'slide' | 'section' | 'summary'

interface ZoomTile {
  target: number
  frame: InsertFrame
  name: string
}

const sortUnique = (indexes: number[]): number[] => [...new Set(indexes)].sort((a, b) => a - b)

async function batched<T>(fn: () => Promise<T>): Promise<T> {
  const opened = await window.slidesApi.beginHistoryBatch()
  try {
    return await fn()
  } finally {
    if (opened) await window.slidesApi.endHistoryBatch()
  }
}

/** Drops the tiles on slideIndex in order (later tiles on top); returns the last page state and the new ids. */
async function placeTiles(
  slides: RenderSlide[],
  images: ActionCtx['images'],
  slideIndex: number,
  tiles: ZoomTile[],
): Promise<{ slide: RenderSlide | null; ids: string[] }> {
  // Lazy: export-render pulls in Konva, which the rest of this module never needs
  const { renderSlidesToPngBase64 } = await import('./export-render')
  const pngs = await renderSlidesToPngBase64(
    tiles.map((tile) => slides[tile.target]!),
    images,
    1,
  )
  let slide: RenderSlide | null = null
  const ids: string[] = []
  for (const [i, tile] of tiles.entries()) {
    const png = pngs[i]
    if (!png) continue
    const r = await window.slidesApi.addImageBytes({
      slideIndex,
      base64: png,
      ext: 'png',
      fitWidthPx: FIT_WIDTH,
      xPx: tile.frame.x,
      yPx: tile.frame.y,
      wPx: tile.frame.w,
      hPx: tile.frame.h,
      name: tile.name,
    })
    if (!r || 'error' in r) continue
    const linked = await window.slidesApi.setLink({
      slideIndex,
      sourceId: r.sourceId,
      target: { kind: 'slide', slideIndex: tile.target },
    })
    slide = linked ?? r.slide
    ids.push(r.sourceId)
  }
  return { slide, ids }
}

async function insertCascade(ctx: ActionCtx, targets: number[], label: string): Promise<void> {
  const { slide, current } = ctx
  if (!slide || !targets.length) return
  const frames = zoomCascadeFrames(slide, targets.length)
  const tiles = targets.map((target, i) => ({
    target,
    frame: frames[i]!,
    name: `${label} ${target + 1}`,
  }))
  const { slide: updated, ids } = await batched(() =>
    placeTiles(ctx.slides, ctx.images, current, tiles),
  )
  if (!updated) return
  ctx.applySlide(current, updated)
  ctx.setSelectedIds(ids)
  ctx.setStatus(
    ids.length === 1
      ? t('appStatusZoomInserted', { page: targets[0]! + 1 })
      : t('appStatusZoomsInserted', { count: ids.length }),
  )
}

export async function insertSlideZooms(ctx: ActionCtx, targets: number[]): Promise<void> {
  const valid = sortUnique(targets).filter((i) => ctx.slides[i])
  await insertCascade(ctx, valid, 'Slide Zoom')
}

/** sectionIndexes address groupSections() entries (the unsectioned lead group counts as PowerPoint's Default Section). */
export async function insertSectionZooms(ctx: ActionCtx, sectionIndexes: number[]): Promise<void> {
  const groups = groupSections(ctx.sections, ctx.slides.length) ?? []
  const targets = sortUnique(sectionIndexes)
    .map((i) => groups[i]?.start)
    .filter((start): start is number => start !== undefined)
  await insertCascade(ctx, targets, 'Section Zoom')
}

/** Title placeholder text of a slide ('' when it has none). */
export function slideTitleText(slide: RenderSlide): string {
  for (const node of slide.nodes) {
    if (node.type !== 'shape' && node.type !== 'text') continue
    if (node.placeholder !== 'title' && node.placeholder !== 'ctrTitle') continue
    return (node.text?.lines ?? [])
      .map((line) =>
        line.runs
          .filter((run) => !run.isBullet)
          .map((run) => run.text)
          .join(''),
      )
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim()
  }
  return ''
}

const isTitlePh = (type: string) => type === 'title' || type === 'ctrTitle'
const isBodyPh = (type: string) => type === 'body' || type === 'obj' || type === ''

/** "Title and Content": layoutType obj, else a title plus exactly one body, else any layout with a body. */
export function summaryLayoutPath(layouts: GetLayoutsResult['layouts'] | null): string | null {
  if (!layouts?.length) return null
  const byType = layouts.find((l) => l.layoutType === 'obj')
  if (byType) return byType.path
  const byShape = layouts.find(
    (l) =>
      l.placeholders.length === 2 &&
      l.placeholders.some((ph) => isTitlePh(ph.type)) &&
      l.placeholders.some((ph) => isBodyPh(ph.type)),
  )
  if (byShape) return byShape.path
  return layouts.find((l) => l.placeholders.some((ph) => isBodyPh(ph.type)))?.path ?? null
}

/**
 * Section list after a summary slide lands at `first` (every original index >= first moves
 * one down): existing sections keep id, name and slides; the new slide gets its own Summary
 * Section; each chosen slide not already heading a section starts one. Computed from the
 * pre-insert state because the engine's moveSlide folds a slide dropped at 0 into the section
 * of the slide it displaces, which would turn the old head into a non-head.
 */
export function summaryZoomSections(
  original: SectionInfo[],
  total: number,
  first: number,
  chosen: number[],
  nameOf: (pos: number) => string,
): SectionInfo[] {
  const heads = new Map<number, { id: string; name: string }>()
  for (const g of groupSections(original, total) ?? []) {
    if (g.id != null) heads.set(g.start + (g.start >= first ? 1 : 0), { id: g.id, name: g.name })
  }
  for (const c of chosen) {
    const pos = c + 1
    if (!heads.has(pos)) heads.set(pos, { id: '', name: nameOf(pos) })
  }
  heads.set(first, { id: '', name: t('appSectionSummary') })
  const starts = [...heads.keys()].sort((a, b) => a - b)
  return starts.map((start, i) => ({
    ...heads.get(start)!,
    slideIndices: indexRange(start, i + 1 < starts.length ? starts[i + 1]! : total + 1),
  }))
}

async function insertSummarySlide(
  ctx: ActionCtx,
  before: number,
): Promise<{ slides: RenderSlide[]; index: number } | null> {
  const layoutPath = summaryLayoutPath(
    ctx.layouts ?? (await window.slidesApi.getLayouts())?.layouts ?? null,
  )
  if (!layoutPath) {
    return window.slidesApi.addBlankSlide({
      sourceIndex: Math.max(before - 1, 0),
      fitWidthPx: FIT_WIDTH,
      ...(before === 0 ? { before: true } : {}),
    })
  }
  const r = await window.slidesApi.addSlideWithLayout({
    sourceIndex: Math.max(before - 1, 0),
    layoutPath,
    fitWidthPx: FIT_WIDTH,
  })
  if (!r || before > 0) return r
  // The layout op only appends after a slide; landing at position 0 takes one more move
  const moved = await window.slidesApi.moveSlide({ fromIndex: r.index, toIndex: 0 })
  return moved ? { slides: moved.slides, index: 0 } : r
}

/**
 * Summary Zoom: a new "Title and Content" slide before the first chosen slide whose body
 * placeholder is replaced by a grid of tiles; the new slide gets its own "Summary Section"
 * and every chosen slide that does not already head a section starts one.
 */
export async function insertSummaryZoom(ctx: ActionCtx, slideIndices: number[]): Promise<void> {
  const chosen = sortUnique(slideIndices).filter((i) => ctx.slides[i])
  if (!chosen.length) return
  const first = chosen[0]!
  const done = await batched(async () => {
    const added = await insertSummarySlide(ctx, first)
    if (!added) return null
    let { slides } = added
    const index = added.index
    // Every chosen slide sits one position later now
    const targets = chosen.map((i) => i + 1)
    const page = slides[index]!
    let box: InsertFrame = graphicFrameInsertFrame(page)
    const body = page.nodes.find(
      (n) => (n.type === 'shape' || n.type === 'text') && n.placeholder && isBodyPh(n.placeholder),
    )
    if (body) {
      box = { x: body.box.x, y: body.box.y, w: body.box.w, h: body.box.h }
      const cleared = await window.slidesApi.deleteElement({
        slideIndex: index,
        sourceId: body.sourceId,
      })
      if (cleared) slides = slides.map((s, i) => (i === index ? cleared : s))
    }
    const frames = summaryZoomLayout(targets.length, box, page.widthPx / page.heightPx)
    const tiles = targets.map((target, i) => ({
      target,
      frame: frames[i]!,
      name: 'Summary Zoom',
    }))
    const placed = await placeTiles(slides, ctx.images, index, tiles)
    if (placed.slide) slides = slides.map((s, i) => (i === index ? placed.slide! : s))

    const wanted = summaryZoomSections(
      ctx.sections,
      ctx.slides.length,
      first,
      chosen,
      (pos) => slideTitleText(slides[pos]!) || t('appSectionN', { n: pos + 1 }),
    )
    const sections = (await window.slidesApi.setSections(wanted)) ?? wanted
    return { slides, index, sections, count: placed.ids.length }
  })
  if (!done) return
  ctx.setSlides(done.slides)
  ctx.setSections(done.sections)
  ctx.setCurrent(done.index)
  ctx.setSelectedSlides([done.index])
  ctx.setSelectedIds([])
  ctx.setEditing(null)
  ctx.setDirty(true)
  ctx.setStatus(t('appStatusSummaryZoomInserted', { count: done.count }))
}
