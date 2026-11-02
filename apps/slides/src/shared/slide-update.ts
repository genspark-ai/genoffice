import type { RenderSlide } from '@genoffice/pptx-render'
import type { SlideUpdate } from './ipc'

let refreshDeck: (() => void) | null = null

/** The full-deck refetch a window falls back to when a partial update names slides it has not seen. */
export function setDeckRefresh(refresh: (() => void) | null): void {
  refreshDeck = refresh
}

/** Merge by stable part identity; null when `order` names a slide neither side has. */
export function tryMergeSlideUpdate(
  current: RenderSlide[],
  update: SlideUpdate,
): RenderSlide[] | null {
  const byPath = new Map(current.map((slide) => [slide.partPath, slide]))
  for (const slide of update.slides) byPath.set(slide.partPath, slide)
  if (update.order.some((path) => !byPath.has(path))) return null
  return update.order.map((path) => byPath.get(path)!)
}

/**
 * Merge for state updaters, including when a full broadcast arrived first.
 * In a shared session a reply can name a slide another window added before
 * that window's broadcast reached us; then the deck is kept and refetched
 * instead of thrown away mid-updater.
 */
export function mergeSlideUpdate(current: RenderSlide[], update: SlideUpdate): RenderSlide[] {
  const merged = tryMergeSlideUpdate(current, update)
  if (merged) return merged
  if (!refreshDeck) {
    const known = new Set([...current, ...update.slides].map((slide) => slide.partPath))
    const missing = update.order.find((path) => !known.has(path))
    throw new Error(`Missing slide in render update: ${missing}`)
  }
  queueMicrotask(refreshDeck)
  return current
}

/** A partial broadcast needs a full refresh when its base deck is unavailable. */
export function mergeBroadcastUpdate(
  current: RenderSlide[] | null,
  update: SlideUpdate,
  refresh: () => void,
): RenderSlide[] | null {
  const paths = new Set([
    ...(current ?? []).map((slide) => slide.partPath),
    ...update.slides.map((slide) => slide.partPath),
  ])
  if (!current?.length || update.order.some((path) => !paths.has(path))) {
    queueMicrotask(refresh)
    return current
  }
  return mergeSlideUpdate(current, update)
}
