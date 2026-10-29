import type { RenderSlide } from '@genoffice/pptx-render'
import type { SlideUpdate } from './ipc'

/** Merge by stable part identity, including when a full broadcast arrived first. */
export function mergeSlideUpdate(current: RenderSlide[], update: SlideUpdate): RenderSlide[] {
  const byPath = new Map(current.map((slide) => [slide.partPath, slide]))
  for (const slide of update.slides) byPath.set(slide.partPath, slide)
  return update.order.map((path) => {
    const slide = byPath.get(path)
    if (!slide) throw new Error(`Missing slide in render update: ${path}`)
    return slide
  })
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
