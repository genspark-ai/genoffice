import type { AgentImage } from '@genoffice/agent-core'
import type { PageImageBytes, PageImageRef } from '../../shared/ipc'

/** Vision attachment budget, shared semantics with the other apps' injectors */
export const MAX_IMAGES = 8
/** Max base64 length of one image */
export const MAX_IMAGE_BASE64 = 5 * 1024 * 1024
/** Max total base64 length across all attached images */
export const MAX_TOTAL_BASE64 = 20 * 1024 * 1024

/** mime types every chat provider's image block accepts */
const VISION_MIME_ALLOWED: Record<string, true> = {
  'image/png': true,
  'image/jpeg': true,
  'image/webp': true,
}

/** Result of a send-time vision collection: what to attach and how to annotate it */
export interface PageImagePlan {
  images: AgentImage[]
  /** page index (0-based) → 1-based attachment ordinals aligned with images (index+1) */
  ordinalsByPage: Map<number, number[]>
  /** Images the budget dropped (past the caps, over-size, undecodable, duplicates) */
  skipped: number
}

/** Shared empty plan: no path / failed collection / non-vision endpoint */
export const EMPTY_PAGE_IMAGE_PLAN: PageImagePlan = {
  images: [],
  ordinalsByPage: new Map(),
  skipped: 0,
}

/** IPC surface the collector needs; injected so tests can mock the bridge */
export interface PageImageBridge {
  listImages(): Promise<PageImageRef[]>
  /** Pixels for one listed ref; null when the object can't be matched or decoded */
  imageBytes(ref: PageImageRef): Promise<PageImageBytes | null>
  /** Original page index → skip images on deleted pages */
  isDeleted(pageIndex: number): boolean
}

/** Page number ordering of the doc's images, stable for the annotation loop */
export function collectPageImages(bridge: PageImageBridge): Promise<PageImagePlan> {
  return (async () => {
    const plan: PageImagePlan = {
      images: [],
      ordinalsByPage: new Map(),
      skipped: 0,
    }
    let refs: PageImageRef[]
    try {
      refs = await bridge.listImages()
    } catch {
      return plan // the model still gets its text-only context
    }
    // deleted pages first, so the budget cap counts real candidates only
    const live = refs.filter((ref) => !bridge.isDeleted(ref.pageIndex))
    const seen = new Set<string>() // dedup by content
    let total = 0
    for (let i = 0; i < live.length; i++) {
      const ref = live[i]!
      // count cap: every remaining listed image is skipped, loop stops
      if (plan.images.length >= MAX_IMAGES) {
        plan.skipped += live.length - i
        break
      }
      let bytes: PageImageBytes | null
      try {
        bytes = await bridge.imageBytes(ref)
      } catch {
        bytes = null
      }
      // mime allowlist → same skip as undecodable (a PDF can embed a TIFF)
      if (!bytes || !VISION_MIME_ALLOWED[bytes.mime]) {
        plan.skipped++
        continue
      }
      if (bytes.base64.length > MAX_IMAGE_BASE64) {
        plan.skipped++
        continue
      }
      if (total + bytes.base64.length > MAX_TOTAL_BASE64) {
        plan.skipped++
        continue
      }
      if (seen.has(bytes.base64)) {
        plan.skipped++
        continue
      }
      seen.add(bytes.base64)
      total += bytes.base64.length
      plan.images.push({ base64: bytes.base64, mime: bytes.mime })
      plan.ordinalsByPage.set(
        ref.pageIndex,
        [...(plan.ordinalsByPage.get(ref.pageIndex) ?? []), plan.images.length],
      )
    }
    return plan
  })()
}

/** Context annotation lines: "page N has M picture(s) — attached as image #K[, #K+1]";
 *  overflow lines carry "(N more pictures not attached)". Empty when nothing attached. */
export function pageImageAnnotation(plan: PageImagePlan): string {
  if (plan.images.length === 0) return ''
  const lines: string[] = []
  const pages = [...plan.ordinalsByPage.keys()].sort((a, b) => a - b)
  for (const page of pages) {
    const ordinals = plan.ordinalsByPage.get(page) ?? []
    if (ordinals.length === 0) continue
    const tags = ordinals.map((k) => `#${k}`).join(', ')
    lines.push(`page ${page + 1} has ${ordinals.length} picture(s) — attached as image ${tags}`)
  }
  if (plan.skipped > 0) {
    lines.push(`(${plan.skipped} more pictures not attached)`)
  }
  return lines.join('\n')
}
