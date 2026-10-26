import type { Editor } from '@tiptap/core'
import type { AgentImage } from '@genoffice/agent-core'
import { endpointSupportsVision } from '@genoffice/ai-provider/browser'
import type { AiSettings } from '../../shared/ipc'
import { inlineImagesOf, type DocImagePlan } from './protocol'

/**
 * Document-picture vision injection: collect the pictures the block list
 * advertises and hand them to loop.run as ordinary multimodal input, so a
 * vision-capable chat model sees them natively instead of routing through the
 * analyze_media side provider. Text-only endpoints skip collection entirely
 * (plan null, previews keep the analyze_media wording).
 */

/** vision-safe wire formats (bmp/svg/tiff never reach the model; the editor transcodes tiff for display already) */
const VISION_MIMES: Record<string, true> = {
  'image/png': true,
  'image/jpeg': true,
  'image/gif': true,
  'image/webp': true,
}

const MAX_DOC_IMAGES = 8
/** one picture's base64 payload cap (≈5MB, matching the attachment rule) */
const MAX_IMAGE_BASE64_CHARS = 5_000_000
/** whole-message picture budget */
const MAX_TOTAL_BASE64_CHARS = 20_000_000

export interface DocImageCollection {
  images: AgentImage[]
  /** block index → doc-local 1-based ordinals (before the attachment offset) */
  ordinalsByBlock: Map<number, number[]>
  skipped: number
}

const DATA_URL_RE = /^data:(image\/[a-z0-9.+-]+);base64,(.+)$/is

/** data URLs parse in place; anything else (lazy media URLs) goes through the preload bridge */
async function resolveSrc(src: string): Promise<{ base64: string; mime: string } | null> {
  const m = DATA_URL_RE.exec(src)
  if (m) return { mime: m[1].toLowerCase(), base64: m[2] }
  const bridge = window.desktop?.readImageBytes
  if (typeof bridge !== 'function') return null
  const r = await bridge(src).catch(() => null)
  if (r?.base64 && r.mime) return { base64: r.base64, mime: r.mime.toLowerCase() }
  return null
}

/**
 * Walk the top-level blocks in buildDocumentContext order and collect their
 * pictures. Dedupes repeated references (headers, repeated shapes) by
 * resolved content. Budget overflow lands in `skipped`.
 */
export async function collectDocImages(
  editor: Editor,
  settings: AiSettings | null | undefined,
): Promise<DocImageCollection | null> {
  if (!settings || !endpointSupportsVision(settings)) return null
  const images: AgentImage[] = []
  const ordinalsByBlock = new Map<number, number[]>()
  const seen = new Map<string, number>()
  let skipped = 0
  let total = 0
  const doc = editor.state.doc
  for (let block = 0; block < doc.childCount; block++) {
    const node = doc.child(block)
    const srcs: string[] = []
    if (node.type.name === 'docProtected') {
      if (node.attrs.blockType === 'image' && typeof node.attrs.imageDataUrl === 'string') {
        srcs.push(node.attrs.imageDataUrl)
      }
    } else {
      for (const pic of inlineImagesOf(node)) {
        const src = String(pic.attrs?.dataUrl ?? '')
        if (src) srcs.push(src)
      }
    }
    if (srcs.length === 0) continue
    const ords: number[] = []
    for (const src of srcs) {
      const resolved = await resolveSrc(src)
      if (!resolved || !VISION_MIMES[resolved.mime]) {
        skipped++
        continue
      }
      const dupOrdinal = seen.get(resolved.base64)
      if (dupOrdinal !== undefined) {
        ords.push(dupOrdinal)
        continue
      }
      if (
        images.length >= MAX_DOC_IMAGES ||
        resolved.base64.length > MAX_IMAGE_BASE64_CHARS ||
        total + resolved.base64.length > MAX_TOTAL_BASE64_CHARS
      ) {
        skipped++
        continue
      }
      images.push({ base64: resolved.base64, mime: resolved.mime })
      total += resolved.base64.length
      const ordinal = images.length
      seen.set(resolved.base64, ordinal)
      ords.push(ordinal)
    }
    if (ords.length > 0) ordinalsByBlock.set(block, ords)
  }
  if (images.length === 0) return { images, ordinalsByBlock, skipped }
  return { images, ordinalsByBlock, skipped }
}

/** shift doc-local ordinals past the attachment images that precede them in the same message */
export function withImageOffset(
  collection: DocImageCollection,
  offset: number,
): DocImagePlan {
  if (offset === 0) return { ordinalsByBlock: collection.ordinalsByBlock, skipped: collection.skipped }
  const shifted = new Map<number, number[]>()
  for (const [block, ords] of collection.ordinalsByBlock) {
    shifted.set(
      block,
      ords.map((o) => o + offset),
    )
  }
  return { ordinalsByBlock: shifted, skipped: collection.skipped }
}
