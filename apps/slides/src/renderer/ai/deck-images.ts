/**
 * Deck-picture collection for AI vision input.
 *
 * When the active endpoint can take images, the send path walks the live
 * RenderSlide tree, gathers picture nodes (their dataUrl when already decoded,
 * otherwise bytes read back through the main process's media resolver), and
 * hands the result to loop.run after the user's attachment images. The deck
 * outline then annotates every collected picture with "attached as image #K"
 * so the model can map pixels back to element ids.
 *
 * Budgets mirror the attachment path (≤8 pictures, ≤5MB base64 each, ≤20MB
 * total, content dedup, png/jpeg/gif/webp only) so one image-heavy deck
 * cannot blow the request.
 */
import type { AgentImage } from '@genoffice/agent-core'
import type { PictureRenderNode, RenderNode, RenderSlide } from '@genoffice/pptx-render'
import type { DeckImageResult } from '../../shared/ipc'

/** Max deck pictures attached to one message */
export const MAX_DECK_IMAGES = 8
/** Per-picture base64 budget in chars (the attachment path's 5MB image cap) */
export const MAX_DECK_IMAGE_B64 = 5 * 1024 * 1024
/** Combined base64 budget across one message's deck pictures */
export const MAX_DECK_IMAGES_B64 = 20 * 1024 * 1024
/** Mimes a vision request accepts; TIFF reaches this layer already transcoded to PNG */
const DECK_IMAGE_MIME: Record<string, true> = {
  'image/png': true,
  'image/jpeg': true,
  'image/gif': true,
  'image/webp': true,
}

/** Main-process bridge: one deck picture's bytes as base64 (slides:read-deck-image) */
export type DeckImageReader = (op: {
  slideIndex: number
  elementId: string
}) => Promise<DeckImageResult>

/** What one send collected: the images, how the outline refers to them, what didn't fit */
export interface DeckImagePlan {
  /** Multimodal inputs in send order */
  images: AgentImage[]
  /**
   * Collected pictures per slide in walk order. Ordinals are 1-based positions
   * in the *message's* images array — the caller passes how many images (user
   * attachments) precede the deck pictures. A picture whose content is already
   * attached reuses that ordinal: the pixels are attached, just once.
   */
  ordinalsBySlide: Map<number, Array<{ elementId: string; ordinal: number }>>
  /** Pictures seen but not attached: over budget, non-whitelisted mime, or failed read */
  skipped: number
}

/**
 * One picture as an AgentImage: its dataUrl when present (no IPC round trip),
 * otherwise bytes read through the main-process bridge. Null when the picture
 * cannot be attached (unreadable, a mime the endpoint cannot take, or over
 * the per-picture budget).
 */
async function readPictureImage(
  pic: PictureRenderNode,
  slideIndex: number,
  readBytes: DeckImageReader,
): Promise<AgentImage | null> {
  let base64: string | undefined
  let mime: string | undefined
  const url = pic.dataUrl
  if (typeof url === 'string') {
    const comma = url.indexOf(',')
    // dataUrl head is "data:<mime>;base64" — the mime ends at the semicolon, not the comma
    if (comma > 0 && /^data:[^;]+;base64$/i.test(url.slice(0, comma))) {
      base64 = url.slice(comma + 1)
      mime = url.slice(5, url.indexOf(';', 5))
    }
  }
  if (base64 === undefined || mime === undefined) {
    try {
      const r = await readBytes({ slideIndex, elementId: pic.durableId ?? pic.sourceId })
      if (!r.ok || !r.base64 || !r.mime) return null
      base64 = r.base64
      mime = r.mime
    } catch {
      return null
    }
  }
  if (!DECK_IMAGE_MIME[mime]) return null
  if (base64.length > MAX_DECK_IMAGE_B64) return null
  return { base64, mime }
}

/** Depth-first walk of a slide's nodes yielding picture nodes in z-order (group children included) */
function* walkPictures(nodes: RenderNode[]): Generator<PictureRenderNode> {
  for (const n of nodes) {
    if (n.type === 'picture') yield n
    else if (n.type === 'group') yield* walkPictures(n.children)
  }
}

/**
 * Collect the deck's pictures for one outgoing message.
 *
 * Walk order = outline order (slide order, z-order within a slide), so the
 * first attached picture is the first the outline annotates. Pictures marked
 * `redact` are never collected — withholding a picture means the model must
 * not receive its pixels any more than its reference.
 */
export async function collectDeckImages(
  slides: RenderSlide[],
  readBytes: DeckImageReader,
  /** How many images (user attachments) precede the deck pictures in the message */
  offset = 0,
): Promise<DeckImagePlan> {
  const images: AgentImage[] = []
  const ordinalsBySlide = new Map<number, Array<{ elementId: string; ordinal: number }>>()
  let totalB64 = 0
  let skipped = 0
  const attach = (slideIndex: number, elementId: string, ordinal: number) => {
    const list = ordinalsBySlide.get(slideIndex)
    if (list) list.push({ elementId, ordinal })
    else ordinalsBySlide.set(slideIndex, [{ elementId, ordinal }])
  }
  for (let slideIndex = 0; slideIndex < slides.length; slideIndex++) {
    const slide = slides[slideIndex]
    if (!slide) continue
    for (const pic of walkPictures(slide.nodes)) {
      if (pic.redact) continue
      const image = await readPictureImage(pic, slideIndex, readBytes)
      if (!image) {
        skipped++
        continue
      }
      const elementId = pic.durableId ?? pic.sourceId
      // content dedup: the same bitmap on many pages attaches once; a later
      // copy still annotates with the ordinal its pixels sit at
      const dup = images.findIndex((a) => a.mime === image.mime && a.base64 === image.base64)
      if (dup >= 0) {
        attach(slideIndex, elementId, offset + dup + 1)
        continue
      }
      if (
        images.length >= MAX_DECK_IMAGES ||
        totalB64 + image.base64.length > MAX_DECK_IMAGES_B64
      ) {
        skipped++
        continue
      }
      images.push(image)
      totalB64 += image.base64.length
      attach(slideIndex, elementId, offset + images.length)
    }
  }
  return { images, ordinalsBySlide, skipped }
}
