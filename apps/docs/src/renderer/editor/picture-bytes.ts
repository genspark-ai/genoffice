import type { Editor } from '@tiptap/core'
import { originalPictureSize, pictureDensityOf } from './picture-density'

export async function imageSizeOf(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight })
    img.onerror = () => reject(new Error('Failed to read image'))
    img.src = dataUrl
  })
}

// enough of the file for its header segments (an Exif block is at most 64 KB)
const HEAD_BYTES = 192 * 1024

async function headBytesOf(url: string): Promise<Uint8Array> {
  const data = /^data:[^;,]+;base64,/.exec(url)
  if (!data) return new Uint8Array(await (await fetch(url)).arrayBuffer()).subarray(0, HEAD_BYTES)
  const head = atob(url.slice(data[0].length, data[0].length + (HEAD_BYTES / 3) * 4))
  return Uint8Array.from(head, (ch) => ch.charCodeAt(0))
}

/** Word's original size of a picture in px: its pixels at the density the file states */
export async function originalPictureSizeOf(
  url: string,
): Promise<{ width: number; height: number }> {
  const pixels = await imageSizeOf(url)
  const density = await headBytesOf(url).then(pictureDensityOf, () => null)
  return originalPictureSize(pixels, density)
}

/**
 * Replace the selected image's bytes (shared by Change Picture / remove background / crop).
 * Original images (docxIndex set) swap bytes in place via the imageReplace patch: the
 * drawing XML survives, so wrap/position/docxIndex — and with them the Position gallery —
 * keep working. Images not yet saved (genImage) just update their pending payload.
 * Display size keeps the current width; height adapts to the new image's aspect ratio.
 */
export async function replaceSelectedPictureBytes(
  editor: Editor,
  dataUrl: string,
  /** the size to show the new bytes at, when they are a part of the old picture */
  shown?: { widthPx: number; heightPx: number },
): Promise<void> {
  if (!editor.isEditable) return
  const m = /^data:(image\/(?:png|jpeg|gif));base64,(.*)$/s.exec(dataUrl)
  if (!m) return
  const attrs = editor.getAttributes('docProtected')
  if (attrs?.blockType !== 'image') return
  try {
    const natural = await imageSizeOf(dataUrl)
    const currentW = Number(attrs.imageWidthPx) || Math.min(natural.width, 620)
    const w = Math.max(1, Math.round(shown?.widthPx ?? currentW))
    const h = Math.max(
      1,
      Math.round(shown?.heightPx ?? (currentW * natural.height) / natural.width),
    )
    const isOriginal = attrs.docxIndex !== null && attrs.docxIndex !== undefined
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', {
        imageDataUrl: dataUrl,
        imageWidthPx: w,
        imageHeightPx: h,
        // The new bytes are the full picture (crop/cutout bake destructively) and
        // the replace pipeline strips a:srcRect on save — drop a Word-authored
        // crop/fill window or it would keep clipping the new image until reload
        imageCrop: null,
        imageFillRect: null,
        ...(isOriginal
          ? { imageReplace: { base64: m[2], mime: m[1] } }
          : { genImage: { base64: m[2], mime: m[1], widthPx: w, heightPx: h } }),
      })
      .run()
  } catch {
    /* image decode failed: keep the original untouched */
  }
}
