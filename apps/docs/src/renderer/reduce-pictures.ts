import {
  croppedPixelRect,
  planMediaReduction,
  rasterInfo,
  reducedPixelSize,
  replaceDocxParts,
  tiffToDataUrlAsync,
  type MediaReductionTarget,
  type RasterInfo,
} from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'
import { NodeSelection, type EditorState } from '@tiptap/pm/state'
import { selectedObjectPositions } from './editor/object-multi-select'

/** Word's picture qualities: for printing, for the screen, for e-mail */
export const REDUCE_PPI = [220, 150, 96] as const
export type ReducePpi = (typeof REDUCE_PPI)[number]

export interface ReduceOptions {
  /** null keeps the resolution the pictures are stored at */
  ppi: ReducePpi | null
  /** cut every cropped picture down to the part it shows */
  deleteCropped: boolean
  /** the pictures to work on, as the editor holds them (data or lazy-media URLs); absent for all */
  only?: readonly string[]
}

export interface ReduceOutcome {
  /** picture parts made smaller, or cut down to what their crop shows */
  reduced: number
  /** parts already small enough, or that came out no smaller */
  smallEnough: number
  /** parts left alone: no pixels to resample, a size that cannot be read, or bytes a canvas would change */
  untouched: number
  bytesBefore: number
  bytesAfter: number
}

const JPEG_QUALITY = 0.9

async function partBytes(target: MediaReductionTarget): Promise<Uint8Array | null> {
  if (target.data) return target.data
  if (!target.lazyUrl) return null
  const res = await fetch(target.lazyUrl)
  return res.ok ? new Uint8Array(await res.arrayBuffer()) : null
}

interface Redraw {
  width: number
  height: number
  /** the stored pixels to keep; absent for the whole picture */
  source?: { sx: number; sy: number; sw: number; sh: number }
}

/** `mime` is what the bytes are, `output` what the redrawn picture is stored as */
async function resample(
  data: Uint8Array,
  mime: string,
  size: Redraw,
  output: string = mime,
): Promise<Uint8Array | null> {
  let bitmap: ImageBitmap
  try {
    const blob = new Blob([data as BlobPart], { type: mime })
    const resize: ImageBitmapOptions = {
      resizeWidth: size.width,
      resizeHeight: size.height,
      resizeQuality: 'high',
    }
    const { source } = size
    bitmap = source
      ? await createImageBitmap(blob, source.sx, source.sy, source.sw, source.sh, resize)
      : await createImageBitmap(blob, resize)
  } catch {
    return null
  }
  try {
    const canvas = new OffscreenCanvas(size.width, size.height)
    const ctx = canvas.getContext('2d')
    if (!ctx) return null
    ctx.drawImage(bitmap, 0, 0)
    const blob = await canvas.convertToBlob({ type: output, quality: JPEG_QUALITY })
    return blob.type === output ? new Uint8Array(await blob.arrayBuffer()) : null
  } finally {
    bitmap.close()
  }
}

function dataUrlBytes(url: string): Uint8Array | null {
  const m = /^data:[^;,]*;base64,(.*)$/s.exec(url)
  if (!m) return null
  try {
    return Uint8Array.from(atob(m[1]), (c) => c.charCodeAt(0))
  } catch {
    return null
  }
}

const sameBytes = (a: Uint8Array | Uint8ClampedArray, b: Uint8Array | Uint8ClampedArray): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i])

async function pixelsOf(png: Uint8Array): Promise<Uint8ClampedArray | null> {
  try {
    const bitmap = await createImageBitmap(new Blob([png as BlobPart], { type: 'image/png' }))
    try {
      const ctx = new OffscreenCanvas(bitmap.width, bitmap.height).getContext('2d')
      if (!ctx) return null
      ctx.drawImage(bitmap, 0, 0)
      return ctx.getImageData(0, 0, bitmap.width, bitmap.height).data
    } finally {
      bitmap.close()
    }
  } catch {
    return null
  }
}

/** whether a part of the package is one of the pictures the editor named */
function partPicker(only: readonly string[]): (target: MediaReductionTarget) => Promise<boolean> {
  const urls = new Set(only)
  const blobs = only.map(dataUrlBytes).filter((b): b is Uint8Array => b !== null)
  // the editor holds a TIFF as the PNG it was decoded to: the part is found by decoding it again,
  // which is only worth it while a PNG is among the selection
  const pngs = blobs.filter((blob) => rasterInfo(blob)?.mime === 'image/png')
  const pngPixels = new Map<Uint8Array, Promise<Uint8ClampedArray | null>>()
  const pixelsOfSelected = (blob: Uint8Array): Promise<Uint8ClampedArray | null> => {
    let pixels = pngPixels.get(blob)
    if (!pixels) pngPixels.set(blob, (pixels = pixelsOf(blob)))
    return pixels
  }
  const shownAs = async (target: MediaReductionTarget): Promise<boolean> => {
    if (!pngs.length) return false
    const data = await partBytes(target).catch(() => null)
    const png = data && dataUrlBytes((await tiffToDataUrlAsync(data).catch(() => null)) ?? '')
    if (!png) return false
    if (pngs.some((blob) => sameBytes(blob, png))) return true
    // the same pixels encoded on another thread need not be the same bytes
    const pixels = await pixelsOf(png)
    if (!pixels) return false
    for (const blob of pngs) {
      const other = await pixelsOfSelected(blob)
      if (other && sameBytes(other, pixels)) return true
    }
    return false
  }
  return async (target) =>
    target.mime === 'image/tiff'
      ? shownAs(target)
      : target.lazyUrl
        ? urls.has(target.lazyUrl)
        : !!target.data && blobs.some((blob) => sameBytes(blob, target.data!))
}

/** a TIFF as the PNG the page shows it as: a canvas cannot read the TIFF itself */
async function decodable(
  data: Uint8Array,
  info: RasterInfo,
): Promise<{ data: Uint8Array; info: RasterInfo } | null> {
  if (info.mime !== 'image/tiff') return { data, info }
  const png = dataUrlBytes((await tiffToDataUrlAsync(data).catch(() => null)) ?? '')
  const decoded = png && rasterInfo(png)
  return png && decoded ? { data: png, info: decoded } : null
}

/**
 * The document with its picture parts resampled down to what their largest
 * use needs at the chosen quality, and cut down to what their crop shows.
 * PNG and JPEG stay what they are; one-frame GIF, BMP, WebP and TIFF are
 * stored again as PNG (JPEG for a lossy opaque WebP). `bytes` is null when
 * no part got smaller.
 */
export async function reducePictures(
  docx: Uint8Array,
  options: ReduceOptions,
): Promise<{ bytes: Uint8Array | null; outcome: ReduceOutcome }> {
  const plan = await planMediaReduction(docx, options.ppi, {
    deleteCropped: options.deleteCropped,
  })
  const picked = options.only ? partPicker(options.only) : null
  const outcome: ReduceOutcome = {
    reduced: 0,
    smallEnough: 0,
    untouched: picked ? 0 : plan.kept.length,
    bytesBefore: 0,
    bytesAfter: 0,
  }
  const parts: Record<string, Uint8Array> = {}
  const uncropped: string[] = []
  const stored: Record<string, RasterInfo['output']> = {}
  for (const target of plan.targets) {
    if (picked && !(await picked(target))) continue
    const data = await partBytes(target).catch(() => null)
    const read = data && rasterInfo(data)
    const source =
      data && read && read.mime === target.mime && read.resamplable
        ? await decodable(data, read)
        : null
    if (!data || !read || !source) {
      outcome.untouched++
      continue
    }
    const { info } = source
    const converts = read.output !== read.mime
    outcome.bytesBefore += data.length
    const { crop } = target
    let size: Redraw | null
    if (crop) {
      const { width, height, ...source } = croppedPixelRect(info, { ...target, crop })
      size = { width, height, source }
    } else {
      // lossless pixels stored again as PNG are worth it at the size they have
      size =
        reducedPixelSize(info, target) ?? (converts && read.output === 'image/png' ? info : null)
    }
    const redrawn = size && (await resample(source.data, info.mime, size, read.output))
    // resampling is only worth keeping when it saves bytes; a crop is asked for so
    // that what it hides is gone from the file, whatever the bytes come to
    if (redrawn && (crop || redrawn.length < data.length)) {
      parts[target.partPath] = redrawn
      if (converts) stored[target.partPath] = read.output
      if (crop) uncropped.push(target.partPath)
      outcome.reduced++
      outcome.bytesAfter += redrawn.length
    } else {
      outcome.smallEnough++
      outcome.bytesAfter += data.length
    }
  }
  return {
    bytes: outcome.reduced > 0 ? await replaceDocxParts(docx, parts, uncropped, stored) : null,
    outcome,
  }
}

/**
 * The pictures of the selection as the editor holds them: a picked object,
 * the objects picked together, or the pictures inside a selected stretch.
 */
export function selectedPictureSources(state: EditorState): string[] {
  const out = new Set<string>()
  const take = (node: PmNode | null | undefined): void => {
    if (!node) return
    const attrs = node.attrs
    if (node.type.name === 'docInlineImage') {
      if (typeof attrs.dataUrl === 'string' && attrs.dataUrl) out.add(attrs.dataUrl)
    } else if (node.type.name === 'docProtected' && attrs.blockType === 'image') {
      const pending = (attrs.imageReplace ?? attrs.genImage) as {
        base64: string
        mime: string
      } | null
      if (pending) out.add(`data:${pending.mime};base64,${pending.base64}`)
      else if (typeof attrs.imageDataUrl === 'string' && attrs.imageDataUrl)
        out.add(attrs.imageDataUrl)
    }
  }
  for (const pos of selectedObjectPositions(state)) take(state.doc.nodeAt(pos))
  const { selection } = state
  if (selection instanceof NodeSelection) take(selection.node)
  else if (!selection.empty) state.doc.nodesBetween(selection.from, selection.to, (n) => take(n))
  return [...out]
}
