import {
  croppedPixelRect,
  planMediaReduction,
  rasterInfo,
  reducedPixelSize,
  replaceDocxParts,
  type MediaReductionTarget,
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
  /** PNG and JPEG parts already small enough, or that came out no smaller */
  smallEnough: number
  /** parts left alone: another format, a size that cannot be read, or bytes a canvas would change */
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

async function resample(data: Uint8Array, mime: string, size: Redraw): Promise<Uint8Array | null> {
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
    const blob = await canvas.convertToBlob({ type: mime, quality: JPEG_QUALITY })
    return blob.type === mime ? new Uint8Array(await blob.arrayBuffer()) : null
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

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i])

/** whether a part of the package is one of the pictures the editor named */
function partPicker(only: readonly string[]): (target: MediaReductionTarget) => boolean {
  const urls = new Set(only)
  const blobs = only.map(dataUrlBytes).filter((b): b is Uint8Array => b !== null)
  return (target) =>
    target.lazyUrl
      ? urls.has(target.lazyUrl)
      : !!target.data && blobs.some((blob) => sameBytes(blob, target.data!))
}

/**
 * The document with its PNG and JPEG parts resampled down to what their
 * largest use needs at the chosen quality, and cut down to what their crop
 * shows. `bytes` is null when no part got smaller.
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
  for (const target of plan.targets) {
    if (picked && !picked(target)) continue
    const data = await partBytes(target).catch(() => null)
    const info = data && rasterInfo(data)
    if (!data || !info || info.mime !== target.mime || !info.resamplable) {
      outcome.untouched++
      continue
    }
    outcome.bytesBefore += data.length
    const { crop } = target
    let size: Redraw | null
    if (crop) {
      const { width, height, ...source } = croppedPixelRect(info, { ...target, crop })
      size = { width, height, source }
    } else {
      size = reducedPixelSize(info, target)
    }
    const redrawn = size && (await resample(data, target.mime, size))
    // resampling is only worth keeping when it saves bytes; a crop is asked for so
    // that what it hides is gone from the file, whatever the bytes come to
    if (redrawn && (crop || redrawn.length < data.length)) {
      parts[target.partPath] = redrawn
      if (crop) uncropped.push(target.partPath)
      outcome.reduced++
      outcome.bytesAfter += redrawn.length
    } else {
      outcome.smallEnough++
      outcome.bytesAfter += data.length
    }
  }
  return {
    bytes: outcome.reduced > 0 ? await replaceDocxParts(docx, parts, uncropped) : null,
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
