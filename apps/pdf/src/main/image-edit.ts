import { nativeImage } from 'electron'
import { FPDF_PAGEOBJ_TEXT, chainPdfium, loadPdfium, saveDoc, withDocument } from './text-edit'
import type { Pdfium } from './text-edit'
import type {
  ImageEditFailure,
  ImageEditInput,
  PageImageRef,
  PagePreviewRequest,
} from '../shared/ipc'

const FPDF_PAGEOBJ_IMAGE = 3
const FPDF_BITMAP_BGRA = 4

type Rect = readonly [number, number, number, number]

interface PageObj {
  obj: number
  index: number
  type: number
  bounds: [number, number, number, number]
}

function collectObjects(m: Pdfium, page: number): PageObj[] {
  const out: PageObj[] = []
  const bl = m._malloc(4)
  const bb = m._malloc(4)
  const br = m._malloc(4)
  const bt = m._malloc(4)
  const count = m._FPDFPage_CountObjects(page)
  for (let i = 0; i < count; i++) {
    const obj = m._FPDFPage_GetObject(page, i)
    if (!m._FPDFPageObj_GetBounds(obj, bl, bb, br, bt)) continue
    out.push({
      obj,
      index: i,
      type: m._FPDFPageObj_GetType(obj),
      bounds: [m.HEAPF32[bl >> 2]!, m.HEAPF32[bb >> 2]!, m.HEAPF32[br >> 2]!, m.HEAPF32[bt >> 2]!],
    })
  }
  for (const p of [bl, bb, br, bt]) m._free(p)
  return out
}

/** Bounds travel renderer→main unchanged, so matching is tight; the tolerance only
    absorbs float32 round-trips. Closest candidate within tolerance wins. */
const MATCH_TOL = 2

function matchImage(objects: PageObj[], rect: Rect): PageObj | null {
  let best: PageObj | null = null
  let bestD = Infinity
  for (const o of objects) {
    if (o.type !== FPDF_PAGEOBJ_IMAGE) continue
    const d = Math.max(
      Math.abs(o.bounds[0] - rect[0]),
      Math.abs(o.bounds[1] - rect[1]),
      Math.abs(o.bounds[2] - rect[2]),
      Math.abs(o.bounds[3] - rect[3]),
    )
    if (d <= MATCH_TOL && d < bestD) {
      best = o
      bestD = d
    }
  }
  return best
}

const firstTextIndex = (objects: PageObj[]): number | undefined =>
  objects.find((o) => o.type === FPDF_PAGEOBJ_TEXT)?.index

/** Decode a PNG/JPEG into straight-alpha BGRA (pdfium splits alpha into an SMask itself) */
function decodeImage(b64: string): { width: number; height: number; bgra: Buffer } {
  const img = nativeImage.createFromBuffer(Buffer.from(b64, 'base64'))
  if (img.isEmpty()) throw new Error('could not decode the image data')
  const { width, height } = img.getSize()
  // toBitmap() hands back premultiplied BGRA; un-premultiply so translucent pixels
  // keep their color once pdfium separates them into image + SMask
  const bgra = Buffer.from(img.toBitmap())
  for (let i = 0; i < bgra.length; i += 4) {
    const a = bgra[i + 3]!
    if (a > 0 && a < 255) {
      bgra[i] = Math.min(255, Math.round((bgra[i]! * 255) / a))
      bgra[i + 1] = Math.min(255, Math.round((bgra[i + 1]! * 255) / a))
      bgra[i + 2] = Math.min(255, Math.round((bgra[i + 2]! * 255) / a))
    }
  }
  return { width, height, bgra }
}

/** Content matrix for a footprint rect, counter-rotated against the page's display
    rotation so the image shows upright (same table as save-pdf's addImageStamp) */
function imageMatrix(rect: Rect, rotate: number): number[] {
  const [x1, y1, x2, y2] = rect
  const w = x2 - x1
  const h = y2 - y1
  switch (((rotate % 360) + 360) % 360) {
    case 90:
      return [0, h, -w, 0, x2, y1]
    case 180:
      return [-w, 0, 0, -h, x2, y2]
    case 270:
      return [0, -h, w, 0, x1, y2]
    default:
      return [w, 0, 0, h, x1, y1]
  }
}

function insertImage(m: Pdfium, doc: number, page: number, edit: ImageEditInput): void {
  if (edit.kind !== 'insertImage') return
  const { width, height, bgra } = decodeImage(edit.image)
  const bufPtr = m._malloc(bgra.length)
  m.HEAPU8.set(bgra, bufPtr)
  const bmp = m._FPDFBitmap_CreateEx(width, height, FPDF_BITMAP_BGRA, bufPtr, width * 4)
  if (!bmp) {
    m._free(bufPtr)
    throw new Error('FPDFBitmap_CreateEx failed')
  }
  const imgObj = m._FPDFPageObj_NewImageObj(doc)
  const ok = imgObj && m._FPDFImageObj_SetBitmap(0, 0, imgObj, bmp)
  m._FPDFBitmap_Destroy(bmp)
  m._free(bufPtr)
  if (!ok) {
    if (imgObj) m._FPDFPageObj_Destroy(imgObj)
    throw new Error('FPDFImageObj_SetBitmap failed')
  }
  const matPtr = m._malloc(24)
  m.HEAPF32.set(imageMatrix(edit.rect, edit.rotate ?? 0), matPtr >> 2)
  m._FPDFPageObj_SetMatrix(imgObj, matPtr)
  m._free(matPtr)
  const textIdx = edit.layer === 'belowText' ? firstTextIndex(collectObjects(m, page)) : undefined
  if (textIdx !== undefined && m._FPDFPage_InsertObjectAtIndex) {
    if (!m._FPDFPage_InsertObjectAtIndex(page, imgObj, textIdx)) {
      m._FPDFPageObj_Destroy(imgObj)
      throw new Error('FPDFPage_InsertObjectAtIndex failed')
    }
  } else {
    m._FPDFPage_InsertObject(page, imgObj)
  }
}

/** Move an object to the requested z-band (remove + reinsert; the page owns it again after) */
function moveToLayer(m: Pdfium, page: number, obj: number, layer: 'belowText' | 'aboveText'): void {
  if (!m._FPDFPage_RemoveObject(page, obj)) throw new Error('FPDFPage_RemoveObject failed')
  const textIdx = layer === 'belowText' ? firstTextIndex(collectObjects(m, page)) : undefined
  if (textIdx !== undefined && m._FPDFPage_InsertObjectAtIndex) {
    if (!m._FPDFPage_InsertObjectAtIndex(page, obj, textIdx)) {
      // The object is already off the page here — re-attach it (topmost) instead of
      // destroying it, or a skipped layer change would silently delete page content
      m._FPDFPage_InsertObject(page, obj)
      throw new Error('FPDFPage_InsertObjectAtIndex failed')
    }
  } else {
    m._FPDFPage_InsertObject(page, obj)
  }
}

/** Screen-clockwise quarter turns as PDF user-space matrices (y-up flips handedness) */
const QUARTER_TURN_MATS: ReadonlyArray<readonly [number, number, number, number]> = [
  [1, 0, 0, 1],
  [0, -1, 1, 0],
  [-1, 0, 0, -1],
  [0, 1, -1, 0],
]

/** Optional center rotation, then the axis-aligned map old bounds → final rect
    (content orientation otherwise survives, as before) */
function applyImageGeometry(
  m: Pdfium,
  obj: number,
  bounds: Rect,
  rect: Rect,
  quarterTurns?: number,
): void {
  let [ox1, oy1, ox2, oy2] = bounds
  const turns = (((quarterTurns ?? 0) % 4) + 4) % 4
  if (turns) {
    const cx = (ox1 + ox2) / 2
    const cy = (oy1 + oy2) / 2
    const [a, b, c, d] = QUARTER_TURN_MATS[turns]!
    m._FPDFPageObj_Transform(obj, a, b, c, d, cx - a * cx - c * cy, cy - b * cx - d * cy)
    if (turns % 2 === 1) {
      const w = ox2 - ox1
      const h = oy2 - oy1
      ;[ox1, oy1, ox2, oy2] = [cx - h / 2, cy - w / 2, cx + h / 2, cy + w / 2]
    }
  }
  const [nx1, ny1, nx2, ny2] = rect
  const ow = ox2 - ox1
  const oh = oy2 - oy1
  if (ow > 1e-6 && oh > 1e-6 && (nx1 !== ox1 || ny1 !== oy1 || nx2 !== ox2 || ny2 !== oy2)) {
    const sx = (nx2 - nx1) / ow
    const sy = (ny2 - ny1) / oh
    // Post-multiplied axis-aligned map old bounds → new rect
    m._FPDFPageObj_Transform(obj, sx, 0, 0, sy, nx1 - sx * ox1, ny1 - sy * oy1)
  }
}

function transformImage(m: Pdfium, page: number, edit: ImageEditInput): void {
  if (edit.kind !== 'transformImage') return
  const target = matchImage(collectObjects(m, page), edit.oldRect)
  if (!target) throw new Error('the image could not be located on the page')
  // Layer move first: it is the only part that can fail, so a skipped edit is
  // reported before any geometry has been touched
  if (edit.layer) moveToLayer(m, page, target.obj, edit.layer)
  applyImageGeometry(m, target.obj, target.bounds, edit.rect, edit.quarterTurns)
}

/** Swap the image object's pixels for a new bitmap (matrix untouched: the footprint,
    z-order and any original rotation stay), then apply the shared geometry pass */
function replaceImage(m: Pdfium, page: number, edit: ImageEditInput): void {
  if (edit.kind !== 'replaceImage') return
  const target = matchImage(collectObjects(m, page), edit.oldRect)
  if (!target) throw new Error('the image could not be located on the page')
  if (edit.layer) moveToLayer(m, page, target.obj, edit.layer)
  const { width, height, bgra } = decodeImage(edit.image)
  const bufPtr = m._malloc(bgra.length)
  m.HEAPU8.set(bgra, bufPtr)
  const bmp = m._FPDFBitmap_CreateEx(width, height, FPDF_BITMAP_BGRA, bufPtr, width * 4)
  if (!bmp) {
    m._free(bufPtr)
    throw new Error('FPDFBitmap_CreateEx failed')
  }
  // Pass the owning page so pdfium drops its cached render of the old pixels
  const pagesPtr = m._malloc(4)
  m.HEAP32[pagesPtr >> 2] = page
  const ok = m._FPDFImageObj_SetBitmap(pagesPtr, 1, target.obj, bmp)
  m._free(pagesPtr)
  m._FPDFBitmap_Destroy(bmp)
  m._free(bufPtr)
  if (!ok) throw new Error('FPDFImageObj_SetBitmap failed')
  applyImageGeometry(m, target.obj, target.bounds, edit.rect, edit.quarterTurns)
}

function deleteImage(m: Pdfium, page: number, edit: ImageEditInput): void {
  if (edit.kind !== 'deleteImage') return
  const target = matchImage(collectObjects(m, page), edit.oldRect)
  if (!target) throw new Error('the image could not be located on the page')
  if (!m._FPDFPage_RemoveObject(page, target.obj)) throw new Error('FPDFPage_RemoveObject failed')
  m._FPDFPageObj_Destroy(target.obj)
}

export interface ImageEditsResult {
  bytes: Uint8Array
  skipped: ImageEditFailure[]
}

/** Apply content-stream image ops, fail-soft per edit (mirrors applyTextEdits) */
export function applyImageEdits(
  bytes: Uint8Array,
  edits: ImageEditInput[],
): Promise<ImageEditsResult> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    const skipped: ImageEditFailure[] = []
    return withDocument(m, bytes, async (doc) => {
      const pageCount = m._FPDF_GetPageCount(doc)
      const byPage = new Map<number, { edit: ImageEditInput; editIndex: number }[]>()
      for (const [editIndex, edit] of edits.entries()) {
        if (edit.pageIndex < 0 || edit.pageIndex >= pageCount) {
          skipped.push({ editIndex, pageIndex: edit.pageIndex, reason: 'page does not exist' })
          continue
        }
        byPage.set(edit.pageIndex, [...(byPage.get(edit.pageIndex) ?? []), { edit, editIndex }])
      }
      let appliedTotal = 0
      for (const [pageIndex, pageEdits] of byPage) {
        const page = m._FPDF_LoadPage(doc, pageIndex)
        if (!page) throw new Error(`could not load page ${pageIndex + 1}`)
        try {
          let applied = 0
          // Objects are re-collected inside each op: every mutation shifts indices
          for (const { edit, editIndex } of pageEdits) {
            try {
              if (edit.kind === 'insertImage') insertImage(m, doc, page, edit)
              else if (edit.kind === 'transformImage') transformImage(m, page, edit)
              else if (edit.kind === 'replaceImage') replaceImage(m, page, edit)
              else deleteImage(m, page, edit)
              applied++
            } catch (err) {
              skipped.push({
                editIndex,
                pageIndex,
                reason: err instanceof Error ? err.message : String(err),
              })
            }
          }
          if (applied > 0 && !m._FPDFPage_GenerateContent(page)) {
            throw new Error(`could not regenerate page ${pageIndex + 1}`)
          }
          appliedTotal += applied
        } finally {
          m._FPDF_ClosePage(page)
        }
      }
      return { bytes: appliedTotal > 0 ? saveDoc(m, doc) : bytes, skipped }
    })
  })
}

/** Enumerate every page's content-stream images (rect + z-band relative to text) */
export function listPageImages(bytes: Uint8Array): Promise<PageImageRef[]> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async (doc) => {
      const out: PageImageRef[] = []
      const pageCount = m._FPDF_GetPageCount(doc)
      for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
        const page = m._FPDF_LoadPage(doc, pageIndex)
        if (!page) continue
        try {
          const objects = collectObjects(m, page)
          const textIdx = firstTextIndex(objects)
          for (const o of objects) {
            if (o.type !== FPDF_PAGEOBJ_IMAGE) continue
            // Sub-3pt objects are decoration/artifacts; not worth an edit target
            if (o.bounds[2] - o.bounds[0] < 3 || o.bounds[3] - o.bounds[1] < 3) continue
            out.push({
              pageIndex,
              rect: o.bounds,
              aboveText: textIdx !== undefined && o.index > textIdx,
            })
          }
        } finally {
          m._FPDF_ClosePage(page)
        }
      }
      return out
    })
  })
}

/** Render one existing image object to PNG (base64) for the renderer's ghost preview */
export function renderImagePng(
  bytes: Uint8Array,
  pageIndex: number,
  rect: Rect,
): Promise<string | null> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async (doc) => {
      const page = m._FPDF_LoadPage(doc, pageIndex)
      if (!page) return null
      try {
        const target = matchImage(collectObjects(m, page), rect)
        if (!target) return null
        const bmp = m._FPDFImageObj_GetRenderedBitmap(doc, page, target.obj)
        if (!bmp) return null
        try {
          const w = m._FPDFBitmap_GetWidth(bmp)
          const h = m._FPDFBitmap_GetHeight(bmp)
          const stride = m._FPDFBitmap_GetStride(bmp)
          const buf = m._FPDFBitmap_GetBuffer(bmp)
          if (!w || !h || !buf) return null
          const tight = Buffer.alloc(w * h * 4)
          for (let row = 0; row < h; row++) {
            tight.set(
              m.HEAPU8.subarray(buf + row * stride, buf + row * stride + w * 4),
              row * w * 4,
            )
          }
          const png = nativeImage.createFromBitmap(tight, { width: w, height: h }).toPNG()
          return png.toString('base64')
        } finally {
          m._FPDFBitmap_Destroy(bmp)
        }
      } finally {
        m._FPDF_ClosePage(page)
      }
    })
  })
}

/**
 * Live-preview render: rasterize a page region with the given image objects removed
 * (in memory only — the file is untouched). The renderer patches the region over its
 * PDF.js canvas so a moved/resized/deleted image disappears immediately instead of
 * lingering until save. Unmatched rects are skipped fail-soft.
 */
export function renderPagePreviewPng(
  bytes: Uint8Array,
  request: Omit<PagePreviewRequest, 'path'>,
): Promise<string | null> {
  const { pageIndex, excludeRects, clip, pxWidth, rotate } = request
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async (doc) => {
      const page = m._FPDF_LoadPage(doc, pageIndex)
      if (!page) return null
      try {
        for (const rect of excludeRects) {
          const target = matchImage(collectObjects(m, page), rect)
          if (target && m._FPDFPage_RemoveObject(page, target.obj)) {
            m._FPDFPageObj_Destroy(target.obj)
          }
        }
        // Page size in display orientation (pdfium already applies /Rotate; the
        // unsaved delta passed as quarter turns swaps the axes again when odd)
        const baseW = m._FPDF_GetPageWidthF(page)
        const baseH = m._FPDF_GetPageHeightF(page)
        const turns = ((rotate % 4) + 4) % 4
        const dispW = turns % 2 === 1 ? baseH : baseW
        const dispH = turns % 2 === 1 ? baseW : baseH
        if (clip.width <= 0 || clip.height <= 0 || pxWidth <= 0) return null
        const k = pxWidth / clip.width
        const w = Math.max(1, Math.round(pxWidth))
        const h = Math.max(1, Math.round(clip.height * k))
        const bufPtr = m._malloc(w * h * 4)
        const bmp = m._FPDFBitmap_CreateEx(w, h, FPDF_BITMAP_BGRA, bufPtr, w * 4)
        if (!bmp) {
          m._free(bufPtr)
          return null
        }
        try {
          m._FPDFBitmap_FillRect(bmp, 0, 0, w, h, 0xffffffff)
          m._FPDF_RenderPageBitmap(
            bmp,
            page,
            Math.round(-clip.x * k),
            Math.round(-clip.y * k),
            Math.round(dispW * k),
            Math.round(dispH * k),
            turns,
            0,
          )
          const tight = Buffer.from(m.HEAPU8.subarray(bufPtr, bufPtr + w * h * 4))
          const png = nativeImage.createFromBitmap(tight, { width: w, height: h }).toPNG()
          return png.toString('base64')
        } finally {
          m._FPDFBitmap_Destroy(bmp)
          m._free(bufPtr)
        }
      } finally {
        m._FPDF_ClosePage(page)
      }
    })
  })
}

/**
 * Read-back verification on the final saved bytes (pageIndex already remapped to the
 * final document): inserts and transforms must be findable at their target rect.
 * Deletes are not re-checked — RemoveObject's status is verified at apply time, and a
 * second identical image at the same bounds would read as a false failure.
 */
export function verifyImageEdits(
  bytes: Uint8Array,
  edits: { pageIndex: number; rect: Rect }[],
): Promise<{ pageIndex: number; reason: string }[]> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async (doc) => {
      const failures: { pageIndex: number; reason: string }[] = []
      const pageCount = m._FPDF_GetPageCount(doc)
      const byPage = new Map<number, Rect[]>()
      for (const e of edits) {
        if (e.pageIndex < 0 || e.pageIndex >= pageCount) {
          failures.push({ pageIndex: e.pageIndex, reason: 'page missing from saved output' })
          continue
        }
        byPage.set(e.pageIndex, [...(byPage.get(e.pageIndex) ?? []), e.rect])
      }
      for (const [pageIndex, rects] of byPage) {
        const page = m._FPDF_LoadPage(doc, pageIndex)
        if (!page) {
          for (const _ of rects)
            failures.push({ pageIndex, reason: 'page unreadable in saved output' })
          continue
        }
        try {
          const objects = collectObjects(m, page)
          for (const rect of rects) {
            if (!matchImage(objects, rect)) {
              failures.push({ pageIndex, reason: 'image missing from saved output' })
            }
          }
        } finally {
          m._FPDF_ClosePage(page)
        }
      }
      return failures
    })
  })
}
