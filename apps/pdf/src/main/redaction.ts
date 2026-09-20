import { createHash } from 'node:crypto'
import { chainPdfium, loadPdfium, saveDoc, withDocument } from './text-edit'
import type { RedactionInput } from '../shared/ipc'

/** PDF user-space rectangle. Callers retain original page indices until the redaction copy is made. */
export type RedactionRegion = RedactionInput

// public/fpdf_annot.h: FPDF_ANNOT_REDACT. Keep this local instead of importing
// a viewer enum: this file calls the C/WASM surface directly.
const FPDF_ANNOT_REDACT = 28
const FPDF_ANNOT_WIDGET = 20
const FPDF_PAGEOBJ_IMAGE = 3
const FPDF_PAGEOBJ_FORM = 5
// @embedpdf/models PdfAnnotationColorType.OverlayColor. Unlike /IC, /OC is
// flattened by ApplyRedaction as the permanent post-removal fill.
const PDF_ANNOTATION_OVERLAY_COLOR = 2

function intersects(a: readonly number[], b: readonly number[]): boolean {
  const [al, ab, ar, at] = [
    Math.min(a[0]!, a[2]!),
    Math.min(a[1]!, a[3]!),
    Math.max(a[0]!, a[2]!),
    Math.max(a[1]!, a[3]!),
  ]
  const [bl, bb, br, bt] = [
    Math.min(b[0]!, b[2]!),
    Math.min(b[1]!, b[3]!),
    Math.max(b[0]!, b[2]!),
    Math.max(b[1]!, b[3]!),
  ]
  return al < br && ar > bl && ab < bt && at > bb
}

function objectBounds(
  m: Awaited<ReturnType<typeof loadPdfium>>,
  object: number,
): [number, number, number, number] | null {
  const left = m._malloc(4)
  const bottom = m._malloc(4)
  const right = m._malloc(4)
  const top = m._malloc(4)
  try {
    if (!m._FPDFPageObj_GetBounds(object, left, bottom, right, top)) return null
    return [
      m.HEAPF32[left >> 2]!,
      m.HEAPF32[bottom >> 2]!,
      m.HEAPF32[right >> 2]!,
      m.HEAPF32[top >> 2]!,
    ]
  } finally {
    m._free(left)
    m._free(bottom)
    m._free(right)
    m._free(top)
  }
}

function imageHash(m: Awaited<ReturnType<typeof loadPdfium>>, image: number): string {
  const size = m._FPDFImageObj_GetImageDataRaw(image, 0, 0)
  if (!size) throw new Error('PDFium could not inspect image data safely')
  const ptr = m._malloc(size)
  try {
    if (m._FPDFImageObj_GetImageDataRaw(image, ptr, size) !== size) {
      throw new Error('PDFium could not inspect image data safely')
    }
    return createHash('sha256')
      .update(m.HEAPU8.subarray(ptr, ptr + size))
      .digest('hex')
  } finally {
    m._free(ptr)
  }
}

function collectImageHashes(
  m: Awaited<ReturnType<typeof loadPdfium>>,
  object: number,
  hashes: Map<string, number>,
  seen: Set<number>,
): boolean {
  if (!object || seen.has(object)) return false
  seen.add(object)
  const type = m._FPDFPageObj_GetType(object)
  if (type === FPDF_PAGEOBJ_IMAGE) {
    const hash = imageHash(m, object)
    hashes.set(hash, (hashes.get(hash) ?? 0) + 1)
    return true
  }
  if (type !== FPDF_PAGEOBJ_FORM) return false
  let hasImage = false
  for (let i = 0; i < m._FPDFFormObj_CountObjects(object); i++) {
    hasImage = collectImageHashes(m, m._FPDFFormObj_GetObject(object, i), hashes, seen) || hasImage
  }
  return hasImage
}

/** PDFium 2.15.1 edits a shared image stream in place. Allow a single-instance
 * image, but reject a selected shared image or an image-bearing Form XObject whose
 * child transform cannot be inspected through this API. */
function assertImageRedactionIsSafe(
  m: Awaited<ReturnType<typeof loadPdfium>>,
  doc: number,
  page: number,
  rect: readonly number[],
  pageIndex: number,
): void {
  const hashes = new Map<string, number>()
  for (let p = 0; p < m._FPDF_GetPageCount(doc); p++) {
    const candidate = m._FPDF_LoadPage(doc, p)
    if (!candidate) throw new Error(`PDFium could not inspect page ${p + 1}`)
    try {
      for (let i = 0; i < m._FPDFPage_CountObjects(candidate); i++) {
        collectImageHashes(m, m._FPDFPage_GetObject(candidate, i), hashes, new Set())
      }
    } finally {
      m._FPDF_ClosePage(candidate)
    }
  }
  for (let i = 0; i < m._FPDFPage_CountObjects(page); i++) {
    const object = m._FPDFPage_GetObject(page, i)
    const bounds = objectBounds(m, object)
    if (!object || !bounds || !intersects(rect, bounds)) continue
    if (
      m._FPDFPageObj_GetType(object) === FPDF_PAGEOBJ_IMAGE &&
      (hashes.get(imageHash(m, object)) ?? 0) > 1
    ) {
      throw new Error(
        `redaction intersects a shared image on page ${pageIndex + 1}; refusing unsafe image mutation`,
      )
    }
    if (
      m._FPDFPageObj_GetType(object) === FPDF_PAGEOBJ_FORM &&
      collectImageHashes(m, object, new Map(), new Set())
    ) {
      throw new Error(
        `redaction intersects an image Form XObject on page ${pageIndex + 1}; refusing unsafe image mutation`,
      )
    }
  }
}

/** PDFium 2.15.1 removes a hit Widget annotation but can leave its AcroForm
 * field value and appearance reachable through the field tree. Refuse that
 * region until the whole field graph can be removed transactionally. */
function assertFormRedactionIsSafe(
  m: Awaited<ReturnType<typeof loadPdfium>>,
  page: number,
  rect: readonly number[],
  pageIndex: number,
): void {
  for (let i = 0; i < m._FPDFPage_GetAnnotCount(page); i++) {
    const annot = m._FPDFPage_GetAnnot(page, i)
    if (!annot) throw new Error(`PDFium could not inspect annotation on page ${pageIndex + 1}`)
    try {
      if (m._FPDFAnnot_GetSubtype(annot) !== FPDF_ANNOT_WIDGET) continue
      const ptr = m._malloc(16)
      try {
        if (!m._FPDFAnnot_GetRect(annot, ptr)) {
          throw new Error(`PDFium could not inspect form field on page ${pageIndex + 1}`)
        }
        const bounds = Array.from(m.HEAPF32.subarray(ptr >> 2, (ptr >> 2) + 4))
        if (intersects(rect, bounds)) {
          throw new Error(
            `redaction intersects a form field on page ${pageIndex + 1}; refusing recoverable field data`,
          )
        }
      } finally {
        m._free(ptr)
      }
    } finally {
      m._FPDFPage_CloseAnnot(annot)
    }
  }
}

/** Remove non-widget annotations whose alternate text, comments, or appearance
 * would otherwise remain a recoverable side channel. Widgets are preflighted and
 * rejected above because their AcroForm parent values can outlive page removal. */
function removeHitAnnotations(
  m: Awaited<ReturnType<typeof loadPdfium>>,
  page: number,
  rect: readonly number[],
  pageIndex: number,
): void {
  for (let i = m._FPDFPage_GetAnnotCount(page) - 1; i >= 0; i--) {
    const annot = m._FPDFPage_GetAnnot(page, i)
    if (!annot) throw new Error(`PDFium could not inspect annotation on page ${pageIndex + 1}`)
    let hit: boolean
    try {
      const ptr = m._malloc(16)
      try {
        if (!m._FPDFAnnot_GetRect(annot, ptr)) {
          throw new Error(`PDFium could not inspect annotation on page ${pageIndex + 1}`)
        }
        hit = intersects(rect, Array.from(m.HEAPF32.subarray(ptr >> 2, (ptr >> 2) + 4)))
      } finally {
        m._free(ptr)
      }
    } finally {
      m._FPDFPage_CloseAnnot(annot)
    }
    if (hit && !m._FPDFPage_RemoveAnnot(page, i)) {
      throw new Error(`PDFium could not remove annotation on page ${pageIndex + 1}`)
    }
  }
}

const MAX_FLOAT32 = 3.402823466e38

function opLabel(index: number): string {
  return `redaction ${index + 1}`
}

function formatRect(rect: unknown): string {
  try {
    return JSON.stringify(rect) ?? String(rect)
  } catch {
    return String(rect)
  }
}

/** Shared page-index gate: integer and non-negative. Upper-bound range is checked later with pageCount. */
function checkPageIndex(pageIndex: unknown, label: string): number {
  if (typeof pageIndex !== 'number' || !Number.isInteger(pageIndex)) {
    throw new Error(
      `${label}: redaction page index must be an integer (got ${formatRect(pageIndex)})`,
    )
  }
  if (pageIndex < 0) {
    throw new Error(`${label}: redaction page index ${pageIndex} is out of range (must be >= 0)`)
  }
  return pageIndex
}

/** Shared rectangle gate: four finite coords inside PDFium range and nonempty. */
function checkRect(
  rect: unknown,
  label: string,
  pageIndex: number,
): [number, number, number, number] {
  const where = `${label} on page ${pageIndex + 1}`
  if (!Array.isArray(rect) || rect.length !== 4 || !rect.every(Number.isFinite)) {
    throw new Error(
      `${where}: redaction rectangle must contain four finite coordinates (got ${formatRect(rect)})`,
    )
  }
  const [x1, y1, x2, y2] = rect as [number, number, number, number]
  if (rect.some((n) => Math.abs(n) > MAX_FLOAT32)) {
    throw new Error(
      `${where}: redaction rectangle is outside PDFium coordinate range (got ${formatRect(rect)})`,
    )
  }
  if (!(x1 !== x2 && y1 !== y2)) {
    throw new Error(`${where}: redaction rectangle must be nonempty (got ${formatRect(rect)})`)
  }
  return [x1, y1, x2, y2]
}

export function validateRedactionRegions(value: unknown): RedactionRegion[] {
  if (!Array.isArray(value) || value.length === 0)
    throw new Error('at least one redaction rectangle is required')
  return value.map((region, index) => {
    const label = opLabel(index)
    if (!region || typeof region !== 'object')
      throw new Error(`${label}: invalid redaction region (got ${formatRect(region)})`)
    const input = region as Partial<RedactionRegion>
    const pageIndex = checkPageIndex(input.pageIndex, label)
    const rect = checkRect(input.rect, label, pageIndex)
    return { pageIndex, rect }
  })
}

function checkedRegion(
  region: RedactionRegion,
  pageCount: number,
  index = 0,
): [number, number, number, number] {
  const label = opLabel(index)
  const pageIndex = checkPageIndex(region.pageIndex, label)
  if (pageIndex >= pageCount) {
    throw new Error(
      `${label}: redaction page index ${pageIndex} is out of range (pages 1-${pageCount}, got rect ${formatRect(region.rect)})`,
    )
  }
  const [x1, y1, x2, y2] = checkRect(region.rect, label, pageIndex)
  const left = Math.min(x1, x2)
  const bottom = Math.min(y1, y2)
  const right = Math.max(x1, x2)
  const top = Math.max(y1, y2)
  if (!(right > left && top > bottom))
    throw new Error(
      `${label} on page ${pageIndex + 1}: redaction rectangle must be nonempty (got ${formatRect(region.rect)})`,
    )
  return [left, top, right, bottom]
}

/**
 * Permanently apply every requested area using EmbedPDF's native redaction API.
 * This deliberately does not persist the temporary Redact annotations: native code
 * removes both them and intersecting annotations after deleting the covered content.
 */
export function redactPdf(bytes: Uint8Array, regions: RedactionRegion[]): Promise<Uint8Array> {
  let validRegions: RedactionRegion[]
  try {
    validRegions = validateRedactionRegions(regions)
  } catch (err) {
    return Promise.reject(err)
  }
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async (doc) => {
      const pageCount = m._FPDF_GetPageCount(doc)
      // Validate the complete request before constructing a single annotation, so an
      // invalid trailing region cannot yield a partially-redacted output buffer.
      const checked = validRegions.map((region, index) => ({
        ...region,
        nativeRect: checkedRegion(region, pageCount, index),
      }))
      // Do every fail-closed check before creating annotations, preserving all-or-nothing
      // behavior even when one later page contains a shared image placement.
      for (const region of checked) {
        const page = m._FPDF_LoadPage(doc, region.pageIndex)
        if (!page) throw new Error(`PDFium could not load redaction page ${region.pageIndex + 1}`)
        try {
          assertImageRedactionIsSafe(m, doc, page, region.nativeRect, region.pageIndex)
          assertFormRedactionIsSafe(m, page, region.nativeRect, region.pageIndex)
        } finally {
          m._FPDF_ClosePage(page)
        }
      }
      for (const region of checked) {
        const page = m._FPDF_LoadPage(doc, region.pageIndex)
        if (!page) throw new Error(`PDFium could not load redaction page ${region.pageIndex + 1}`)
        try {
          removeHitAnnotations(m, page, region.nativeRect, region.pageIndex)
          const annot = m._FPDFPage_CreateAnnot(page, FPDF_ANNOT_REDACT)
          if (!annot)
            throw new Error(`PDFium could not create redaction on page ${region.pageIndex + 1}`)
          try {
            const ptr = m._malloc(16)
            try {
              // FS_RECTF is {left, top, right, bottom}; the native implementation
              // normalizes it and uses it when QuadPoints are absent.
              m.HEAPF32.set(region.nativeRect, ptr >> 2)
              if (!m._FPDFAnnot_SetRect(annot, ptr)) {
                throw new Error(
                  `PDFium rejected redaction rectangle on page ${region.pageIndex + 1}`,
                )
              }
              // /OC is flattened as page content during ApplyRedaction. /IC only
              // controls the review annotation and would leave a white hole here.
              if (!m._EPDFAnnot_SetColor(annot, PDF_ANNOTATION_OVERLAY_COLOR, 0, 0, 0)) {
                throw new Error(`PDFium could not style redaction on page ${region.pageIndex + 1}`)
              }
            } finally {
              m._free(ptr)
            }
            if (!m._EPDFAnnot_ApplyRedaction(page, annot)) {
              throw new Error(`PDFium could not apply redaction on page ${region.pageIndex + 1}`)
            }
            // 2.15.1 removes image/text content correctly but its annotation overlay
            // is not painted when a Rect (rather than QuadPoints) defines the area.
            // Add the blackout to the *already redacted native page* before its sole
            // GenerateContent/save pass; this never serializes the pre-redaction doc.
            const [left, top, right, bottom] = region.nativeRect
            const blackout = m._FPDFPageObj_CreateNewRect(left, bottom, right - left, top - bottom)
            if (
              !blackout ||
              !m._FPDFPageObj_SetFillColor(blackout, 0, 0, 0, 255) ||
              !m._FPDFPath_SetDrawMode(blackout, 1, 0)
            ) {
              if (blackout) m._FPDFPageObj_Destroy(blackout)
              throw new Error(`PDFium could not create blackout on page ${region.pageIndex + 1}`)
            }
            m._FPDFPage_InsertObject(page, blackout)
          } finally {
            m._FPDFPage_CloseAnnot(annot)
          }
          if (!m._FPDFPage_GenerateContent(page)) {
            throw new Error(`PDFium could not serialize redaction on page ${region.pageIndex + 1}`)
          }
        } finally {
          m._FPDF_ClosePage(page)
        }
      }
      return saveDoc(m, doc)
    })
  })
}
