/**
 * Emscripten module surface this package calls into (raw FPDF_* exports +
 * heap access). Same pattern as apps/pdf/src/main/text-edit.ts: the wasm is
 * initialized by the CALLER (`init({ wasmBinary })` + `_PDFiumExt_Init()`);
 * this package never touches the filesystem. Signatures follow
 * @embedpdf/pdfium@2.15 dist/index.d.ts (verified against its export list).
 */
export interface PdfiumModule {
  HEAPU8: Uint8Array
  HEAPU32: Uint32Array
  HEAPF32: Float32Array
  HEAPF64: Float64Array
  _malloc(size: number): number
  _free(ptr: number): void

  // document / page
  _FPDF_LoadMemDocument(ptr: number, size: number, password: number): number
  _FPDF_GetLastError(): number
  _FPDF_CloseDocument(doc: number): void
  _FPDF_GetPageCount(doc: number): number
  /** Info-dictionary text (Producer/Creator/…), UTF-16LE incl. NUL; returns bytes needed */
  _FPDF_GetMetaText(doc: number, tagPtr: number, buffer: number, buflen: number): number
  _FPDF_LoadPage(doc: number, index: number): number
  _FPDF_ClosePage(page: number): void
  _FPDF_GetPageWidthF(page: number): number
  _FPDF_GetPageHeightF(page: number): number
  _FPDFPage_GetRotation(page: number): number
  /** float* left, bottom, right, top */
  _FPDFPage_GetMediaBox(page: number, l: number, b: number, r: number, t: number): number
  _FPDFPage_GetCropBox(page: number, l: number, b: number, r: number, t: number): number

  // text page: character-level access
  _FPDFText_LoadPage(page: number): number
  _FPDFText_ClosePage(textPage: number): void
  _FPDFText_CountChars(textPage: number): number
  _FPDFText_GetUnicode(textPage: number, index: number): number
  /** double* left, right, bottom, top */
  _FPDFText_GetCharBox(
    textPage: number,
    index: number,
    l: number,
    r: number,
    b: number,
    t: number,
  ): number
  /** FS_RECTF* (4 floats: left, top, right, bottom) */
  _FPDFText_GetLooseCharBox(textPage: number, index: number, rect: number): number
  /** double* x, y */
  _FPDFText_GetCharOrigin(textPage: number, index: number, x: number, y: number): number
  /** returns radians */
  _FPDFText_GetCharAngle(textPage: number, index: number): number
  _FPDFText_GetFontSize(textPage: number, index: number): number
  /** utf-8 name into buffer, int* flags; returns bytes needed incl. NUL */
  _FPDFText_GetFontInfo(
    textPage: number,
    index: number,
    buffer: number,
    buflen: number,
    flags: number,
  ): number
  _FPDFText_GetFontWeight(textPage: number, index: number): number
  /** unsigned int* R, G, B, A */
  _FPDFText_GetFillColor(
    textPage: number,
    index: number,
    r: number,
    g: number,
    b: number,
    a: number,
  ): number
  _FPDFText_GetStrokeColor(
    textPage: number,
    index: number,
    r: number,
    g: number,
    b: number,
    a: number,
  ): number
  /** FS_MATRIX* (6 floats) of the char's text object */
  _FPDFText_GetMatrix(textPage: number, index: number, matrix: number): number
  _FPDFText_IsGenerated(textPage: number, index: number): number
  _FPDFText_IsHyphen(textPage: number, index: number): number
  _FPDFText_GetTextObject(textPage: number, index: number): number
  /** FPDF_TEXTRENDERMODE_* of a page text object (3/7 = invisible) */
  _FPDFTextObj_GetTextRenderMode(textObj: number): number
  // line rectangles (P2 cross-check material)
  _FPDFText_CountRects(textPage: number, startIndex: number, count: number): number
  _FPDFText_GetRect(
    textPage: number,
    rectIndex: number,
    left: number,
    top: number,
    right: number,
    bottom: number,
  ): number
  _FPDFText_GetBoundedText(
    textPage: number,
    left: number,
    top: number,
    right: number,
    bottom: number,
    buffer: number,
    buflen: number,
  ): number
  _FPDFText_GetText(textPage: number, startIndex: number, count: number, result: number): number

  // page objects
  _FPDFPage_CountObjects(page: number): number
  _FPDFPage_GetObject(page: number, index: number): number
  _FPDFPage_RemoveObject(page: number, obj: number): number
  _FPDFPageObj_Destroy(obj: number): void
  _FPDFPageObj_GetType(obj: number): number
  /** float* left, bottom, right, top */
  _FPDFPageObj_GetBounds(obj: number, l: number, b: number, r: number, t: number): number
  /** FS_MATRIX* (6 floats a,b,c,d,e,f) */
  _FPDFPageObj_GetMatrix(obj: number, matrix: number): number
  _FPDFPageObj_SetMatrix(obj: number, matrix: number): number
  /** unsigned int* R, G, B, A */
  _FPDFPageObj_GetFillColor(obj: number, r: number, g: number, b: number, a: number): number
  _FPDFPageObj_GetStrokeColor(obj: number, r: number, g: number, b: number, a: number): number
  /** float* width */
  _FPDFPageObj_GetStrokeWidth(obj: number, width: number): number
  // form XObjects (fpdf_edit.h; optional in older builds — feature-detect)
  _FPDFFormObj_CountObjects?(form: number): number
  _FPDFFormObj_GetObject?(form: number, index: number): number

  // annotations + AcroForm widgets (fpdf_annot.h; optional — feature-detect)
  _FPDFPage_GetAnnotCount?(page: number): number
  _FPDFPage_GetAnnot?(page: number, index: number): number
  _FPDFPage_CloseAnnot?(annot: number): void
  _FPDFAnnot_GetSubtype?(annot: number): number
  /** FS_RECTF* (4 floats: left, top, right, bottom) */
  _FPDFAnnot_GetRect?(annot: number, rect: number): number
  _FPDFAnnot_GetFormFieldType?(form: number, annot: number): number
  _FPDFAnnot_IsChecked?(form: number, annot: number): number
  /** UTF-16LE value of a string/name key ('AS', 'V', …); returns byte length */
  _FPDFAnnot_GetStringValue?(annot: number, key: number, buffer: number, buflen: number): number
  _PDFiumExt_OpenFormFillInfo?(): number
  _PDFiumExt_InitFormFillEnvironment?(doc: number, formInfo: number): number
  _PDFiumExt_ExitFormFillEnvironment?(form: number): void
  _PDFiumExt_CloseFormFillInfo?(formInfo: number): void

  // images
  _FPDFImageObj_GetImageDataDecoded(obj: number, buffer: number, buflen: number): number
  _FPDFImageObj_GetImageDataRaw(obj: number, buffer: number, buflen: number): number
  _FPDFImageObj_GetImageFilterCount(obj: number): number
  _FPDFImageObj_GetImageFilter(obj: number, index: number, buffer: number, buflen: number): number
  /** FPDF_IMAGEOBJ_METADATA* (28 bytes) */
  _FPDFImageObj_GetImageMetadata(obj: number, page: number, metadata: number): number
  /** unsigned int* width, height */
  _FPDFImageObj_GetImagePixelSize(obj: number, width: number, height: number): number
  _FPDFImageObj_GetBitmap(obj: number): number
  _FPDFImageObj_GetRenderedBitmap(doc: number, page: number, obj: number): number

  // bitmaps / page rendering
  _FPDFBitmap_Create(width: number, height: number, alpha: number): number
  _FPDFBitmap_Destroy(bitmap: number): void
  _FPDFBitmap_FillRect(
    bitmap: number,
    left: number,
    top: number,
    width: number,
    height: number,
    color: number,
  ): number
  _FPDFBitmap_GetBuffer(bitmap: number): number
  _FPDFBitmap_GetFormat(bitmap: number): number
  _FPDFBitmap_GetWidth(bitmap: number): number
  _FPDFBitmap_GetHeight(bitmap: number): number
  _FPDFBitmap_GetStride(bitmap: number): number
  _FPDF_RenderPageBitmap(
    bitmap: number,
    page: number,
    startX: number,
    startY: number,
    sizeX: number,
    sizeY: number,
    rotate: number,
    flags: number,
  ): void

  // paths (P2 table-line raw material; declared per plan §3.1)
  _FPDFPath_CountSegments(path: number): number
  _FPDFPath_GetPathSegment(path: number, index: number): number
  _FPDFPathSegment_GetType(segment: number): number
  /** float* x, y */
  _FPDFPathSegment_GetPoint(segment: number, x: number, y: number): number
  _FPDFPathSegment_GetClose(segment: number): number
  /** int* fillmode, FPDF_BOOL* stroke */
  _FPDFPath_GetDrawMode(path: number, fillmode: number, stroke: number): number

  // fonts
  _FPDFTextObj_GetFont(obj: number): number
  _FPDFFont_GetFamilyName(font: number, buffer: number, buflen: number): number
  _FPDFFont_GetBaseFontName(font: number, buffer: number, buflen: number): number
  _FPDFFont_GetWeight(font: number): number
  /** int* angle */
  _FPDFFont_GetItalicAngle(font: number, angle: number): number
  _FPDFFont_GetFlags(font: number): number
  /** float* ascent (fontSize-scaled) */
  _FPDFFont_GetAscent(font: number, fontSize: number, ascent: number): number
  _FPDFFont_GetDescent(font: number, fontSize: number, descent: number): number
  _FPDFFont_GetIsEmbedded(font: number): number
  /** uint8 buffer, size_t* out_buflen */
  _FPDFFont_GetFontData(font: number, buffer: number, buflen: number, outBuflen: number): number

  // structure tree (P1: probe only)
  _FPDF_StructTree_GetForPage(page: number): number
  _FPDF_StructTree_Close(tree: number): void
  _FPDF_StructTree_CountChildren(tree: number): number
  _FPDF_StructTree_GetChildAtIndex(tree: number, index: number): number
}

// page object types (fpdf_edit.h)
export const FPDF_PAGEOBJ_TEXT = 1
export const FPDF_PAGEOBJ_PATH = 2
export const FPDF_PAGEOBJ_IMAGE = 3
export const FPDF_PAGEOBJ_SHADING = 4
export const FPDF_PAGEOBJ_FORM = 5

// fpdf_annot.h subtypes / form-field types
export const FPDF_ANNOT_WIDGET = 20
export const FPDF_FORMFIELD_CHECKBOX = 2
export const FPDF_FORMFIELD_RADIOBUTTON = 3

// path segment types (fpdf_edit.h FPDF_SEGMENT_*)
export const FPDF_SEGMENT_LINETO = 0
export const FPDF_SEGMENT_BEZIERTO = 1
export const FPDF_SEGMENT_MOVETO = 2

// bitmap formats (fpdf_view.h FPDFBitmap_*)
export const BITMAP_FORMAT_GRAY = 1
export const BITMAP_FORMAT_BGR = 2
export const BITMAP_FORMAT_BGRX = 3
export const BITMAP_FORMAT_BGRA = 4

/** scoped malloc: frees even when fn throws */
export function withAlloc<T>(m: PdfiumModule, size: number, fn: (ptr: number) => T): T {
  const ptr = m._malloc(size)
  try {
    return fn(ptr)
  } finally {
    m._free(ptr)
  }
}

// FPDF_ERR_* (fpdf_view.h FPDF_GetLastError)
export const FPDF_ERR_UNKNOWN = 1
export const FPDF_ERR_FILE = 2
export const FPDF_ERR_FORMAT = 3
export const FPDF_ERR_PASSWORD = 4
export const FPDF_ERR_SECURITY = 5

export type PdfLoadErrorCode = 'password-required' | 'corrupt' | 'unsupported'

/**
 * Structured load failure (P22). Extends Error so pre-existing callers that
 * only catch/print keep working; new callers switch on `code`.
 * `password-required` covers both "no password given" and "wrong password".
 */
export class PdfLoadError extends Error {
  readonly code: PdfLoadErrorCode
  /** raw FPDF_GetLastError() value */
  readonly pdfiumError: number

  constructor(code: PdfLoadErrorCode, pdfiumError: number) {
    super(`PDFium could not load the document (${code}, FPDF error ${pdfiumError})`)
    this.name = 'PdfLoadError'
    this.code = code
    this.pdfiumError = pdfiumError
  }
}

function classifyLoadError(err: number): PdfLoadErrorCode {
  if (err === FPDF_ERR_PASSWORD) return 'password-required'
  if (err === FPDF_ERR_FILE || err === FPDF_ERR_FORMAT || err === FPDF_ERR_UNKNOWN) return 'corrupt'
  return 'unsupported' // FPDF_ERR_SECURITY (unsupported handler) and anything new
}

/**
 * Load a PDF from bytes, run fn, always close + free. PDFium itself retries
 * an owner-locked file with the empty user password when `password` is
 * omitted, so those convert normally; a real user password lands here as
 * FPDF_ERR_PASSWORD → PdfLoadError('password-required').
 */
export function withPdfDocument<T>(
  m: PdfiumModule,
  bytes: Uint8Array,
  fn: (doc: number) => T,
  password?: string,
): T {
  const ptr = m._malloc(bytes.length)
  m.HEAPU8.set(bytes, ptr)
  let pwPtr = 0
  if (password !== undefined) {
    const pw = new TextEncoder().encode(password + '\0')
    pwPtr = m._malloc(pw.length)
    m.HEAPU8.set(pw, pwPtr)
  }
  const doc = m._FPDF_LoadMemDocument(ptr, bytes.length, pwPtr)
  if (!doc) {
    if (pwPtr) m._free(pwPtr)
    m._free(ptr)
    const err = m._FPDF_GetLastError()
    throw new PdfLoadError(classifyLoadError(err), err)
  }
  try {
    return fn(doc)
  } finally {
    m._FPDF_CloseDocument(doc)
    if (pwPtr) m._free(pwPtr)
    m._free(ptr)
  }
}
