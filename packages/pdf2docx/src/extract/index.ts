/**
 * Extraction layer — the ONLY layer that talks to PDFium. Produces plain data
 * (chars, images, page geometry, quality flags) for the pure-geometry analysis
 * layer. No filesystem access; the initialized wasm module comes from the caller.
 */
import { detectVectorRegions } from '../analyze/vector'
import type { Rect } from '../geometry'
import { intersectArea, overlapRatio, rectArea } from '../geometry'
import type { PdfChar, PageRender, RawPath, RawSubpath } from '../ir'
import { scriptOf } from '../script'
import type { PdfiumModule } from './pdfium'
import {
  BITMAP_FORMAT_BGR,
  BITMAP_FORMAT_BGRA,
  BITMAP_FORMAT_BGRX,
  BITMAP_FORMAT_GRAY,
  FPDF_ANNOT_WIDGET,
  FPDF_FORMFIELD_CHECKBOX,
  FPDF_FORMFIELD_RADIOBUTTON,
  FPDF_PAGEOBJ_FORM,
  FPDF_PAGEOBJ_IMAGE,
  FPDF_PAGEOBJ_PATH,
  FPDF_PAGEOBJ_SHADING,
  FPDF_PAGEOBJ_TEXT,
  FPDF_SEGMENT_BEZIERTO,
  FPDF_SEGMENT_MOVETO,
  withAlloc,
} from './pdfium'
import { familyFromPsName, italicFromPsName, SUBSET_PREFIX, weightFromPsName } from './fontname'
import type { PsNameWeight } from './fontname'
import { encodeRgbaPng } from './png'

export { withPdfDocument, withAlloc, PdfLoadError } from './pdfium'
export type { PdfiumModule, PdfLoadErrorCode } from './pdfium'
export { encodeRgbaPng } from './png'
export { familyFromPsName } from './fontname'

export interface ExtractedImage {
  /** placement box on the page (PDF points) */
  box: Rect
  data: Uint8Array
  mime: 'image/png' | 'image/jpeg'
  pixelWidth: number
  pixelHeight: number
  /** source paint order (top-level page-object index, P16 A) */
  z?: number
}

export interface ExtractedPage {
  index: number
  widthPt: number
  heightPt: number
  rotation: number
  chars: PdfChar[]
  images: ExtractedImage[]
  /** raw page path objects (table-line / shading raw material) */
  paths: RawPath[]
  degraded: boolean
  degradedReason?: string
  scanned: boolean
  hasStructTree: boolean
  render?: PageRender
  /** vector-illustration regions rasterized into `images` (their chars/paths are dropped) */
  vectorRegions: Rect[]
  /** searchable scan: the invisible OCR layer was flipped visible and the scan bitmap dropped (P27) */
  ocrTextRecovered?: boolean
  /** graphics-dominant searchable scan: exported as its full-page render (P29 B) */
  ocrImageKept?: boolean
  /** extraction lost ALL visible body text with nothing carrying it (P27 guard input) */
  textLost?: boolean
  /** share of bad (U+FFFD / private-use) code points among text chars (P4 confidence input) */
  badUnicodeRatio: number
  /**
   * page-covering background stack (gradient shading / raster wallpaper)
   * rendered to a bitmap (P9 B) — the rebuild layer pins it behind the text
   * as a full-page behindDoc float. Absent on plain (white / flat-fill) pages.
   */
  bgRender?: PageRender
  /**
   * non-page-covering gradient shadings rasterized transparent (P19): slide
   * title bars / accent strips. Consumed ONLY by the canvas rebuild path —
   * flow pages drop them exactly as before.
   */
  decorImages?: ExtractedImage[]
}

export interface ExtractOptions {
  /** raster scale for full-page fallback renders (pixels per PDF point), default 2 */
  renderScale?: number
}

/** ToUnicode quality gate: this share of bad code points marks the page degraded */
const BAD_UNICODE_RATIO = 0.15
/** minimum sample size before the bad-unicode ratio is trusted */
const BAD_UNICODE_MIN_CHARS = 10
/** |angle| above this (radians, ~15°) counts a char as rotated/vertical */
const ANGLED_CHAR_RAD = 0.26
/** share of rotated chars that triggers the vertical-text fallback */
const ANGLED_RATIO = 0.3
/** image covering this share of the page + no text ⇒ scanned page */
const SCANNED_IMAGE_COVER = 0.8
/** "no text layer" threshold for scanned detection */
const SCANNED_MAX_CHARS = 2
/** content is on-page when it reaches within this many points of the page box (P14) */
const PAGE_CLIP_TOL_PT = 0.01

const utf8Decoder = new TextDecoder()

function toHexColor(r: number, g: number, b: number): string {
  const hex = (v: number) => v.toString(16).padStart(2, '0').toUpperCase()
  return hex(r) + hex(g) + hex(b)
}

const isWhitespaceCode = (code: number): boolean =>
  code === 0x20 || code === 0xa0 || code === 0x0d || code === 0x0a || code === 0x09

const isBadUnicode = (code: number): boolean =>
  code === 0xfffd || code === 0xfffe || (code >= 0xe000 && code <= 0xf8ff) || isControlCode(code)

/** C0/C1 control code (not tab/newline/CR): never real text — a broken
 * ToUnicode map voicing raw glyph ids (mojibake headings render as tofu) */
const isControlCode = (code: number): boolean =>
  (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) || (code >= 0x7f && code <= 0x9f)

interface FontInfo {
  family: string
  italic: boolean
  /** explicit weight declared by the PS name's style tokens (the family name
   * no longer carries them); overrides the descriptor's StemV-derived weight */
  nameWeight: PsNameWeight
}

/** PDF FontDescriptor flags bit 7 (value 64) = italic */
const FONT_FLAG_ITALIC = 1 << 6

/** the embedded font's real family name (FPDFFont_GetFamilyName), '' when absent */
function readFontFamilyName(m: PdfiumModule, textPage: number, index: number): string {
  const obj = m._FPDFText_GetTextObject(textPage, index)
  if (!obj) return ''
  const font = m._FPDFTextObj_GetFont(obj)
  if (!font) return ''
  const needed = m._FPDFFont_GetFamilyName(font, 0, 0)
  if (needed <= 1) return ''
  return withAlloc(m, needed, (buf) => {
    m._FPDFFont_GetFamilyName(font, buf, needed)
    return utf8Decoder
      .decode(m.HEAPU8.subarray(buf, buf + needed - 1))
      .replace(SUBSET_PREFIX, '')
      .trim()
  })
}

function readCharFontInfo(m: PdfiumModule, textPage: number, index: number): FontInfo | null {
  return withAlloc(m, 4, (flagsPtr) => {
    const needed = m._FPDFText_GetFontInfo(textPage, index, 0, 0, flagsPtr)
    if (needed <= 0) return null
    return withAlloc(m, needed, (buf) => {
      m._FPDFText_GetFontInfo(textPage, index, buf, needed, flagsPtr)
      // buffer is UTF-8 with a trailing NUL; usually the PostScript name
      const psName = utf8Decoder
        .decode(m.HEAPU8.subarray(buf, buf + needed - 1))
        .replace(SUBSET_PREFIX, '')
      const flags = m.HEAPU32[flagsPtr >> 2]! | 0
      const italic =
        (flags & FONT_FLAG_ITALIC) !== 0 ||
        /italic|oblique/i.test(psName) ||
        italicFromPsName(psName)
      const nameWeight = weightFromPsName(psName)
      // rFonts needs the FAMILY name — a PS name makes Word substitute fonts.
      // Subset fonts often carry the PS name even in their name table, so the
      // heuristic runs over GetFamilyName's answer too (space-containing real
      // family names pass through untouched).
      const family = familyFromPsName(readFontFamilyName(m, textPage, index) || psName)
      return { family, italic, nameWeight }
    })
  })
}

// ── invisible-glyph verification (P20) ──
// Word's hidden formatting marks (section/page-break labels) can arrive as Type 3 glyphs
// with broken bounding boxes: the text layer extracts them but no renderer
// paints a single pixel. Chars with an EMPTY font family are the suspects;
// a tiny region render confirms — no dark ink where a dark glyph should be
// means the char is invisible in the source and rides to the docx as
// w:vanish. Direction is conservative: any ink found keeps the char visible.

/** per-page cap on verification renders (cost guard) */
const INKLESS_MAX_CLUSTERS = 12
/** suspects darker than this luminance are verifiable against "no dark ink" */
const INKLESS_MAX_LUMINANCE = 0.5
/** only chars set well below the page's median size are suspects — hidden
 * formatting marks are tiny (6.5pt vs 10.5pt body); body text never qualifies
 * and footnote refs merely cost one verification render each */
const INKLESS_MAX_SIZE_RATIO = 0.7
/** a rendered pixel darker than this channel value counts as ink */
const INK_CHANNEL_MAX = 120
const INKLESS_RENDER_SCALE = 3

function luminanceOfHex(hex: string): number {
  const v = parseInt(hex, 16)
  if (Number.isNaN(v)) return 0
  const r = (v >> 16) & 0xff
  const g = (v >> 8) & 0xff
  const b = v & 0xff
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255
}

function markInklessChars(m: PdfiumModule, page: number, chars: PdfChar[]): void {
  const sizes = chars
    .filter((c) => c.code > 0x20 && c.fontSize > 0)
    .map((c) => c.fontSize)
    .sort((a, b) => a - b)
  if (sizes.length === 0) return
  const medianSize = sizes[Math.floor(sizes.length / 2)]!
  const suspects = chars.filter(
    (c) =>
      c.code > 0x20 &&
      c.fontSize > 0 &&
      c.fontSize <= INKLESS_MAX_SIZE_RATIO * medianSize &&
      luminanceOfHex(c.color) <= INKLESS_MAX_LUMINANCE &&
      c.box.x1 > c.box.x0 &&
      c.box.y1 > c.box.y0,
  )
  if (suspects.length === 0) return
  // cluster line-wise: same baseline band, no wild x jumps
  suspects.sort((a, b) => b.originY - a.originY || a.box.x0 - b.box.x0)
  const clusters: PdfChar[][] = []
  for (const c of suspects) {
    const last = clusters[clusters.length - 1]
    const tail = last?.[last.length - 1]
    if (
      tail !== undefined &&
      Math.abs(tail.originY - c.originY) <= 2 &&
      c.box.x0 - tail.box.x1 <= 24
    ) {
      last!.push(c)
    } else {
      clusters.push([c])
    }
  }
  for (const cluster of clusters.slice(0, INKLESS_MAX_CLUSTERS)) {
    const region: Rect = {
      x0: Math.min(...cluster.map((c) => c.box.x0)) - 1,
      x1: Math.max(...cluster.map((c) => c.box.x1)) + 1,
      y0: Math.min(...cluster.map((c) => c.box.y0)) - 1,
      y1: Math.max(...cluster.map((c) => c.box.y1)) + 1,
    }
    if (regionHasDarkInk(m, page, region)) continue
    for (const c of cluster) c.invisible = true
  }
}

/** render the region and report whether any pixel reads as dark ink */
function regionHasDarkInk(m: PdfiumModule, page: number, region: Rect): boolean {
  const pageWidthPt = m._FPDF_GetPageWidthF(page)
  const pageHeightPt = m._FPDF_GetPageHeightF(page)
  const width = Math.max(1, Math.round((region.x1 - region.x0) * INKLESS_RENDER_SCALE))
  const height = Math.max(1, Math.round((region.y1 - region.y0) * INKLESS_RENDER_SCALE))
  const bitmap = m._FPDFBitmap_Create(width, height, 0)
  if (!bitmap) return true // cannot verify → keep visible
  try {
    m._FPDFBitmap_FillRect(bitmap, 0, 0, width, height, 0xffffffff)
    m._FPDF_RenderPageBitmap(
      bitmap,
      page,
      Math.round(-region.x0 * INKLESS_RENDER_SCALE),
      Math.round(-(pageHeightPt - region.y1) * INKLESS_RENDER_SCALE),
      Math.round(pageWidthPt * INKLESS_RENDER_SCALE),
      Math.round(pageHeightPt * INKLESS_RENDER_SCALE),
      0,
      0,
    )
    const px = bitmapToRgba(m, bitmap)
    if (!px) return true
    for (let i = 0; i < px.rgba.length; i += 4) {
      if (
        px.rgba[i]! <= INK_CHANNEL_MAX &&
        px.rgba[i + 1]! <= INK_CHANNEL_MAX &&
        px.rgba[i + 2]! <= INK_CHANNEL_MAX
      )
        return true
    }
    return false
  } finally {
    m._FPDFBitmap_Destroy(bitmap)
  }
}

function readChars(m: PdfiumModule, textPage: number): PdfChar[] {
  const count = m._FPDFText_CountChars(textPage)
  const chars: PdfChar[] = []
  if (count <= 0) return chars

  // 8-byte-aligned scratch: 4 doubles for the char box, 2 for the origin,
  // 4 uints for the color, 4 floats for the loose box, 6 for the matrix
  const d4 = m._malloc(4 * 8)
  const d2 = m._malloc(2 * 8)
  const u4 = m._malloc(4 * 4)
  const f4 = m._malloc(4 * 4)
  const f6 = m._malloc(6 * 4)
  // font info is per text object; chars of one object share it
  const fontCache = new Map<number, FontInfo | null>()
  // render mode is per text object too (invisible text = PDF Tr 3/7, the way
  // Word exports hidden formatting marks like section-break labels)
  const renderModeCache = new Map<number, number>()
  let lastFont: FontInfo = { family: '', italic: false, nameWeight: null }
  try {
    for (let i = 0; i < count; i++) {
      const code = m._FPDFText_GetUnicode(textPage, i)
      if (!m._FPDFText_GetCharBox(textPage, i, d4, d4 + 8, d4 + 16, d4 + 24)) continue
      const box: Rect = {
        x0: m.HEAPF64[d4 >> 3]!,
        x1: m.HEAPF64[(d4 >> 3) + 1]!,
        y0: m.HEAPF64[(d4 >> 3) + 2]!,
        y1: m.HEAPF64[(d4 >> 3) + 3]!,
      }
      // FS_RECTF: left, top, right, bottom
      let looseBox = box
      if (m._FPDFText_GetLooseCharBox(textPage, i, f4)) {
        looseBox = {
          x0: m.HEAPF32[f4 >> 2]!,
          y1: m.HEAPF32[(f4 >> 2) + 1]!,
          x1: m.HEAPF32[(f4 >> 2) + 2]!,
          y0: m.HEAPF32[(f4 >> 2) + 3]!,
        }
      }
      let originX = box.x0
      let originY = box.y0
      if (m._FPDFText_GetCharOrigin(textPage, i, d2, d2 + 8)) {
        originX = m.HEAPF64[d2 >> 3]!
        originY = m.HEAPF64[(d2 >> 3) + 1]!
      }
      const obj = m._FPDFText_GetTextObject(textPage, i)
      let font = fontCache.get(obj)
      if (font === undefined) {
        font = readCharFontInfo(m, textPage, i)
        fontCache.set(obj, font)
      }
      if (font) lastFont = font
      let renderMode = renderModeCache.get(obj)
      if (renderMode === undefined) {
        renderMode = obj ? m._FPDFTextObj_GetTextRenderMode(obj) : -1
        renderModeCache.set(obj, renderMode)
      }
      // FPDF_TEXTRENDERMODE_INVISIBLE = 3; mode 7 adds a clip path on top
      const invisible = renderMode === 3 || renderMode === 7

      let color = '000000'
      if (m._FPDFText_GetFillColor(textPage, i, u4, u4 + 4, u4 + 8, u4 + 12)) {
        color = toHexColor(
          m.HEAPU32[u4 >> 2]!,
          m.HEAPU32[(u4 >> 2) + 1]!,
          m.HEAPU32[(u4 >> 2) + 2]!,
        )
      }
      // horizontal squeeze (PDF Tz / docx w:w) shows up as matrix a ≠ d;
      // the matrix also scales the nominal font size into page space (Chromium
      // print PDFs use size 16 with a 0.75 matrix for an effective 12pt)
      let hscale = 1
      let vscale = 1
      if (m._FPDFText_GetMatrix(textPage, i, f6)) {
        const base = f6 >> 2
        const sx = Math.hypot(m.HEAPF32[base]!, m.HEAPF32[base + 1]!)
        const sy = Math.hypot(m.HEAPF32[base + 2]!, m.HEAPF32[base + 3]!)
        if (sx > 0 && sy > 0) {
          hscale = sx / sy
          vscale = sy
        }
      }

      const weight = m._FPDFText_GetFontWeight(textPage, i)
      const info = font ?? lastFont
      chars.push({
        code,
        text: code > 0 ? String.fromCodePoint(code) : '',
        box,
        looseBox,
        originX,
        originY,
        angle: m._FPDFText_GetCharAngle(textPage, i),
        hscale,
        fontSize: m._FPDFText_GetFontSize(textPage, i) * vscale,
        // an explicit PS-name style outranks the descriptor weight in BOTH
        // directions: PDFium derives that weight from /StemV, and producers
        // write StemV values that read as bold on declared-Roman faces (P21 A)
        fontWeight:
          info.nameWeight === 'bold'
            ? Math.max(weight, 700)
            : info.nameWeight === 'regular'
              ? 400
              : weight > 0
                ? weight
                : 400,
        fontFamily: info.family,
        italic: info.italic,
        color,
        isGenerated: m._FPDFText_IsGenerated(textPage, i) === 1,
        ...(invisible ? { invisible: true } : {}),
        isHyphen: m._FPDFText_IsHyphen(textPage, i) === 1,
        script: scriptOf(code),
        pdfCharIndex: i,
      })
    }
  } finally {
    for (const p of [d4, d2, u4, f4, f6]) m._free(p)
  }
  return chars
}

/** Copy a PDFium bitmap out of the heap as tightly-packed RGBA. */
function bitmapToRgba(
  m: PdfiumModule,
  bitmap: number,
): { rgba: Uint8Array; width: number; height: number } | null {
  const width = m._FPDFBitmap_GetWidth(bitmap)
  const height = m._FPDFBitmap_GetHeight(bitmap)
  const stride = m._FPDFBitmap_GetStride(bitmap)
  const format = m._FPDFBitmap_GetFormat(bitmap)
  const buffer = m._FPDFBitmap_GetBuffer(bitmap)
  if (width <= 0 || height <= 0 || !buffer) return null

  const rgba = new Uint8Array(width * height * 4)
  const heap = m.HEAPU8
  for (let y = 0; y < height; y++) {
    const rowIn = buffer + y * stride
    const rowOut = y * width * 4
    if (format === BITMAP_FORMAT_BGRA || format === BITMAP_FORMAT_BGRX) {
      for (let x = 0; x < width; x++) {
        const src = rowIn + x * 4
        const dst = rowOut + x * 4
        rgba[dst] = heap[src + 2]!
        rgba[dst + 1] = heap[src + 1]!
        rgba[dst + 2] = heap[src]!
        rgba[dst + 3] = format === BITMAP_FORMAT_BGRA ? heap[src + 3]! : 255
      }
    } else if (format === BITMAP_FORMAT_BGR) {
      for (let x = 0; x < width; x++) {
        const src = rowIn + x * 3
        const dst = rowOut + x * 4
        rgba[dst] = heap[src + 2]!
        rgba[dst + 1] = heap[src + 1]!
        rgba[dst + 2] = heap[src]!
        rgba[dst + 3] = 255
      }
    } else if (format === BITMAP_FORMAT_GRAY) {
      for (let x = 0; x < width; x++) {
        const v = heap[rowIn + x]!
        const dst = rowOut + x * 4
        rgba[dst] = v
        rgba[dst + 1] = v
        rgba[dst + 2] = v
        rgba[dst + 3] = 255
      }
    } else {
      return null
    }
  }
  return { rgba, width, height }
}

function readImages(
  m: PdfiumModule,
  doc: number,
  page: number,
  /** objects below this index are baked into the background stack (P9 B/P16 B) */
  skipBelowIndex = 0,
  /** pre-extraction drop test — skipping here saves the decode itself (P28):
   * a searchable scan's page-covering tiles would only be filtered out again */
  skipBox?: (box: Rect) => boolean,
): ExtractedImage[] {
  const images: ExtractedImage[] = []
  const count = m._FPDFPage_CountObjects(page)
  const f6 = m._malloc(6 * 4)
  const readMatrix = (obj: number): Matrix => {
    if (!m._FPDFPageObj_GetMatrix(obj, f6)) return IDENTITY
    const base = f6 >> 2
    return [
      m.HEAPF32[base]!,
      m.HEAPF32[base + 1]!,
      m.HEAPF32[base + 2]!,
      m.HEAPF32[base + 3]!,
      m.HEAPF32[base + 4]!,
      m.HEAPF32[base + 5]!,
    ]
  }
  // slide/report exporters wrap page art in form XObjects (P29 E) — a
  // top-level-only walk ships those pages imageless. Nested boxes come from
  // the composed matrix over the image's unit square.
  const visit = (obj: number, parent: Matrix | null, depth: number, z: number): void => {
    const type = m._FPDFPageObj_GetType(obj)
    if (type === FPDF_PAGEOBJ_FORM) {
      if (
        depth >= FORM_RECURSION_MAX ||
        typeof m._FPDFFormObj_CountObjects !== 'function' ||
        typeof m._FPDFFormObj_GetObject !== 'function'
      ) {
        return
      }
      const formMatrix = composeMatrix(parent ?? IDENTITY, readMatrix(obj))
      const children = m._FPDFFormObj_CountObjects(obj)
      for (let c = 0; c < children; c++) {
        const child = m._FPDFFormObj_GetObject!(obj, c)
        if (child) visit(child, formMatrix, depth + 1, z)
      }
      return
    }
    if (type !== FPDF_PAGEOBJ_IMAGE) return
    let box: Rect
    if (parent === null) {
      if (!m._FPDFPageObj_GetBounds(obj, f6, f6 + 4, f6 + 8, f6 + 12)) return
      box = {
        x0: m.HEAPF32[f6 >> 2]!,
        y0: m.HEAPF32[(f6 >> 2) + 1]!,
        x1: m.HEAPF32[(f6 >> 2) + 2]!,
        y1: m.HEAPF32[(f6 >> 2) + 3]!,
      }
    } else {
      const [a, b, c, d, e, f] = composeMatrix(parent, readMatrix(obj))
      const xs = [e, a + e, c + e, a + c + e]
      const ys = [f, b + f, d + f, b + d + f]
      box = {
        x0: Math.min(...xs),
        y0: Math.min(...ys),
        x1: Math.max(...xs),
        y1: Math.max(...ys),
      }
    }
    if (rectArea(box) <= 0) return
    if (skipBox !== undefined && skipBox(box)) return
    const image = extractImagePayload(m, doc, page, obj, box)
    if (image) images.push({ ...image, z })
  }
  try {
    for (let i = skipBelowIndex; i < count; i++) {
      const obj = m._FPDFPage_GetObject(page, i)
      if (obj) visit(obj, null, 0, i)
    }
  } finally {
    m._free(f6)
  }
  return images
}

function imageFilters(m: PdfiumModule, obj: number): string[] {
  const count = m._FPDFImageObj_GetImageFilterCount(obj)
  const filters: string[] = []
  for (let i = 0; i < count; i++) {
    const needed = m._FPDFImageObj_GetImageFilter(obj, i, 0, 0)
    if (needed <= 0) continue
    withAlloc(m, needed, (buf) => {
      m._FPDFImageObj_GetImageFilter(obj, i, buf, needed)
      filters.push(utf8Decoder.decode(m.HEAPU8.subarray(buf, buf + needed - 1)))
    })
  }
  return filters
}

/** alpha below this marks a pixel as meaningfully transparent (rounding-safe) */
const ALPHA_OPAQUE_MIN = 250

function rgbaHasTransparency(rgba: Uint8Array): boolean {
  for (let i = 3; i < rgba.length; i += 4) {
    if (rgba[i]! < ALPHA_OPAQUE_MIN) return true
  }
  return false
}

/** re-rendered transparent images cap at this size per dimension (px) */
const RERENDER_MAX_PX = 4096
/** natural resolution must beat the device-size render by this much to re-render */
const RERENDER_MIN_GAIN = 1.3
/**
 * total-pixel ceiling for one image render: GetRenderedBitmap follows the
 * object matrix, and a scan tile placed with a huge matrix asks the wasm heap
 * for gigabytes (hard 2GiB limit) — the alloc fails and the image silently
 * vanishes. Beyond the cap the matrix is pre-scaled DOWN instead (a
 * downsampled embed always beats a dropped one).
 */
const RENDER_MAX_TOTAL_PX = 24_000_000

/**
 * Render the image object through PDFium (transforms + soft masks applied).
 * GetRenderedBitmap sizes its output from the object matrix (device size); for
 * a transparent image that PNG is the ONLY payload we can emit, so when the
 * natural pixel size is clearly larger the matrix is temporarily scaled up to
 * render near natural resolution, then restored.
 */
function renderImageObjRgba(
  m: PdfiumModule,
  doc: number,
  page: number,
  obj: number,
  scale: number,
): { rgba: Uint8Array; width: number; height: number } | null {
  // clamp the requested scale so the output stays under the total-pixel cap
  const extents = withAlloc(m, 6 * 4, (f6) => {
    if (!m._FPDFPageObj_GetMatrix(obj, f6)) return null
    const v = m.HEAPF32.subarray(f6 >> 2, (f6 >> 2) + 6)
    return { w: Math.hypot(v[0]!, v[1]!), h: Math.hypot(v[2]!, v[3]!) }
  })
  if (extents !== null) {
    const outPx = extents.w * extents.h * scale * scale
    if (outPx > RENDER_MAX_TOTAL_PX) {
      scale = Math.sqrt(RENDER_MAX_TOTAL_PX / Math.max(1, extents.w * extents.h))
    }
  }
  const restore =
    scale !== 1
      ? withAlloc(m, 6 * 4, (f6) => {
          if (!m._FPDFPageObj_GetMatrix(obj, f6)) return null
          const base = f6 >> 2
          const orig = Array.from(m.HEAPF32.subarray(base, base + 6)) as number[]
          // scale extents only (a,b,c,d) — the bitmap size follows the
          // transformed bbox extents, translation does not matter
          for (let i = 0; i < 4; i++) m.HEAPF32[base + i] = orig[i]! * scale
          m._FPDFPageObj_SetMatrix(obj, f6)
          return orig
        })
      : null
  try {
    const bitmap = m._FPDFImageObj_GetRenderedBitmap(doc, page, obj)
    if (!bitmap) return null
    try {
      return bitmapToRgba(m, bitmap)
    } finally {
      m._FPDFBitmap_Destroy(bitmap)
    }
  } finally {
    if (restore) {
      withAlloc(m, 6 * 4, (f6) => {
        const base = f6 >> 2
        for (let i = 0; i < 6; i++) m.HEAPF32[base + i] = restore[i]!
        m._FPDFPageObj_SetMatrix(obj, f6)
      })
    }
  }
}

/**
 * Component count from a JPEG's SOF segment, or 0 when unparseable. Word
 * renders 4-component (CMYK/YCCK) JPEGs as black/inverted slabs, so only
 * 1- and 3-component streams may embed raw — others re-encode via the
 * rendered RGBA (P29 A).
 */
export function jpegSofComponents(data: Uint8Array): number {
  if (data.length < 4 || data[0] !== 0xff || data[1] !== 0xd8) return 0
  let i = 2
  while (i + 3 < data.length) {
    if (data[i] !== 0xff) return 0
    const marker = data[i + 1]!
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7)) {
      i += 2
      continue
    }
    if (marker === 0xda || marker === 0xd9) return 0 // hit scan data without a SOF
    const len = (data[i + 2]! << 8) | data[i + 3]!
    if (len < 2) return 0
    const isSof =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc
    if (isSof) return i + 9 < data.length ? data[i + 9]! : 0
    i += 2 + len
  }
  return 0
}

/** raw JPEG passthrough for a single-DCT image object; null = not eligible */
function tryRawJpeg(
  m: PdfiumModule,
  obj: number,
  box: Rect,
  naturalW: number,
  naturalH: number,
): ExtractedImage | null {
  const size = m._FPDFImageObj_GetImageDataRaw(obj, 0, 0)
  if (size <= 0 || naturalW <= 0 || naturalH <= 0) return null
  const data = withAlloc(m, size, (buf) => {
    m._FPDFImageObj_GetImageDataRaw(obj, buf, size)
    return Uint8Array.from(m.HEAPU8.subarray(buf, buf + size))
  })
  const components = jpegSofComponents(data)
  if (components !== 1 && components !== 3) return null
  return { box, data, mime: 'image/jpeg', pixelWidth: naturalW, pixelHeight: naturalH }
}

function extractImagePayload(
  m: PdfiumModule,
  doc: number,
  page: number,
  obj: number,
  box: Rect,
): ExtractedImage | null {
  const [naturalW, naturalH] = withAlloc(m, 8, (p) => {
    return m._FPDFImageObj_GetImagePixelSize(obj, p, p + 4)
      ? [m.HEAPU32[p >> 2]!, m.HEAPU32[(p >> 2) + 1]!]
      : [0, 0]
  })

  // Render first (transforms + soft masks applied) — the alpha channel decides
  // the payload. A masked image embedded as its bare base stream shows the
  // matte (usually black) where the PDF was transparent (P9 A).
  let px = renderImageObjRgba(m, doc, page, obj, 1)
  const transparent = px !== null && rgbaHasTransparency(px.rgba)

  // Plain DCT-encoded OPAQUE images: the raw stream already IS a JPEG file —
  // embed as-is at natural resolution. (Multi-filter chains like Flate+DCT are
  // not raw JPEG, and transparent JPEGs must carry their mask via PNG.)
  const filters = imageFilters(m, obj)
  if (!transparent && filters.length === 1 && filters[0] === 'DCTDecode') {
    const raw = tryRawJpeg(m, obj, box, naturalW, naturalH)
    if (raw) return raw
  }

  if (!px) return null
  // PNG-payload image rendered well below its natural resolution → re-render
  // scaled up so the PNG keeps the source detail (the raw passthrough above
  // already carries full resolution for eligible JPEGs)
  if (naturalW > px.width * RERENDER_MIN_GAIN && naturalH > 0) {
    const scale = Math.min(
      naturalW / Math.max(1, px.width),
      RERENDER_MAX_PX / Math.max(1, px.width),
      RERENDER_MAX_PX / Math.max(1, px.height),
    )
    if (scale > 1) px = renderImageObjRgba(m, doc, page, obj, scale) ?? px
  }
  return {
    box,
    data: encodeRgbaPng(px.rgba, px.width, px.height),
    mime: 'image/png',
    pixelWidth: px.width,
    pixelHeight: px.height,
  }
}

/** FS_MATRIX as a tuple: page = (a·x + c·y + e, b·x + d·y + f) */
type Matrix = [number, number, number, number, number, number]
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0]

/** outer ∘ inner: apply `inner` first, then `outer` */
function composeMatrix(outer: Matrix, inner: Matrix): Matrix {
  const [oa, ob, oc, od, oe, of_] = outer
  const [ia, ib, ic, id, ie, if_] = inner
  return [
    oa * ia + oc * ib,
    ob * ia + od * ib,
    oa * ic + oc * id,
    ob * ic + od * id,
    oa * ie + oc * if_ + oe,
    ob * ie + od * if_ + of_,
  ]
}

/** nested form XObjects deeper than this are ignored (malformed/pathological) */
const FORM_RECURSION_MAX = 4

/**
 * Read page path objects as raw polylines in page space (object matrix
 * applied). Bezier subpaths are kept but flagged; the analysis layer's shape
 * normalization decides what becomes a Stroke/Fill and what is ignored.
 * Form XObjects are entered recursively (P14 C): PowerPoint's PDF export
 * wraps divider lines in forms, and a top-level-only walk never sees them.
 */
function readPaths(m: PdfiumModule, page: number, skipBelowIndex = 0): RawPath[] {
  const paths: RawPath[] = []
  const count = m._FPDFPage_CountObjects(page)
  // scratch: 6 matrix floats / 2 point floats / 4 color uints / 2 drawmode ints share buffers
  const f6 = m._malloc(6 * 4)
  const u4 = m._malloc(4 * 4)

  const readMatrix = (obj: number): Matrix => {
    if (!m._FPDFPageObj_GetMatrix(obj, f6)) return IDENTITY
    const base = f6 >> 2
    return [
      m.HEAPF32[base]!,
      m.HEAPF32[base + 1]!,
      m.HEAPF32[base + 2]!,
      m.HEAPF32[base + 3]!,
      m.HEAPF32[base + 4]!,
      m.HEAPF32[base + 5]!,
    ]
  }

  const visit = (obj: number, parent: Matrix, depth: number, z: number): void => {
    const type = m._FPDFPageObj_GetType(obj)
    if (type === FPDF_PAGEOBJ_FORM) {
      if (
        depth >= FORM_RECURSION_MAX ||
        typeof m._FPDFFormObj_CountObjects !== 'function' ||
        typeof m._FPDFFormObj_GetObject !== 'function'
      ) {
        return
      }
      // children live in the form's own space; its matrix maps into `parent`
      const formMatrix = composeMatrix(parent, readMatrix(obj))
      const children = m._FPDFFormObj_CountObjects(obj)
      for (let i = 0; i < children; i++) {
        const child = m._FPDFFormObj_GetObject!(obj, i)
        if (child) visit(child, formMatrix, depth + 1, z)
      }
      return
    }
    if (type !== FPDF_PAGEOBJ_PATH) return

    let filled = false
    let stroked = false
    if (m._FPDFPath_GetDrawMode(obj, u4, u4 + 4)) {
      filled = (m.HEAPU32[u4 >> 2]! | 0) !== 0
      stroked = m.HEAPU32[(u4 >> 2) + 1]! !== 0
    }
    if (!filled && !stroked) return // invisible (clip-only) path

    const [ma, mb, mc, md, me, mf] = composeMatrix(parent, readMatrix(obj))

    const readColor = (fn: (o: number, r: number, g: number, b: number, a: number) => number) => {
      if (!fn(obj, u4, u4 + 4, u4 + 8, u4 + 12)) return '000000'
      return toHexColor(m.HEAPU32[u4 >> 2]!, m.HEAPU32[(u4 >> 2) + 1]!, m.HEAPU32[(u4 >> 2) + 2]!)
    }
    let fillAlpha = 255
    const fillColor = (() => {
      if (!m._FPDFPageObj_GetFillColor(obj, u4, u4 + 4, u4 + 8, u4 + 12)) return '000000'
      fillAlpha = m.HEAPU32[(u4 >> 2) + 3]! // P10 C: translucent fills are glows, not shading
      return toHexColor(m.HEAPU32[u4 >> 2]!, m.HEAPU32[(u4 >> 2) + 1]!, m.HEAPU32[(u4 >> 2) + 2]!)
    })()
    const strokeColor = readColor(m._FPDFPageObj_GetStrokeColor.bind(m))
    let strokeWidth = 1
    if (m._FPDFPageObj_GetStrokeWidth(obj, f6)) strokeWidth = m.HEAPF32[f6 >> 2]!
    // scale the stroke width along with the object transform
    strokeWidth *= Math.sqrt(Math.abs(ma * md - mb * mc)) || 1

    const subpaths: RawSubpath[] = []
    let current: RawSubpath | null = null
    const segCount = m._FPDFPath_CountSegments(obj)
    for (let s = 0; s < segCount; s++) {
      const seg = m._FPDFPath_GetPathSegment(obj, s)
      if (!seg) continue
      const type = m._FPDFPathSegment_GetType(seg)
      if (!m._FPDFPathSegment_GetPoint(seg, f6, f6 + 4)) continue
      const px = m.HEAPF32[f6 >> 2]!
      const py = m.HEAPF32[(f6 >> 2) + 1]!
      const point = { x: ma * px + mc * py + me, y: mb * px + md * py + mf }
      if (type === FPDF_SEGMENT_MOVETO || !current) {
        current = { points: [point], closed: false, hasCurves: false }
        subpaths.push(current)
      } else {
        current.points.push(point)
        if (type === FPDF_SEGMENT_BEZIERTO) current.hasCurves = true
      }
      if (m._FPDFPathSegment_GetClose(seg)) current.closed = true
    }
    if (subpaths.length === 0) return
    paths.push({
      subpaths,
      filled,
      stroked,
      fillColor,
      strokeColor,
      strokeWidth,
      fillAlpha,
      z,
      ...(depth > 0 ? { fromForm: true } : {}),
    })
  }

  try {
    for (let i = skipBelowIndex; i < count; i++) {
      const obj = m._FPDFPage_GetObject(page, i)
      if (obj) visit(obj, IDENTITY, 0, i)
    }
  } finally {
    m._free(f6)
    m._free(u4)
  }
  return paths
}

// ── searchable-scan OCR overlay recovery (P27) ──
// OCR tools stamp the recognized text over the scan as render-mode-3
// (invisible) glyphs. Dropping that layer silently empties the page, so when
// nearly ALL body text is invisible AND a page-covering image is present the
// page is a searchable scan: the OCR text is the page's real content.

/** minimum body chars before the OCR-overlay verdict is trusted */
const OCR_OVERLAY_MIN_CHARS = 20
/** visible body chars a page must have had before wholesale loss trips the guard */
const TEXT_LOST_MIN_CHARS = 10
/** share of invisible body chars that marks the page a searchable scan */
const OCR_OVERLAY_MIN_SHARE = 0.8
/** an image covering this share of the page area counts as the scan layer */
const OCR_SCAN_IMAGE_COVER = 0.8
/** char boxes shorter than this × fontSize are degenerate (OCR junk boxes) */
const OCR_BOX_MIN_H_RATIO = 0.3
/** nominal font sizes below this are OCR placeholders, resynth from box height */
const OCR_MIN_FONT_SIZE = 2
/** tight glyph boxes run about this share of an em (cap-height-ish) */
const OCR_BOX_EM_RATIO = 0.72
/** this many page-covering tiles mark a searchable scan graphics-dominant (P29 B) */
const OCR_GRAPHICS_MIN_TILES = 3

/** page-covering raw image objects: how many, and the topmost one's object index */
function pageCoveringImageStats(
  m: PdfiumModule,
  page: number,
  widthPt: number,
  heightPt: number,
): { count: number; topIndex: number } {
  const pageArea = widthPt * heightPt
  if (pageArea <= 0) return { count: 0, topIndex: -1 }
  return withAlloc(m, 16, (f4) => {
    const total = m._FPDFPage_CountObjects(page)
    let count = 0
    let topIndex = -1
    for (let i = 0; i < total; i++) {
      const obj = m._FPDFPage_GetObject(page, i)
      if (m._FPDFPageObj_GetType(obj) !== FPDF_PAGEOBJ_IMAGE) continue
      if (!m._FPDFPageObj_GetBounds(obj, f4, f4 + 4, f4 + 8, f4 + 12)) continue
      const w = Math.min(m.HEAPF32[(f4 >> 2) + 2]!, widthPt) - Math.max(m.HEAPF32[f4 >> 2]!, 0)
      const h =
        Math.min(m.HEAPF32[(f4 >> 2) + 3]!, heightPt) - Math.max(m.HEAPF32[(f4 >> 2) + 1]!, 0)
      if (w > 0 && h > 0 && (w * h) / pageArea >= OCR_SCAN_IMAGE_COVER) {
        count++
        topIndex = i
      }
    }
    return { count, topIndex }
  })
}

/**
 * Flip a searchable scan's OCR layer visible. OCR geometry is sloppy in both
 * directions — pr-136-style zero-height boxes next to sane font sizes, and
 * jpeg2000-style 1pt nominal sizes next to sane boxes — so each char gets the
 * broken half resynthesized from the intact one. Returns true when the page
 * was recovered (caller drops the scan image and surfaces a warning).
 */
function recoverOcrOverlay(
  m: PdfiumModule,
  page: number,
  chars: PdfChar[],
  widthPt: number,
  heightPt: number,
): boolean {
  const body = chars.filter((c) => !c.isGenerated && !isWhitespaceCode(c.code))
  if (body.length < OCR_OVERLAY_MIN_CHARS) return false
  const invisible = body.filter((c) => c.invisible).length
  if (invisible / body.length < OCR_OVERLAY_MIN_SHARE) return false
  if (pageCoveringImageStats(m, page, widthPt, heightPt).count === 0) return false

  for (const c of chars) {
    if (c.invisible) delete c.invisible
    if (isWhitespaceCode(c.code)) continue
    const boxH = c.box.y1 - c.box.y0
    if (c.fontSize < OCR_MIN_FONT_SIZE && boxH >= OCR_MIN_FONT_SIZE) {
      c.fontSize = boxH / OCR_BOX_EM_RATIO
    } else if (c.fontSize >= OCR_MIN_FONT_SIZE && boxH < OCR_BOX_MIN_H_RATIO * c.fontSize) {
      // rebuild the tight box around the origin: ~0.8 em ascent, ~0.2 descent
      c.box = {
        x0: c.box.x0,
        x1: c.box.x1,
        y0: c.originY - 0.2 * c.fontSize,
        y1: c.originY + 0.8 * c.fontSize,
      }
      c.looseBox = c.box
    }
  }
  return true
}

// ── full-page background stack (P9 B) ──

/** a background-layer object must cover this share of BOTH page dimensions */
const BG_LAYER_DIM_COVER = 0.9
/**
 * beyond the z-order prefix the bar is higher (P16 B): a ~92%-cover white
 * CONTENT card over a photo must not extend the stack — only true full-bleed
 * washes/wallpapers do, or the card's own text would leave the flow
 */
const BG_EXT_DIM_COVER = 0.96
/** raster scale for background renders (px per pt) — backgrounds are soft imagery */
const BG_RENDER_SCALE = 1.5

/** stack path fills at/below this alpha are tints, not the page's background */
const BG_FILL_MIN_ALPHA = 128
/** every fill channel at/above this is the default paper tone, not a background */
const BG_FILL_WHITE_MIN = 0xf2

/**
 * The page's background stack: everything below (and including) the LAST
 * page-covering paint object (path fill / image / gradient shading).
 * `count` is the stack length; `rasterIndices` marks the page-covering
 * members a flat w:background color cannot carry — shadings/images, and
 * (P10 B) opaque non-white vector fills: w:background is one color per
 * DOCUMENT, so a lone colored page (deck cover, closing slide) loses its
 * wash unless it rides a per-page behindDoc render.
 *
 * The stack is not required to be a z-order prefix (P16 B): slide templates
 * leave wallpaper + junk furniture at the bottom and designers blank them
 * with a page wash drawn later; the wash's ExtGState transparency is
 * invisible to the object API, so only a true composite render (which also
 * bakes the buried furniture, occluded exactly as authored) is faithful.
 * Without a page-covering raster member the stack never extends past plain
 * washes — `rasterIndices` stays empty and the caller keeps the page as-is.
 */
/** a non-prefix page-covering image this transparent is a watermark overlay, not a wash */
const OVERLAY_TRANSPARENT_SHARE = 0.5

/** share of meaningfully transparent pixels in the image object's own render */
function imageTransparentShare(m: PdfiumModule, doc: number, page: number, obj: number): number {
  const px = renderImageObjRgba(m, doc, page, obj, 1)
  if (!px) return 0
  let transparent = 0
  let sampled = 0
  for (let i = 3; i < px.rgba.length; i += 16) {
    sampled++
    if (px.rgba[i]! < ALPHA_OPAQUE_MIN) transparent++
  }
  return sampled > 0 ? transparent / sampled : 0
}

/** a background-wrapper form may carry at most this many furniture text objects */
const BG_FORM_MAX_TEXT_OBJS = 10

/**
 * Background-wrapper form profile: paint children plus its (few) furniture
 * text objects. Gradient washes hide inside XObjects together with the
 * template's header/footer text — those text handles get buried into the
 * background bake with the rest of the form. A form carrying MORE text is a
 * content wrapper (some generators wrap the whole page) and stays out: baking
 * it would turn the body into pixels. Returns null when not admissible.
 */
function formBgProfile(
  m: PdfiumModule,
  form: number,
  depth = 0,
): { paints: number; textHandles: number[] } | null {
  if (
    depth > 3 ||
    typeof m._FPDFFormObj_CountObjects !== 'function' ||
    typeof m._FPDFFormObj_GetObject !== 'function'
  ) {
    return null
  }
  const children = m._FPDFFormObj_CountObjects(form)
  if (children === 0 || children > 64) return null
  let paints = 0
  const textHandles: number[] = []
  for (let i = 0; i < children; i++) {
    const child = m._FPDFFormObj_GetObject(form, i)
    if (!child) return null
    const type = m._FPDFPageObj_GetType(child)
    if (type === FPDF_PAGEOBJ_TEXT) {
      textHandles.push(child)
      if (textHandles.length > BG_FORM_MAX_TEXT_OBJS) return null
    } else if (type === FPDF_PAGEOBJ_FORM) {
      const sub = formBgProfile(m, child, depth + 1)
      if (sub === null) return null
      paints += sub.paints
      textHandles.push(...sub.textHandles)
      if (textHandles.length > BG_FORM_MAX_TEXT_OBJS) return null
    } else {
      paints++
    }
  }
  return paints > 0 ? { paints, textHandles } : null
}

function detectBackgroundStack(
  m: PdfiumModule,
  doc: number,
  page: number,
  widthPt: number,
  heightPt: number,
): {
  count: number
  rasterIndices: number[]
  formTextHandles: Set<number>
  /** every raster member is a white-reporting fill — probe the bake first */
  suspectOnly: boolean
} {
  const total = m._FPDFPage_CountObjects(page)
  const rasterIndices: number[] = []
  const formTextHandles = new Set<number>()
  let count = 0
  let sureCount = 0
  let inPrefix = true
  withAlloc(m, 4 * 4, (f4) => {
    for (let i = 0; i < total; i++) {
      const obj = m._FPDFPage_GetObject(page, i)
      const type = m._FPDFPageObj_GetType(obj)
      // a page-covering background-wrapper FORM in the bottom prefix joins the
      // stack (gradient washes hide inside XObjects); above live content it
      // could be a watermark riding on text, so only the prefix admits forms
      const formProfile = type === FPDF_PAGEOBJ_FORM && inPrefix ? formBgProfile(m, obj) : null
      const isPaint =
        type === FPDF_PAGEOBJ_PATH || type === FPDF_PAGEOBJ_IMAGE || type === FPDF_PAGEOBJ_SHADING
      if (
        (!isPaint && formProfile === null) ||
        !m._FPDFPageObj_GetBounds(obj, f4, f4 + 4, f4 + 8, f4 + 12)
      ) {
        inPrefix = false
        continue
      }
      const x0 = m.HEAPF32[f4 >> 2]!
      const y0 = m.HEAPF32[(f4 >> 2) + 1]!
      const x1 = m.HEAPF32[(f4 >> 2) + 2]!
      const y1 = m.HEAPF32[(f4 >> 2) + 3]!
      const coveredW = Math.min(x1, widthPt) - Math.max(x0, 0)
      const coveredH = Math.min(y1, heightPt) - Math.max(y0, 0)
      const cover = inPrefix ? BG_LAYER_DIM_COVER : BG_EXT_DIM_COVER
      if (coveredW < widthPt * cover || coveredH < heightPt * cover) {
        inPrefix = false
        continue
      }
      // a page-covering image drawn OVER the content that is mostly
      // transparent pixels is a watermark riding on top, not a background
      // wash — extending the stack would bury the live text under it (P27)
      if (
        !inPrefix &&
        type === FPDF_PAGEOBJ_IMAGE &&
        imageTransparentShare(m, doc, page, obj) > OVERLAY_TRANSPARENT_SHARE
      ) {
        continue
      }
      const sureRaster = type !== FPDF_PAGEOBJ_PATH || pathIsColoredFill(m, obj, f4)
      // pattern/gradient fills voice the white fallback color through the
      // color API, indistinguishable from a real white wash here — a
      // bottom-prefix one is a raster SUSPECT: the caller probe-bakes the
      // stack and deactivates it when the result is near-white (plain page)
      const suspect = !sureRaster && inPrefix && pathIsOpaqueFill(m, obj, f4)
      const raster = sureRaster || suspect
      // plain (near-white / translucent) washes only ride the stack when a
      // raster member is already below them — alone they never pull
      // mid-stack content into the bake
      if (inPrefix || raster || rasterIndices.length > 0) count = i + 1
      if (raster) {
        rasterIndices.push(i)
        if (sureRaster) sureCount++
      }
      if (formProfile !== null && raster) {
        for (const h of formProfile.textHandles) formTextHandles.add(h)
      }
    }
  })
  return { count, rasterIndices, formTextHandles, suspectOnly: sureCount === 0 }
}

/** handles of text objects buried under the background stack (P16 B) */
function buriedTextObjects(m: PdfiumModule, page: number, stackCount: number): Set<number> {
  const handles = new Set<number>()
  for (let i = 0; i < stackCount; i++) {
    const obj = m._FPDFPage_GetObject(page, i)
    if (obj && m._FPDFPageObj_GetType(obj) === FPDF_PAGEOBJ_TEXT) handles.add(obj)
  }
  return handles
}

/** a filled path whose color is opaque and clearly not the white paper tone */
function pathIsColoredFill(m: PdfiumModule, obj: number, scratch16: number): boolean {
  if (!pathIsOpaqueFill(m, obj, scratch16)) return false
  const r = m.HEAPU32[scratch16 >> 2]!
  const g = m.HEAPU32[(scratch16 >> 2) + 1]!
  const b = m.HEAPU32[(scratch16 >> 2) + 2]!
  return r < BG_FILL_WHITE_MIN || g < BG_FILL_WHITE_MIN || b < BG_FILL_WHITE_MIN
}

/** opaque filled path of ANY reported color — leaves the fill color in scratch16 */
function pathIsOpaqueFill(m: PdfiumModule, obj: number, scratch16: number): boolean {
  if (!m._FPDFPath_GetDrawMode(obj, scratch16, scratch16 + 4)) return false
  if ((m.HEAPU32[scratch16 >> 2]! | 0) === 0) return false // not filled
  if (!m._FPDFPageObj_GetFillColor(obj, scratch16, scratch16 + 4, scratch16 + 8, scratch16 + 12)) {
    return false
  }
  return m.HEAPU32[(scratch16 >> 2) + 3]! >= BG_FILL_MIN_ALPHA
}

/** every sampled channel at/above this = white page, background not worth carrying */
const BG_RENDER_WHITE_MIN = 0xf2

/** samples every 4th pixel — a near-white render is the default paper, not a background */
function rgbaIsNearWhite(rgba: Uint8Array): boolean {
  for (let i = 0; i < rgba.length; i += 16) {
    if (
      rgba[i]! < BG_RENDER_WHITE_MIN ||
      rgba[i + 1]! < BG_RENDER_WHITE_MIN ||
      rgba[i + 2]! < BG_RENDER_WHITE_MIN
    ) {
      return false
    }
  }
  return true
}

/**
 * Render only the page's background stack: a SECOND instance of the page is
 * loaded (each FPDF_LoadPage parses its own object list), every object above
 * the stack is removed from that instance, and the remainder renders onto a
 * white-backed bitmap. The shared page dictionary is never regenerated, so
 * the main page instance is unaffected. Near-white results return null.
 */
function renderPageBackground(
  m: PdfiumModule,
  doc: number,
  pageIndex: number,
  stackCount: number,
): PageRender | null {
  const page = m._FPDF_LoadPage(doc, pageIndex)
  if (!page) return null
  try {
    const total = m._FPDFPage_CountObjects(page)
    for (let i = total - 1; i >= stackCount; i--) {
      const obj = m._FPDFPage_GetObject(page, i)
      if (obj && m._FPDFPage_RemoveObject(page, obj)) m._FPDFPageObj_Destroy(obj)
    }
    const width = Math.max(1, Math.round(m._FPDF_GetPageWidthF(page) * BG_RENDER_SCALE))
    const height = Math.max(1, Math.round(m._FPDF_GetPageHeightF(page) * BG_RENDER_SCALE))
    const bitmap = m._FPDFBitmap_Create(width, height, 0)
    if (!bitmap) return null
    try {
      m._FPDFBitmap_FillRect(bitmap, 0, 0, width, height, 0xffffffff)
      m._FPDF_RenderPageBitmap(bitmap, page, 0, 0, width, height, 0, 0)
      const px = bitmapToRgba(m, bitmap)
      if (!px || rgbaIsNearWhite(px.rgba)) return null
      return {
        data: encodeRgbaPng(px.rgba, px.width, px.height),
        mime: 'image/png',
        pixelWidth: px.width,
        pixelHeight: px.height,
      }
    } finally {
      m._FPDFBitmap_Destroy(bitmap)
    }
  } finally {
    m._FPDF_ClosePage(page)
  }
}

/** Render the whole page to an opaque white-backed PNG (fallback fidelity path). */
export function renderPagePng(m: PdfiumModule, page: number, scale: number): PageRender | null {
  const width = Math.max(1, Math.round(m._FPDF_GetPageWidthF(page) * scale))
  const height = Math.max(1, Math.round(m._FPDF_GetPageHeightF(page) * scale))
  const bitmap = m._FPDFBitmap_Create(width, height, 0)
  if (!bitmap) return null
  try {
    m._FPDFBitmap_FillRect(bitmap, 0, 0, width, height, 0xffffffff)
    m._FPDF_RenderPageBitmap(bitmap, page, 0, 0, width, height, 0, 0)
    const px = bitmapToRgba(m, bitmap)
    if (!px) return null
    return {
      data: encodeRgbaPng(px.rgba, px.width, px.height),
      mime: 'image/png',
      pixelWidth: px.width,
      pixelHeight: px.height,
    }
  } finally {
    m._FPDFBitmap_Destroy(bitmap)
  }
}

/**
 * Render one page region to an opaque PNG (P4 vector-illustration path). The
 * page renders scaled with a negative offset so only the region lands on the
 * bitmap — PDFium clips everything outside it.
 */
export function renderPageRegionPng(
  m: PdfiumModule,
  page: number,
  region: Rect,
  scale: number,
): PageRender | null {
  const pageWidthPt = m._FPDF_GetPageWidthF(page)
  const pageHeightPt = m._FPDF_GetPageHeightF(page)
  const width = Math.max(1, Math.round((region.x1 - region.x0) * scale))
  const height = Math.max(1, Math.round((region.y1 - region.y0) * scale))
  const bitmap = m._FPDFBitmap_Create(width, height, 0)
  if (!bitmap) return null
  try {
    m._FPDFBitmap_FillRect(bitmap, 0, 0, width, height, 0xffffffff)
    // device origin is the page's TOP-left corner (y down)
    m._FPDF_RenderPageBitmap(
      bitmap,
      page,
      Math.round(-region.x0 * scale),
      Math.round(-(pageHeightPt - region.y1) * scale),
      Math.round(pageWidthPt * scale),
      Math.round(pageHeightPt * scale),
      0,
      0,
    )
    const px = bitmapToRgba(m, bitmap)
    if (!px) return null
    return {
      data: encodeRgbaPng(px.rgba, px.width, px.height),
      mime: 'image/png',
      pixelWidth: px.width,
      pixelHeight: px.height,
    }
  } finally {
    m._FPDFBitmap_Destroy(bitmap)
  }
}

// ── decor shadings (P19, canvas pages only) ──
// Non-page-covering gradient SHADING objects (slide title bars, accent
// strips) have no path/image representation anywhere in the pipeline and
// were dropped outright. They rasterize here on a TRANSPARENT bitmap so the
// canvas path can pin them at their measured position; the flow path never
// reads the field, so its output is byte-identical.

/** a shading covering at/above this share of the page is background, not decor */
const DECOR_SHADING_MAX_COVER = 0.85
/** ignore sub-2pt slivers (hairline gradient artifacts) */
const DECOR_SHADING_MIN_DIM_PT = 2
const DECOR_RENDER_SCALE = 2

/** rasterize the page's decor shadings (transparent bg), one image per object */
function readShadingDecor(
  m: PdfiumModule,
  doc: number,
  pageIndex: number,
  widthPt: number,
  heightPt: number,
  stackSkip: number,
  vectorRegions: readonly Rect[],
): ExtractedImage[] {
  const page = m._FPDF_LoadPage(doc, pageIndex)
  if (!page) return []
  try {
    const total = m._FPDFPage_CountObjects(page)
    const keep = new Map<number, Rect>()
    withAlloc(m, 4 * 4, (f4) => {
      for (let i = stackSkip; i < total; i++) {
        const obj = m._FPDFPage_GetObject(page, i)
        if (!obj || m._FPDFPageObj_GetType(obj) !== FPDF_PAGEOBJ_SHADING) continue
        if (!m._FPDFPageObj_GetBounds(obj, f4, f4 + 4, f4 + 8, f4 + 12)) continue
        const box: Rect = {
          x0: Math.max(0, m.HEAPF32[f4 >> 2]!),
          y0: Math.max(0, m.HEAPF32[(f4 >> 2) + 1]!),
          x1: Math.min(widthPt, m.HEAPF32[(f4 >> 2) + 2]!),
          y1: Math.min(heightPt, m.HEAPF32[(f4 >> 2) + 3]!),
        }
        const w = box.x1 - box.x0
        const h = box.y1 - box.y0
        if (w < DECOR_SHADING_MIN_DIM_PT || h < DECOR_SHADING_MIN_DIM_PT) continue
        if (w * h >= DECOR_SHADING_MAX_COVER * widthPt * heightPt) continue
        const cx = (box.x0 + box.x1) / 2
        const cy = (box.y0 + box.y1) / 2
        // a vector-region bitmap already shows it
        if (vectorRegions.some((v) => cx >= v.x0 && cx <= v.x1 && cy >= v.y0 && cy <= v.y1)) {
          continue
        }
        keep.set(i, box)
      }
    })
    if (keep.size === 0) return []
    // strip everything else from this SECOND page instance and render the
    // shadings alone over a fully transparent bitmap
    for (let i = total - 1; i >= 0; i--) {
      if (keep.has(i)) continue
      const obj = m._FPDFPage_GetObject(page, i)
      if (obj && m._FPDFPage_RemoveObject(page, obj)) m._FPDFPageObj_Destroy(obj)
    }
    const bw = Math.max(1, Math.round(widthPt * DECOR_RENDER_SCALE))
    const bh = Math.max(1, Math.round(heightPt * DECOR_RENDER_SCALE))
    const bitmap = m._FPDFBitmap_Create(bw, bh, 1)
    if (!bitmap) return []
    try {
      m._FPDFBitmap_FillRect(bitmap, 0, 0, bw, bh, 0x00000000)
      m._FPDF_RenderPageBitmap(bitmap, page, 0, 0, bw, bh, 0, 0)
      const px = bitmapToRgba(m, bitmap)
      if (!px) return []
      const out: ExtractedImage[] = []
      for (const [z, box] of keep) {
        // crop the object's bounds out of the page render (device y is down)
        const cx0 = Math.max(0, Math.floor(box.x0 * DECOR_RENDER_SCALE))
        const cy0 = Math.max(0, Math.floor((heightPt - box.y1) * DECOR_RENDER_SCALE))
        const cw = Math.min(
          bw - cx0,
          Math.max(1, Math.ceil((box.x1 - box.x0) * DECOR_RENDER_SCALE)),
        )
        const ch = Math.min(
          bh - cy0,
          Math.max(1, Math.ceil((box.y1 - box.y0) * DECOR_RENDER_SCALE)),
        )
        if (cw <= 0 || ch <= 0) continue
        const rgba = new Uint8Array(cw * ch * 4)
        for (let y = 0; y < ch; y++) {
          const src = ((cy0 + y) * px.width + cx0) * 4
          rgba.set(px.rgba.subarray(src, src + cw * 4), y * cw * 4)
        }
        out.push({
          box,
          data: encodeRgbaPng(rgba, cw, ch),
          mime: 'image/png',
          pixelWidth: cw,
          pixelHeight: ch,
          z,
        })
      }
      return out
    } finally {
      m._FPDFBitmap_Destroy(bitmap)
    }
  } finally {
    m._FPDF_ClosePage(page)
  }
}

/** Load a page by index and render it whole (P4 low-confidence downgrade path). */
export function renderPageByIndexPng(
  m: PdfiumModule,
  doc: number,
  pageIndex: number,
  scale: number,
): PageRender | null {
  const page = m._FPDF_LoadPage(doc, pageIndex)
  if (!page) return null
  try {
    return renderPagePng(m, page, scale)
  } finally {
    m._FPDF_ClosePage(page)
  }
}

function probeStructTree(m: PdfiumModule, page: number): boolean {
  const tree = m._FPDF_StructTree_GetForPage(page)
  if (!tree) return false
  const hasChildren = m._FPDF_StructTree_CountChildren(tree) > 0
  m._FPDF_StructTree_Close(tree)
  return hasChildren
}

/** Quality verdicts that force the bitmap fallback (computed on raw chars). */
function assessQuality(
  chars: PdfChar[],
  images: ExtractedImage[],
  pageRect: Rect,
  rotation: number,
): { degraded: boolean; reason?: string; scanned: boolean; badUnicodeRatio: number } {
  const textChars = chars.filter((c) => !c.isGenerated && !isWhitespaceCode(c.code))
  const badUnicodeRatio =
    textChars.length >= BAD_UNICODE_MIN_CHARS
      ? textChars.filter((c) => isBadUnicode(c.code)).length / textChars.length
      : 0

  if (textChars.length <= SCANNED_MAX_CHARS) {
    const pageArea = rectArea(pageRect)
    const covered = images.some(
      (img) => pageArea > 0 && intersectArea(img.box, pageRect) / pageArea > SCANNED_IMAGE_COVER,
    )
    if (covered) return { degraded: false, scanned: true, badUnicodeRatio }
  }

  // /Rotate pages are geometry-normalized into display space at extraction
  // (P27) — no hard degrade; producers that pre-rotate content anyway are
  // caught by the angled-char check below exactly like unrotated pages.
  void rotation

  if (textChars.length >= BAD_UNICODE_MIN_CHARS) {
    if (badUnicodeRatio > BAD_UNICODE_RATIO) {
      return { degraded: true, reason: 'bad-tounicode', scanned: false, badUnicodeRatio }
    }
    const angled = textChars.filter((c) => Math.abs(c.angle) > ANGLED_CHAR_RAD).length
    if (angled / textChars.length > ANGLED_RATIO) {
      return { degraded: true, reason: 'vertical-text', scanned: false, badUnicodeRatio }
    }
  }
  return { degraded: false, scanned: false, badUnicodeRatio }
}

/** Info-dictionary metadata for the P19 slide prior (Producer/Creator). */
export interface DocMetadata {
  producer?: string
  creator?: string
}

/** one Info-dictionary text field, decoded from PDFium's UTF-16LE buffer */
function readMetaText(m: PdfiumModule, doc: number, tag: string): string | undefined {
  return withAlloc(m, tag.length + 1, (tagPtr) => {
    for (let i = 0; i < tag.length; i++) m.HEAPU8[tagPtr + i] = tag.charCodeAt(i)
    m.HEAPU8[tagPtr + tag.length] = 0
    const needed = m._FPDF_GetMetaText(doc, tagPtr, 0, 0)
    if (needed <= 2) return undefined
    return withAlloc(m, needed, (buf) => {
      m._FPDF_GetMetaText(doc, tagPtr, buf, needed)
      const text = new TextDecoder('utf-16le')
        .decode(m.HEAPU8.subarray(buf, buf + needed - 2))
        .trim()
      return text.length > 0 ? text : undefined
    })
  })
}

export function readDocMetadata(m: PdfiumModule, doc: number): DocMetadata {
  return {
    producer: readMetaText(m, doc, 'Producer'),
    creator: readMetaText(m, doc, 'Creator'),
  }
}

/**
 * CropBox lower-left origin (P21): PDFium reports page width/height in crop
 * space and renders the crop area, but char/object coordinates stay in
 * absolute user space — on a page cropped away from (0,0) every coordinate
 * must shift by the crop origin or the whole rebuild is offset by it.
 */
function pageOriginOffset(m: PdfiumModule, page: number): { dx: number; dy: number } {
  return withAlloc(m, 16, (ptr) => {
    // some producers write flipped boxes (top < bottom / left > right); PDFium
    // reports |height| but keeps object coordinates in the authored space, so
    // the origin is the MINIMUM corner, not the nominal left/bottom (P27)
    const read = (fn: (p: number, l: number, b: number, r: number, t: number) => number) =>
      fn(page, ptr, ptr + 4, ptr + 8, ptr + 12)
        ? {
            l: Math.min(m.HEAPF32[ptr >> 2]!, m.HEAPF32[(ptr >> 2) + 2]!),
            b: Math.min(m.HEAPF32[(ptr >> 2) + 1]!, m.HEAPF32[(ptr >> 2) + 3]!),
          }
        : null
    const box = read(m._FPDFPage_GetCropBox.bind(m)) ?? read(m._FPDFPage_GetMediaBox.bind(m))
    return box ? { dx: -box.l, dy: -box.b } : { dx: 0, dy: 0 }
  })
}

const shiftRect = (r: Rect, dx: number, dy: number): void => {
  r.x0 += dx
  r.x1 += dx
  r.y0 += dy
  r.y1 += dy
}

/**
 * Display-space normalization for /Rotate pages (P27): PDFium reports the
 * rotated display dims and renders display space, but every text / object
 * coordinate stays in the unrotated user space. Rotating the geometry here
 * lets rotated pages (landscape tables in portrait media boxes) flow through
 * the regular pipeline instead of hard-degrading to a full-page bitmap.
 * Note: embedded image PIXELS keep their authored orientation — only their
 * placement box is normalized.
 */
// ── AcroForm widget checkboxes (P29) ──
// Interactive checkbox/radio fields are ANNOTATIONS, not page content: the
// object walk never sees them and every form lost its boxes. Synthesize the
// same '☐' glyph chars the drawn-checkbox pass mints (checked state included)
// so they ride line grouping, cell routing, and the form-table detector.

/** widget squares outside this range are decoration or full-field outlines */
const WIDGET_BOX_MIN_PT = 4
const WIDGET_BOX_MAX_PT = 24
/** snap a glyph beside its label when the gap is under this (pt) */
const WIDGET_SNAP_MAX_GAP_PT = 24

/** UTF-16LE annotation string/name value ('' when absent or API missing) */
function annotStringValue(m: PdfiumModule, annot: number, key: string): string {
  if (typeof m._FPDFAnnot_GetStringValue !== 'function') return ''
  return withAlloc(m, key.length + 1, (kb) => {
    for (let i = 0; i < key.length; i++) m.HEAPU8[kb + i] = key.charCodeAt(i)
    m.HEAPU8[kb + key.length] = 0
    const len = m._FPDFAnnot_GetStringValue!(annot, kb, 0, 0)
    if (len <= 2) return ''
    return withAlloc(m, len, (buf) => {
      m._FPDFAnnot_GetStringValue!(annot, kb, buf, len)
      return new TextDecoder('utf-16le').decode(m.HEAPU8.subarray(buf, buf + len - 2))
    })
  })
}

function readWidgetCheckboxChars(m: PdfiumModule, doc: number, page: number): PdfChar[] {
  if (
    typeof m._FPDFPage_GetAnnotCount !== 'function' ||
    typeof m._FPDFPage_GetAnnot !== 'function' ||
    typeof m._FPDFPage_CloseAnnot !== 'function' ||
    typeof m._FPDFAnnot_GetSubtype !== 'function' ||
    typeof m._FPDFAnnot_GetRect !== 'function' ||
    typeof m._FPDFAnnot_GetFormFieldType !== 'function' ||
    typeof m._FPDFAnnot_IsChecked !== 'function' ||
    typeof m._PDFiumExt_OpenFormFillInfo !== 'function' ||
    typeof m._PDFiumExt_InitFormFillEnvironment !== 'function' ||
    typeof m._PDFiumExt_ExitFormFillEnvironment !== 'function' ||
    typeof m._PDFiumExt_CloseFormFillInfo !== 'function'
  ) {
    return []
  }
  const total = m._FPDFPage_GetAnnotCount(page)
  if (total <= 0) return []
  const out: PdfChar[] = []
  const formInfo = m._PDFiumExt_OpenFormFillInfo()
  if (!formInfo) return []
  let form = 0
  try {
    form = m._PDFiumExt_InitFormFillEnvironment(doc, formInfo)
    if (!form) return []
    withAlloc(m, 4 * 4, (f4) => {
      for (let i = 0; i < total; i++) {
        const annot = m._FPDFPage_GetAnnot!(page, i)
        if (!annot) continue
        try {
          if (m._FPDFAnnot_GetSubtype!(annot) !== FPDF_ANNOT_WIDGET) continue
          const fieldType = m._FPDFAnnot_GetFormFieldType!(form, annot)
          const radio = fieldType === FPDF_FORMFIELD_RADIOBUTTON
          if (fieldType !== FPDF_FORMFIELD_CHECKBOX && !radio) continue
          if (!m._FPDFAnnot_GetRect!(annot, f4)) continue
          const v = m.HEAPF32.subarray(f4 >> 2, (f4 >> 2) + 4) // left top right bottom
          const box: Rect = {
            x0: Math.min(v[0]!, v[2]!),
            x1: Math.max(v[0]!, v[2]!),
            y0: Math.min(v[1]!, v[3]!),
            y1: Math.max(v[1]!, v[3]!),
          }
          const side = Math.max(box.x1 - box.x0, box.y1 - box.y0)
          if (side < WIDGET_BOX_MIN_PT || side > WIDGET_BOX_MAX_PT) continue
          // IsChecked misses widgets whose value lives on the parent field
          // dict only — the appearance state name is the standard fallback
          // ('Off' = unchecked, any on-state name = checked)
          const appearanceState = annotStringValue(m, annot, 'AS')
          const checked =
            m._FPDFAnnot_IsChecked!(form, annot) === 1 ||
            (appearanceState !== '' && appearanceState !== 'Off')
          const code = radio ? (checked ? 0x25c9 : 0x25cb) : checked ? 0x2612 : 0x2610
          out.push({
            code,
            text: String.fromCodePoint(code),
            box,
            looseBox: { ...box },
            originX: box.x0,
            originY: box.y0,
            angle: 0,
            fontSize: Math.min(Math.max(side, 6), 14),
            fontWeight: 400,
            fontFamily: 'Segoe UI Symbol',
            italic: false,
            color: '000000',
            isGenerated: true, // synthesized — not in the content stream
            isHyphen: false,
            script: scriptOf(code),
          })
        } finally {
          m._FPDFPage_CloseAnnot!(annot)
        }
      }
    })
  } finally {
    if (form) m._PDFiumExt_ExitFormFillEnvironment(form)
    m._PDFiumExt_CloseFormFillInfo(formInfo)
  }
  return out
}

function rotateToDisplay(
  rotation: number,
  widthPt: number,
  heightPt: number,
): {
  point: (x: number, y: number) => [number, number]
  rect: (r: Rect) => void
  angleDelta: number
} {
  const userW = rotation % 2 === 1 ? heightPt : widthPt
  const userH = rotation % 2 === 1 ? widthPt : heightPt
  const point: (x: number, y: number) => [number, number] =
    rotation === 1
      ? (x, y) => [y, userW - x]
      : rotation === 2
        ? (x, y) => [userW - x, userH - y]
        : (x, y) => [userH - y, x]
  const rect = (r: Rect): void => {
    const [ax, ay] = point(r.x0, r.y0)
    const [bx, by] = point(r.x1, r.y1)
    r.x0 = Math.min(ax, bx)
    r.x1 = Math.max(ax, bx)
    r.y0 = Math.min(ay, by)
    r.y1 = Math.max(ay, by)
  }
  // FPDFText angles are clockwise-positive relative to display: upright text
  // on a /Rotate 90 page reads back as -π/2, so the delta ADDS the rotation
  return { point, rect, angleDelta: rotation * (Math.PI / 2) }
}

/** wrap an angle into (-π, π] */
const normalizeAngle = (a: number): number => a - 2 * Math.PI * Math.round(a / (2 * Math.PI))

/** Extract one page: chars + images + geometry + quality flags (+ fallback render). */
export function extractPage(
  m: PdfiumModule,
  doc: number,
  pageIndex: number,
  options: ExtractOptions = {},
): ExtractedPage {
  const page = m._FPDF_LoadPage(doc, pageIndex)
  if (!page) throw new Error(`could not load page ${pageIndex + 1}`)
  try {
    const widthPt = m._FPDF_GetPageWidthF(page)
    const heightPt = m._FPDF_GetPageHeightF(page)
    const rotation = m._FPDFPage_GetRotation(page)
    const pageRect: Rect = { x0: 0, y0: 0, x1: widthPt, y1: heightPt }

    const { dx, dy } = pageOriginOffset(m, page)
    const shifted = dx !== 0 || dy !== 0

    const textPage = m._FPDFText_LoadPage(page)
    let chars: PdfChar[] = []
    try {
      if (textPage) {
        chars = readChars(m, textPage)
        // crop-origin shift BEFORE the ink check: it renders in crop space
        if (shifted) {
          for (const c of chars) {
            shiftRect(c.box, dx, dy)
            shiftRect(c.looseBox, dx, dy)
            c.originX += dx
            c.originY += dy
          }
        }
        // /Rotate pages: normalize chars into display space BEFORE the ink
        // check — it samples the rendered (display-space) bitmap (P27)
        if (rotation === 1 || rotation === 2 || rotation === 3) {
          const t = rotateToDisplay(rotation, widthPt, heightPt)
          for (const c of chars) {
            t.rect(c.box)
            t.rect(c.looseBox)
            const [ox, oy] = t.point(c.originX, c.originY)
            c.originX = ox
            c.originY = oy
            c.angle = normalizeAngle(c.angle + t.angleDelta)
          }
        }
        // over a page-covering image (scan) the region renders find ink
        // everywhere — a guaranteed no-op that decodes the whole tile stack
        // into PDFium's page cache and exhausts the wasm heap (P29 B)
        if (pageCoveringImageStats(m, page, widthPt, heightPt).count === 0) {
          markInklessChars(m, page, chars)
        }
      }
    } finally {
      if (textPage) m._FPDFText_ClosePage(textPage)
    }

    // interactive checkbox/radio widgets → synthesized glyph chars (P29),
    // through the same crop-shift / rotation normalization as real chars
    const widgetChars = readWidgetCheckboxChars(m, doc, page)
    if (widgetChars.length > 0) {
      for (const c of widgetChars) {
        if (shifted) {
          shiftRect(c.box, dx, dy)
          shiftRect(c.looseBox, dx, dy)
          c.originX += dx
          c.originY += dy
        }
        if (rotation === 1 || rotation === 2 || rotation === 3) {
          const t = rotateToDisplay(rotation, widthPt, heightPt)
          t.rect(c.box)
          t.rect(c.looseBox)
          const [ox, oy] = t.point(c.originX, c.originY)
          c.originX = ox
          c.originY = oy
        }
        // ride the neighboring label's line: a widget box is taller than the
        // text beside it, and an unaligned baseline mints a lone line that
        // adds flow height per field (a 261-field form grew 24 → 36 pages).
        // Snap the glyph ADJACENT to its closest same-row label too — the
        // widget rect often sits across a drawn input-box boundary, and a
        // center-routed glyph then lands in the wrong (empty) grid cell.
        const cy = (c.box.y0 + c.box.y1) / 2
        const mates = chars.filter(
          (r) => !r.isGenerated && r.box.y0 <= cy && r.box.y1 >= cy && !isWhitespaceCode(r.code),
        )
        if (mates.length === 0) {
          chars.push(c)
          continue
        }
        // closest same-row label; a label INSIDE a padded widget rect
        // (overlap) counts as gap 0, center distance breaks ties — skipping
        // overlaps attached the glyph to a different field on the row
        const widgetCx = (c.box.x0 + c.box.x1) / 2
        let best = mates[0]!
        let bestGap = Infinity
        let bestKey = Infinity
        for (const r of mates) {
          const overlaps = r.box.x0 < c.box.x1 && r.box.x1 > c.box.x0
          const gap = overlaps
            ? 0
            : r.box.x0 >= c.box.x1
              ? r.box.x0 - c.box.x1
              : c.box.x0 - r.box.x1
          const key = gap + Math.abs((r.box.x0 + r.box.x1) / 2 - widgetCx) * 0.001
          if (key < bestKey) {
            bestKey = key
            bestGap = gap
            best = r
          }
        }
        c.originY = best.originY
        c.fontSize = best.fontSize
        c.fontFamily = best.fontFamily
        const side = c.fontSize
        let x0 = c.box.x0
        if (bestGap < WIDGET_SNAP_MAX_GAP_PT) {
          const bestCx = (best.box.x0 + best.box.x1) / 2
          x0 = bestCx >= widgetCx ? best.box.x0 - side - 2 : best.box.x1 + 2
        }
        c.box = { x0, x1: x0 + side, y0: c.originY, y1: c.originY + side }
        c.looseBox = { ...c.box }
        c.originX = x0
        // line grouping clusters consecutive STREAM-order chars — an appended
        // glyph would gather with its siblings into a lone "☐ ☐ ☐" line, so
        // splice it right beside its label instead
        const at = chars.indexOf(best)
        chars.splice(c.box.x0 <= best.box.x0 ? at : at + 1, 0, c)
      }
    }

    // P27: searchable scans (invisible OCR layer over a page-covering image)
    // recover their OCR text instead of silently dropping it
    const ocrTextRecovered = recoverOcrOverlay(m, page, chars, widthPt, heightPt)
    // P29 B: a graphics-dominant searchable scan (flyer/poster built from a
    // stack of page-covering tiles, often clipped patches over a base scan)
    // is its IMAGE — text-only output would ship a near-blank page, and no
    // single tile carries the composite. The page exports as its full-page
    // render, like a plain scan.
    const ocrImageDominant =
      ocrTextRecovered &&
      pageCoveringImageStats(m, page, widthPt, heightPt).count >= OCR_GRAPHICS_MIN_TILES
    const countVisibleBody = (list: PdfChar[]): number => {
      let n = 0
      for (const c of list) {
        if (!c.isGenerated && !c.invisible && !isWhitespaceCode(c.code)) n++
      }
      return n
    }
    const rawVisibleBody = countVisibleBody(chars)

    // P9 B: a page-covering background stack with raster members (gradient
    // shadings / wallpaper images) becomes one behindDoc bitmap. Text-less
    // pages stay out — the scanned/degraded fallback serves them better, and
    // excluding a scan's own image here would break scanned detection.
    const bgStack = detectBackgroundStack(m, doc, page, widthPt, heightPt)
    const hasRealText =
      chars.filter((c) => !c.isGenerated && !isWhitespaceCode(c.code)).length > SCANNED_MAX_CHARS
    // P29 B: a graphics-dominant scan's tile stack must NOT bake — the bake
    // decodes every tile (heap death); the tiles embed raw via readImages
    let bgActive = hasRealText && bgStack.rasterIndices.length > 0 && !ocrImageDominant
    // suspect-only stacks (white-reporting fills that may be pattern washes):
    // probe the bake up front — a near-white result means a plain white page,
    // and activating the stack would drop its below-count vectors for nothing
    let probedBg: PageRender | null | undefined
    if (bgActive && bgStack.suspectOnly) {
      probedBg = renderPageBackground(m, doc, pageIndex, bgStack.count)
      if (probedBg === null) bgActive = false
    }

    // text buried under the background stack (template junk hidden by the
    // wallpaper/wash) is baked into the bgRender bitmap, occluded exactly as
    // authored — it must leave the char stream or it resurfaces (P16 B)
    if (bgActive && bgStack.count > 0) {
      const buried = buriedTextObjects(m, page, bgStack.count)
      for (const h of bgStack.formTextHandles) buried.add(h)
      if (buried.size > 0) {
        const textPage2 = m._FPDFText_LoadPage(page)
        try {
          if (textPage2) {
            chars = chars.filter(
              (c) =>
                c.pdfCharIndex === undefined ||
                !buried.has(m._FPDFText_GetTextObject(textPage2, c.pdfCharIndex)),
            )
          }
        } finally {
          if (textPage2) m._FPDFText_ClosePage(textPage2)
        }
      }
    }

    // searchable scans drop their page-covering tiles anyway — skip the
    // decodes up front: 90+ full-page tiles otherwise exhaust the wasm heap
    // (PDFium caches every decode) and later images vanish silently
    const pageAreaRaw = widthPt * heightPt
    const coversPage = (box: Rect): boolean =>
      pageAreaRaw > 0 &&
      intersectArea(box, { x0: 0, y0: 0, x1: widthPt, y1: heightPt }) / pageAreaRaw >=
        OCR_SCAN_IMAGE_COVER
    const ocrSkip = ocrTextRecovered ? coversPage : undefined
    let images = readImages(m, doc, page, bgActive ? bgStack.count : 0, ocrSkip)
    let paths = readPaths(m, page, bgActive ? bgStack.count : 0)
    if (shifted) {
      for (const img of images) shiftRect(img.box, dx, dy)
      for (const p of paths) {
        for (const sub of p.subpaths) {
          for (const pt of sub.points) {
            pt.x += dx
            pt.y += dy
          }
        }
      }
    }
    // /Rotate pages: image boxes and path points into display space too (P27)
    if (rotation === 1 || rotation === 2 || rotation === 3) {
      const t = rotateToDisplay(rotation, widthPt, heightPt)
      for (const img of images) t.rect(img.box)
      for (const p of paths) {
        for (const sub of p.subpaths) {
          for (const pt of sub.points) {
            const [px, py] = t.point(pt.x, pt.y)
            pt.x = px
            pt.y = py
          }
        }
      }
    }
    // a recovered searchable scan keeps its OCR text ONLY: leaving the scan
    // bitmap in would paint the same words twice (pixels + recovered text).
    // Graphics-dominant scans keep the tiles instead (text re-hides below).
    if (ocrTextRecovered && !ocrImageDominant) {
      const pageArea = rectArea(pageRect)
      images = images.filter(
        (img) =>
          pageArea <= 0 || intersectArea(img.box, pageRect) / pageArea < OCR_SCAN_IMAGE_COVER,
      )
    }
    // content ENTIRELY outside the page box is invisible in every renderer —
    // PowerPoint exports keep neighboring-slide leftovers parked off-canvas
    // (whole card lists at x < 0, P14) and un-clipped they duplicate text and
    // wreck the column solve. Partially-visible content stays whole.
    const intersectsPage = (r: Rect): boolean =>
      r.x1 > -PAGE_CLIP_TOL_PT &&
      r.x0 < widthPt + PAGE_CLIP_TOL_PT &&
      r.y1 > -PAGE_CLIP_TOL_PT &&
      r.y0 < heightPt + PAGE_CLIP_TOL_PT
    chars = chars.filter((c) => intersectsPage(c.box))
    images = images.filter((img) => intersectsPage(img.box))
    paths = paths.filter((p) =>
      p.subpaths.some((sub) => {
        if (sub.points.length === 0) return false
        const xs = sub.points.map((pt) => pt.x)
        const ys = sub.points.map((pt) => pt.y)
        return intersectsPage({
          x0: Math.min(...xs),
          y0: Math.min(...ys),
          x1: Math.max(...xs),
          y1: Math.max(...ys),
        })
      }),
    )
    // quality verdicts judge the page as authored (before vector-art rewriting)
    const quality = assessQuality(chars, images, pageRect, rotation)
    if (ocrImageDominant) quality.scanned = true
    const needsRender = quality.scanned || quality.degraded
    const render = needsRender
      ? (renderPagePng(m, page, options.renderScale ?? 2) ?? undefined)
      : undefined
    // control-code chars on an otherwise-healthy page (one mojibake heading
    // on a valid CJK page) leave the stream AFTER the quality verdict — they
    // render as tofu in Word and shift every following span
    if (!needsRender) chars = chars.filter((c) => c.isGenerated || !isControlCode(c.code))

    const bgRender =
      bgActive && !needsRender
        ? (probedBg ?? renderPageBackground(m, doc, pageIndex, bgStack.count) ?? undefined)
        : undefined

    // P4: vector-illustration regions rasterize via PDFium; chars (axis
    // labels…), paths and images they cover leave the normal streams — the
    // bitmap already shows them
    let outChars = chars
    let outPaths = paths
    let outImages = images
    const vectorRegions: Rect[] = []
    if (!needsRender) {
      const centerIn = (r: Rect, region: Rect): boolean => {
        const cx = (r.x0 + r.x1) / 2
        const cy = (r.y0 + r.y1) / 2
        return cx >= region.x0 && cx <= region.x1 && cy >= region.y0 && cy <= region.y1
      }
      const regionImages: ExtractedImage[] = []
      for (const region of detectVectorRegions(paths, chars, pageRect)) {
        const art = renderPageRegionPng(m, page, region, options.renderScale ?? 2)
        if (!art) continue // rendering failed → leave the region's content in place
        vectorRegions.push(region)
        regionImages.push({
          box: region,
          data: art.data,
          mime: art.mime,
          pixelWidth: art.pixelWidth,
          pixelHeight: art.pixelHeight,
        })
      }
      if (vectorRegions.length > 0) {
        const covered = (r: Rect): boolean => vectorRegions.some((v) => centerIn(r, v))
        outChars = chars.filter((c) => !covered(c.box))
        outPaths = paths.filter(
          (p) =>
            !p.subpaths.every((sub) => {
              if (sub.points.length === 0) return true
              const xs = sub.points.map((pt) => pt.x)
              const ys = sub.points.map((pt) => pt.y)
              return covered({
                x0: Math.min(...xs),
                y0: Math.min(...ys),
                x1: Math.max(...xs),
                y1: Math.max(...ys),
              })
            }),
        )
        outImages = [
          ...images.filter((img) => !vectorRegions.some((v) => overlapRatio(img.box, v) >= 0.8)),
          ...regionImages,
        ]
      }
    }

    // P27 guard input: extraction kept NONE of the page's visible body text
    // and nothing visible (bake/region render) carries it either — the
    // pipeline degrades such a page instead of shipping it silently empty
    const textLost =
      !needsRender &&
      rawVisibleBody >= TEXT_LOST_MIN_CHARS &&
      countVisibleBody(outChars) === 0 &&
      vectorRegions.length === 0 &&
      bgRender === undefined

    const decorImages = needsRender
      ? []
      : readShadingDecor(
          m,
          doc,
          pageIndex,
          widthPt,
          heightPt,
          bgActive ? bgStack.count : 0,
          vectorRegions,
        )
    if (shifted) for (const img of decorImages) shiftRect(img.box, dx, dy)

    return {
      index: pageIndex,
      widthPt,
      heightPt,
      rotation,
      chars: outChars,
      images: outImages,
      paths: outPaths,
      degraded: quality.degraded,
      degradedReason: quality.reason,
      scanned: quality.scanned,
      hasStructTree: probeStructTree(m, page),
      render,
      vectorRegions,
      badUnicodeRatio: quality.badUnicodeRatio,
      ...(ocrTextRecovered ? { ocrTextRecovered } : {}),
      ...(ocrImageDominant ? { ocrImageKept: true } : {}),
      ...(textLost ? { textLost } : {}),
      ...(bgRender !== undefined ? { bgRender } : {}),
      ...(decorImages.length > 0 ? { decorImages } : {}),
    }
  } finally {
    m._FPDF_ClosePage(page)
  }
}
