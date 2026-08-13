import { readFileSync } from 'node:fs'
import { PDFDict, PDFDocument, PDFName, PDFRawStream, decodePDFRawStream } from 'pdf-lib'
import { fontCoversText } from './font-cmap'
import { findSystemFont, isTruetype } from './font-locate'
import { identityCffCharset, subsetTtf } from './font-subset'
import { pdfiumWasmPath } from './wasm-path'
import type {
  TextEditFailure,
  TextEditInput,
  TextEditValidation,
  TextInsertFailure,
  TextInsertInput,
} from '../shared/ipc'
import { foldRadicals } from '../shared/radicals'
import { chainLayers } from '../shared/x-layers'

export const FPDF_PAGEOBJ_TEXT = 1
const FPDF_FONT_TYPE1 = 1
const FPDF_FONT_TRUETYPE = 2

/** Emscripten module surface we call into (raw FPDF_* exports + heap access) */
export interface Pdfium {
  HEAPU8: Uint8Array
  HEAP32: Int32Array
  HEAPF32: Float32Array
  HEAPF64: Float64Array
  _malloc(size: number): number
  _free(ptr: number): void
  _PDFiumExt_Init(): void
  _PDFiumExt_OpenFileWriter(): number
  _PDFiumExt_SaveAsCopy(doc: number, writer: number): number
  _PDFiumExt_GetFileWriterSize(writer: number): number
  _PDFiumExt_GetFileWriterData(writer: number, buf: number, size: number): number
  _PDFiumExt_CloseFileWriter(writer: number): void
  _FPDF_LoadMemDocument(ptr: number, size: number, password: number): number
  _FPDF_CloseDocument(doc: number): void
  _FPDF_LoadPage(doc: number, index: number): number
  _FPDF_ClosePage(page: number): void
  _FPDF_GetPageCount(doc: number): number
  _FPDFText_LoadPage(page: number): number
  _FPDFText_ClosePage(textPage: number): void
  _FPDFText_CountChars(textPage: number): number
  _FPDFText_GetTextObject(textPage: number, index: number): number
  _FPDFText_GetLooseCharBox(textPage: number, index: number, rect: number): number
  _FPDFText_GetCharOrigin(textPage: number, index: number, x: number, y: number): number
  _FPDFText_LoadFont(doc: number, data: number, size: number, fontType: number, cid: number): number
  _FPDFText_SetText(textObj: number, text: number): number
  _FPDFPage_CountObjects(page: number): number
  _FPDFPage_GetObject(page: number, index: number): number
  _FPDFPage_RemoveObject(page: number, obj: number): number
  _FPDFPage_InsertObject(page: number, obj: number): void
  _FPDFPage_InsertObjectAtIndex?(page: number, obj: number, index: number): number
  _FPDFPage_GenerateContent(page: number): number
  _FPDFPageObj_GetType(obj: number): number
  _FPDFPageObj_Destroy(obj: number): void
  _FPDFPageObj_GetBounds(obj: number, l: number, b: number, r: number, t: number): number
  _FPDFPageObj_GetMatrix(obj: number, matrix: number): number
  _FPDFPageObj_SetMatrix(obj: number, matrix: number): number
  _FPDFPageObj_GetFillColor(obj: number, r: number, g: number, b: number, a: number): number
  _FPDFPageObj_SetFillColor(obj: number, r: number, g: number, b: number, a: number): number
  _FPDFPageObj_CreateTextObj(doc: number, font: number, size: number): number
  _FPDFTextObj_GetText(obj: number, textPage: number, buf: number, len: number): number
  _FPDFTextObj_GetFont(obj: number): number
  _FPDFTextObj_GetFontSize(obj: number, size: number): number
  _FPDFFont_GetIsEmbedded(font: number): number
  _FPDFFont_GetFontData(font: number, buffer: number, buflen: number, outBuflen: number): number
  _FPDFFont_GetGlyphWidth(font: number, glyph: number, fontSize: number, width: number): number
  _FPDFFont_GetBaseFontName(font: number, buffer: number, buflen: number): number
  _FPDFFont_GetFamilyName(font: number, buffer: number, buflen: number): number
  _FPDFPageObj_Transform(
    obj: number,
    a: number,
    b: number,
    c: number,
    d: number,
    e: number,
    f: number,
  ): void
  _FPDFPageObj_NewImageObj(doc: number): number
  _FPDFImageObj_SetBitmap(pages: number, count: number, obj: number, bitmap: number): number
  _FPDFImageObj_GetRenderedBitmap(doc: number, page: number, obj: number): number
  _FPDFBitmap_CreateEx(w: number, h: number, format: number, buf: number, stride: number): number
  _FPDFBitmap_FillRect(
    bitmap: number,
    left: number,
    top: number,
    width: number,
    height: number,
    color: number,
  ): void
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
  _FPDF_GetPageWidthF(page: number): number
  _FPDF_GetPageHeightF(page: number): number
  _FPDFBitmap_Destroy(bitmap: number): void
  _FPDFBitmap_GetBuffer(bitmap: number): number
  _FPDFBitmap_GetWidth(bitmap: number): number
  _FPDFBitmap_GetHeight(bitmap: number): number
  _FPDFBitmap_GetStride(bitmap: number): number
  // Annotation access (annot-delete.ts); EPDF_* are embedpdf extensions
  _FPDFPage_GetAnnotCount(page: number): number
  _FPDFPage_GetAnnot(page: number, index: number): number
  _FPDFPage_CloseAnnot(annot: number): void
  _FPDFPage_RemoveAnnot(page: number, index: number): number
  _FPDFAnnot_GetSubtype(annot: number): number
  _FPDFAnnot_GetRect(annot: number, rect: number): number
  _FPDFAnnot_GetStringValue(annot: number, key: number, buffer: number, buflen: number): number
  _EPDFPage_GetAnnotByObjectNumber(page: number, objNum: number): number
  _EPDFPage_RemoveAnnotByObjectNumber(page: number, objNum: number): number
}

let pdfiumPromise: Promise<Pdfium> | null = null

/** Load the wasm bytes ourselves: the bundled main process must not rely on
    emscripten's __dirname-based lookup (shell bundles this file with no runtime deps) */
export function loadPdfium(): Promise<Pdfium> {
  pdfiumPromise ??= (async () => {
    const { init } = (await import('@embedpdf/pdfium')) as unknown as {
      init(overrides: object): Promise<{ pdfium: Pdfium } | Pdfium>
    }
    const raw = readFileSync(pdfiumWasmPath())
    // Exact slice: Buffer.buffer may be a shared pool larger than the file
    const wasmBinary = raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength)
    // thisProgram: emscripten synthesizes an environ whose `_` entry defaults
    // to process.argv[1] and writes it through ASCII-asserting stringToAscii,
    // so a non-ASCII path in argv — a CJK checkout running vitest, or a
    // document path handed to the packaged app by a file association —
    // aborts the runtime at first use. Every other synthetic entry is a
    // hardcoded ASCII constant; a fixed ASCII program name defuses the
    // only variable one.
    const wrapped = await init({ wasmBinary, thisProgram: 'genoffice-pdf' })
    const m = ('pdfium' in wrapped ? wrapped.pdfium : wrapped) as Pdfium
    m._PDFiumExt_Init()
    return m
  })()
  return pdfiumPromise
}

/** Fallback fonts for rebuilt runs, first readable file wins; must be single-face sfnt (no .ttc) */
const FALLBACK_FONTS = [
  '/System/Library/Fonts/Supplemental/Arial Unicode.ttf',
  'C:\\Windows\\Fonts\\arialuni.ttf',
  '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
]

let fallbackFontBytes: Buffer | null | undefined

function loadFallbackFont(): Buffer {
  if (fallbackFontBytes === undefined) {
    fallbackFontBytes = null
    for (const p of FALLBACK_FONTS) {
      try {
        fallbackFontBytes = readFileSync(p)
        break
      } catch {
        /* try next */
      }
    }
  }
  if (!fallbackFontBytes) throw new Error('no fallback font available for text editing')
  return fallbackFontBytes
}

export type EditFontStyle = 'regular' | 'bold' | 'italic' | 'bolditalic'

/** Font files for the user-selectable rebuild fonts by style, first readable wins
    (single-face sfnt only) */
const EDIT_FONT_PATHS: Record<string, Record<EditFontStyle, string[]>> = {
  arial: {
    regular: [
      '/System/Library/Fonts/Supplemental/Arial.ttf',
      'C:\\Windows\\Fonts\\arial.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
    ],
    bold: [
      '/System/Library/Fonts/Supplemental/Arial Bold.ttf',
      'C:\\Windows\\Fonts\\arialbd.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf',
    ],
    italic: [
      '/System/Library/Fonts/Supplemental/Arial Italic.ttf',
      'C:\\Windows\\Fonts\\ariali.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf',
    ],
    bolditalic: [
      '/System/Library/Fonts/Supplemental/Arial Bold Italic.ttf',
      'C:\\Windows\\Fonts\\arialbi.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf',
    ],
  },
  times: {
    regular: [
      '/System/Library/Fonts/Supplemental/Times New Roman.ttf',
      'C:\\Windows\\Fonts\\times.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf',
    ],
    bold: [
      '/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf',
      'C:\\Windows\\Fonts\\timesbd.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf',
    ],
    italic: [
      '/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf',
      'C:\\Windows\\Fonts\\timesi.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf',
    ],
    bolditalic: [
      '/System/Library/Fonts/Supplemental/Times New Roman Bold Italic.ttf',
      'C:\\Windows\\Fonts\\timesbi.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationSerif-BoldItalic.ttf',
    ],
  },
  courier: {
    regular: [
      '/System/Library/Fonts/Supplemental/Courier New.ttf',
      'C:\\Windows\\Fonts\\cour.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf',
    ],
    bold: [
      '/System/Library/Fonts/Supplemental/Courier New Bold.ttf',
      'C:\\Windows\\Fonts\\courbd.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationMono-Bold.ttf',
    ],
    italic: [
      '/System/Library/Fonts/Supplemental/Courier New Italic.ttf',
      'C:\\Windows\\Fonts\\couri.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationMono-Italic.ttf',
    ],
    bolditalic: [
      '/System/Library/Fonts/Supplemental/Courier New Bold Italic.ttf',
      'C:\\Windows\\Fonts\\courbi.ttf',
      '/usr/share/fonts/truetype/liberation/LiberationMono-BoldItalic.ttf',
    ],
  },
}

const editFontCache = new Map<string, Buffer | null>()

function loadEditFont(id: string, style: EditFontStyle = 'regular'): Buffer | null {
  const key = `${id}#${style}`
  let cached = editFontCache.get(key)
  if (cached === undefined) {
    cached = null
    for (const p of EDIT_FONT_PATHS[id]?.[style] ?? []) {
      try {
        cached = readFileSync(p)
        break
      } catch {
        /* try next */
      }
    }
    editFontCache.set(key, cached)
  }
  return cached
}

/** Edit-font ids that resolve to a readable font file on this machine */
export function listEditFonts(): string[] {
  return Object.keys(EDIT_FONT_PATHS).filter((id) => loadEditFont(id) !== null)
}

/** Whitespace-insensitive, radical- and NFKC-folded comparison key. pdf.js and pdfium
    disagree on compatibility codepoints — some fonts' cmaps yield radical-block
    codepoints (Kangxi U+2F00, which NFKC decomposes, and the Radicals Supplement
    U+2E80, which it does not) where pdfium extracts unified ideographs — and on
    synthesized spaces; both engines' text must land on the same key or every
    whole-line match fails on such documents. */
export const norm = (s: string) => foldRadicals(s).normalize('NFKC').replace(/\s+/g, '')

/** NFKC-fold `raw` keeping maps from each folded unit back to its source char's start
    and end index. Whitespace runs collapse to one ' ' unit when keepSpaces, else
    disappear. */
export function foldMap(
  raw: string,
  keepSpaces = false,
): { units: string[]; idx: number[]; end: number[] } {
  const units: string[] = []
  const idx: number[] = []
  const end: number[] = []
  let i = 0
  let inWs = false
  for (const ch of raw) {
    if (/\s/.test(ch)) {
      if (keepSpaces && !inWs) {
        units.push(' ')
        idx.push(i)
        end.push(i + ch.length)
      }
      inWs = true
    } else {
      inWs = false
      for (const u of foldRadicals(ch).normalize('NFKC')) {
        if (/\s/.test(u)) continue
        units.push(u)
        idx.push(i)
        end.push(i + ch.length)
      }
    }
    i += ch.length
  }
  idx.push(raw.length)
  return { units, idx, end }
}

/** Splice the user's replacement into the engine's own text: the untouched prefix and
    suffix keep the engine's original codepoints (writing back pdf.js-extracted variants
    would silently transliterate unedited text), only the changed middle comes from
    newText as typed. Boundary whitespace retained on both sides follows the engine —
    a renderer-synthesized gap space must not become a real space in the rebuilt run.
    Precondition: norm(engineRaw) === norm(oldText). */
export function spliceIntoEngine(engineRaw: string, oldText: string, newText: string): string {
  const oldK = foldMap(oldText, true)
  const newK = foldMap(newText, true)
  const eng = foldMap(engineRaw)
  // Common prefix/suffix of old vs new in space-keeping folded units
  let p = 0
  const maxP = Math.min(oldK.units.length, newK.units.length)
  while (p < maxP && oldK.units[p] === newK.units[p]) p++
  let s = 0
  const maxS = maxP - p
  while (
    s < maxS &&
    oldK.units[oldK.units.length - 1 - s] === newK.units[newK.units.length - 1 - s]
  )
    s++
  // Multi-line replacements (block reflow) must not take the plain prefix/suffix cut:
  // the retained regions would come from the engine join, which has no '\n's and no
  // separator chars between joined line objects, so the reflow's line breaks (and
  // Latin between-line spaces) would vanish and rebuildRun would draw one overflowing
  // line. Rebuild line by line instead.
  if (newText.includes('\n')) return spliceLines(engineRaw, eng, newK, p, s, newText)
  const nonSpace = (units: string[]) => units.filter((u) => u !== ' ').length
  // Map retained counts onto the engine's space-free fold; back off while a cut would
  // split a raw char whose fold expanded to several units (ligatures)
  for (;;) {
    const np = nonSpace(oldK.units.slice(0, p))
    const ns = nonSpace(oldK.units.slice(oldK.units.length - s))
    const splitsEng =
      (np > 0 && eng.idx[np] === eng.idx[np - 1]) ||
      (ns > 0 && eng.idx[eng.units.length - ns] === eng.idx[eng.units.length - ns - 1])
    const splitsNew =
      (p > 0 && newK.idx[p] === newK.idx[p - 1]) ||
      (s > 0 && newK.idx[newK.units.length - s] === newK.idx[newK.units.length - s - 1])
    if (!splitsEng && !splitsNew) {
      // Retained boundary space (present in old and new alike): cut past it on the
      // engine side so the engine's own spacing survives, and start the typed middle
      // after new's copy of it. A non-space boundary cuts right at the char edge.
      const prefixEndsSpace = p > 0 && oldK.units[p - 1] === ' '
      const suffixStartsSpace = s > 0 && oldK.units[oldK.units.length - s] === ' '
      const engPrefixEnd = p === 0 ? 0 : prefixEndsSpace ? eng.idx[np]! : eng.end[np - 1]!
      const engSuffixStart =
        s === 0
          ? engineRaw.length
          : suffixStartsSpace || ns === 0
            ? eng.end[eng.units.length - ns - 1]!
            : eng.idx[eng.units.length - ns]!
      return (
        engineRaw.slice(0, engPrefixEnd) +
        newText.slice(newK.idx[p]!, newK.idx[newK.units.length - s]!) +
        engineRaw.slice(engSuffixStart)
      )
    }
    if (p > 0) p--
    else if (s > 0) s--
    else return newText
  }
}

/** Per-line splice for multi-line replacements: each line of newText whose folded
    units all sit inside the retained prefix/suffix is rebuilt from the engine's own
    codepoints, every other line stays as typed. `p`/`s` are the unit counts of the
    old↔new common prefix/suffix from spliceIntoEngine. */
function spliceLines(
  engineRaw: string,
  eng: ReturnType<typeof foldMap>,
  newK: ReturnType<typeof foldMap>,
  p: number,
  s: number,
  newText: string,
): string {
  // The engine fold keeps no spaces, so every unit is non-space and the engine's
  // non-space unit count matches oldText's (norm-equality precondition)
  const E = eng.units.length
  const nsPos: number[] = [] // newK unit index of each non-space unit, by rank
  for (let k = 0; k < newK.units.length; k++) if (newK.units[k] !== ' ') nsPos.push(k)
  const N = nsPos.length
  const out: string[] = []
  let k = 0 // walking unit cursor in newK
  let rank = 0 // non-space units consumed so far
  let off = 0 // raw offset of the current line in newText
  for (const line of newText.split('\n')) {
    const lineEnd = off + line.length
    // Skip the collapsed whitespace unit of the preceding '\n' (attributed to neither line)
    while (k < newK.units.length && newK.idx[k]! < off) {
      if (newK.units[k] !== ' ') rank++
      k++
    }
    const kStart = k
    const g1 = rank
    while (k < newK.units.length && newK.idx[k]! < lineEnd) {
      if (newK.units[k] !== ' ') rank++
      k++
    }
    const g2 = rank
    off = lineEnd + 1
    const inPrefix = k <= p
    const inSuffix = kStart >= newK.units.length - s
    if (g2 === g1 || (!inPrefix && !inSuffix)) {
      out.push(line)
      continue
    }
    // Prefix units map 1:1 onto engine ranks; suffix units sit `E - N` ranks later
    const shift = inPrefix ? 0 : E - N
    out.push(engineLineText(engineRaw, eng, nsPos, g1 + shift, g2 + shift, g1) ?? line)
  }
  return out.join('\n')
}

/** Engine-raw text for engine unit ranks [e1, e2); null when the line boundary would
    split a raw char whose fold expanded to several units (ligatures). `g1` is the
    new-text rank of e1, used to restore spaces the engine's line join ran together. */
function engineLineText(
  engineRaw: string,
  eng: ReturnType<typeof foldMap>,
  nsPos: number[],
  e1: number,
  e2: number,
  g1: number,
): string | null {
  if (e1 > 0 && eng.idx[e1]! < eng.end[e1 - 1]!) return null
  if (e2 < eng.units.length && eng.idx[e2]! < eng.end[e2 - 1]!) return null
  let out = ''
  for (let e = e1; e < e2; e++) {
    if (e > e1) {
      if (eng.idx[e]! < eng.end[e - 1]!) continue // same raw char (fold expansion)
      const gap = engineRaw.slice(eng.end[e - 1]!, eng.idx[e]!)
      out += gap
      // The new text separates these units but the engine join put nothing between
      // them (adjacent line objects): a reflow moved a word across the old break, so
      // restore the space the visual line break used to stand for
      const g = g1 + (e - e1)
      if (!/\s/.test(gap) && nsPos[g]! - nsPos[g - 1]! > 1) out += ' '
    }
    out += engineRaw.slice(eng.idx[e]!, eng.end[e]!)
  }
  return out
}

/** Rebuild a paragraph replacement keeping the engine's own codepoints for the
    unchanged head and tail. spliceIntoEngine keeps the engine's raw text (per-line
    for multi-line replacements) — right for single-run edits, wrong for paragraphs:
    engine text joined across a block's objects carries no spaces or breaks at line
    seams, so an unchanged head/tail would collapse the reflowed lines and glue seam
    words together. Here newText's whitespace and '\n' structure survive in full;
    only the non-space codepoints of the unchanged units are swapped back to the
    engine's (Kangxi-radical extraction variants must not be written back).
    Precondition: norm(engineRaw) === norm(oldText). */
export function mergeEngineCodepoints(engineRaw: string, oldText: string, newText: string): string {
  const oldU = foldMap(oldText)
  const newU = foldMap(newText)
  const eng = foldMap(engineRaw)
  if (eng.units.length !== oldU.units.length) return newText
  let p = 0
  const maxP = Math.min(oldU.units.length, newU.units.length)
  while (p < maxP && oldU.units[p] === newU.units[p]) p++
  let s = 0
  const maxS = maxP - p
  while (
    s < maxS &&
    oldU.units[oldU.units.length - 1 - s] === newU.units[newU.units.length - 1 - s]
  )
    s++
  // Back off boundaries that would split a raw char whose fold expanded to several
  // units (ligatures): emitting the engine's whole glyph for a partially-retained
  // unit span would keep an extra letter or drop a typed one
  const splits = (fm: ReturnType<typeof foldMap>, i: number) =>
    i > 0 && i < fm.units.length && fm.idx[i]! < fm.end[i - 1]!
  while (p > 0 && (splits(eng, p) || splits(newU, p))) p--
  while (s > 0 && (splits(eng, eng.units.length - s) || splits(newU, newU.units.length - s))) s--
  const total = newU.units.length
  const engShift = eng.units.length - total
  let out = ''
  let k = 0
  let engEnd = 0
  for (const ch of newText) {
    if (/\s/.test(ch)) {
      out += ch
      continue
    }
    const m = [...ch.normalize('NFKC')].filter((u) => !/\s/.test(u)).length
    if (m === 0) {
      out += ch
      continue
    }
    const inPrefix = k + m <= p
    const inSuffix = k >= total - s
    if (inPrefix || inSuffix) {
      const at = inPrefix ? k : k + engShift
      const rawStart = eng.idx[at]!
      const rawEnd = eng.end[at + m - 1]!
      // Overlap guard: an engine ligature spans several units; emit its raw char once
      out += engineRaw.slice(Math.max(rawStart, engEnd), rawEnd)
      engEnd = Math.max(engEnd, rawEnd)
    } else {
      out += ch
    }
    k += m
  }
  return out
}

interface PageTextObj {
  obj: number
  index: number
  text: string
  font: number
  /** [x1, y1, x2, y2] */
  bounds: [number, number, number, number]
}

function collectTextObjects(m: Pdfium, page: number, textPage: number): PageTextObj[] {
  const out: PageTextObj[] = []
  const bl = m._malloc(4)
  const bb = m._malloc(4)
  const br = m._malloc(4)
  const bt = m._malloc(4)
  const count = m._FPDFPage_CountObjects(page)
  for (let i = 0; i < count; i++) {
    const obj = m._FPDFPage_GetObject(page, i)
    if (m._FPDFPageObj_GetType(obj) !== FPDF_PAGEOBJ_TEXT) continue
    // Two-pass read: FPDFTextObj_GetText's length argument and return value are BYTES
    // including the 2-byte NUL terminator, and a too-small buffer is left untouched
    // (not truncated) — a fixed buffer would silently yield garbage for long runs
    const len = m._FPDFTextObj_GetText(obj, textPage, 0, 0)
    let text = ''
    if (len > 2) {
      const buf = m._malloc(len)
      m._FPDFTextObj_GetText(obj, textPage, buf, len)
      text = Buffer.from(m.HEAPU8.buffer, buf, len - 2).toString('utf16le')
      m._free(buf)
    }
    if (!m._FPDFPageObj_GetBounds(obj, bl, bb, br, bt)) continue
    out.push({
      obj,
      index: i,
      text,
      font: m._FPDFTextObj_GetFont(obj),
      bounds: [m.HEAPF32[bl >> 2]!, m.HEAPF32[bb >> 2]!, m.HEAPF32[br >> 2]!, m.HEAPF32[bt >> 2]!],
    })
  }
  for (const p of [bl, bb, br, bt]) m._free(p)
  return out
}

type Rect = readonly [number, number, number, number]

function overlapArea(a: Rect, b: Rect): number {
  const w = Math.min(a[2], b[2]) - Math.max(a[0], b[0])
  const h = Math.min(a[3], b[3]) - Math.max(a[1], b[1])
  return w > 0 && h > 0 ? w * h : 0
}

/** Overlap area ÷ object area; edits address whole objects, so partial bleed from
    neighboring lines must not capture them */
const overlapRatio = (o: Rect, r: Rect) =>
  overlapArea(o, r) / Math.max((o[2] - o[0]) * (o[3] - o[1]), 1e-6)

/** Overlap area ÷ edit-rect area: how much of the edited span sits inside the object */
const rectCoverage = (o: Rect, r: Rect) =>
  overlapArea(o, r) / Math.max((r[2] - r[0]) * (r[3] - r[1]), 1e-6)

function utf16Ptr(m: Pdfium, str: string): number {
  const bytes = Buffer.from(`${str}\0`, 'utf16le')
  const p = m._malloc(bytes.length)
  m.HEAPU8.set(bytes, p)
  return p
}

/** SetText re-encodes against the object's existing font. Only safe when the replacement
    is ASCII and every char is provably in the embedded subset (i.e. already drawn somewhere
    with the same font). CJK subset CID fonts fail even for in-subset chars (broken cmaps),
    so anything non-ASCII always goes the rebuild path. Style overrides, alignment offsets
    and multi-line replacements also rebuild: they need new objects (or a new matrix), not
    just re-encoded text. */
function canReuseFont(
  edit: TextEditInput,
  newText: string,
  matches: PageTextObj[],
  all: PageTextObj[],
): boolean {
  if (matches.length !== 1) return false
  if (edit.newFontSize !== undefined || edit.newColor !== undefined || edit.newFont !== undefined)
    return false
  // Selection-level styles split the line into several objects: always rebuild
  const runs = editStyleRuns(edit)
  if (runs && runs.length > 0) return false
  // Centered/right blocks reposition every line: SetText keeps the object's own
  // matrix, so a shorter replacement would stay at the old x instead of re-centering
  if (edit.lineXOffsets !== undefined) return false
  // Bold/italic toggles swap the face: the object's existing font cannot draw them
  if (edit.newBold || edit.newItalic) return false
  if (!/^[\x20-\x7e]*$/.test(newText)) return false
  const font = matches[0]!.font
  const charset = new Set<string>()
  for (const o of all) if (o.font === font) for (const ch of o.text) charset.add(ch)
  return [...newText].every((ch) => ch === ' ' || charset.has(ch))
}

/** The text objects an edit resolves to, plus the full replacement for them
    (fragment edits expand to the container object's text with the fragment spliced in) */
interface PlannedMatch {
  matches: PageTextObj[]
  newText: string
  /** True when the matched objects ARE the edited text (not a fragment of a larger
      run); only then may position overrides (origin/lineLeading) be honored */
  whole: boolean
}

/** Split same-baseline objects into z-stacked layers: an earlier multi-line edit that
    overflowed onto the next line leaves two runs drawn over each other, and a whole-line
    rect then captures both. Objects joining one visual line never overlap horizontally,
    so chaining by x-extent recovers the individual layers. */
function xLayers(objs: PageTextObj[]): PageTextObj[][] {
  const sorted = [...objs].sort((a, b) => a.bounds[0] - b.bounds[0])
  const heights = sorted.map((o) => o.bounds[3] - o.bounds[1]).sort((a, b) => a - b)
  const tieMargin = (heights[heights.length >> 1] ?? 0) * 0.5
  const assign = chainLayers(
    sorted.map((o) => ({ left: o.bounds[0], right: o.bounds[2] })),
    tieMargin,
  )
  const layers: PageTextObj[][] = []
  sorted.forEach((o, i) => {
    ;(layers[assign[i]!] ??= []).push(o)
  })
  return layers
}

/** Subarray search over folded units (indexOf on joined strings would return UTF-16
    offsets, which drift from unit offsets on astral codepoints) */
function indexOfUnits(hay: string[], needle: string[]): number {
  if (needle.length === 0) return -1
  outer: for (let i = 0; i + needle.length <= hay.length; i++) {
    for (let j = 0; j < needle.length; j++) if (hay[i + j] !== needle[j]) continue outer
    return i
  }
  return -1
}

/** Joined text of a candidate object set, x-order (how the renderer reads a line) */
const joinX = (objs: PageTextObj[]) =>
  [...objs]
    .sort((a, b) => a.bounds[0] - b.bounds[0])
    .map((o) => o.text)
    .join('')

/** Objects in visual reading order — rows top→bottom (≥50% vertical overlap joins
    a row), x within each row. How a multi-line paragraph edit reads its lines; plain
    x-order interleaves them and stream order is not guaranteed. */
function readingOrder(objs: PageTextObj[]): PageTextObj[] {
  const sorted = [...objs].sort((a, b) => b.bounds[3] - a.bounds[3])
  const rows: PageTextObj[][] = []
  for (const o of sorted) {
    const row = rows[rows.length - 1]
    const first = row?.[0]
    const overlap = first
      ? Math.min(first.bounds[3], o.bounds[3]) - Math.max(first.bounds[1], o.bounds[1])
      : 0
    if (
      first &&
      overlap >= 0.5 * Math.min(first.bounds[3] - first.bounds[1], o.bounds[3] - o.bounds[1])
    ) {
      row.push(o)
    } else rows.push([o])
  }
  return rows.flatMap((row) => [...row].sort((a, b) => a.bounds[0] - b.bounds[0]))
}

const joinRows = (objs: PageTextObj[]): string =>
  readingOrder(objs)
    .map((o) => o.text)
    .join('')

/**
 * Text-first rescue: the rect is only a hint here, not a filter. The renderer's rect
 * comes from pdf.js layout boxes, which can sit a few points inside the engine's ink
 * bounds — on a line built from many text objects the edge objects then fail the
 * primary path's ≥50%-containment test and the whole-line join never equals oldText.
 * Instead, order every object merely touching a slightly padded rect by reading order,
 * search their joined folded units for oldText's units, and pick the occurrence whose
 * union bounds best overlap the edit rect (disambiguates repeated text; bleed from
 * neighboring lines falls outside the occurrence and is skipped naturally). The match
 * may start or end mid-object — the uncovered head/tail of the edge objects is kept
 * verbatim around the replacement, like the single-container fragment path.
 */
function matchByText(objects: PageTextObj[], edit: TextEditInput): PlannedMatch | null {
  const target = foldMap(edit.oldText)
  if (target.units.length === 0) return null
  const pad = Math.max(2, edit.fontSize * 0.2)
  const r: Rect = [edit.rect[0] - pad, edit.rect[1] - pad, edit.rect[2] + pad, edit.rect[3] + pad]
  const touching = readingOrder(objects.filter((o) => overlapArea(o.bounds, r) > 0))
  if (touching.length === 0) return null
  const joined = touching.map((o) => o.text).join('')
  const eng = foldMap(joined)
  // Which object each folded unit came from, plus each object's raw span in `joined`
  // (per-object folds concatenate to the joined fold: folding is per-character)
  const objAt: number[] = []
  const rawStartOf: number[] = []
  const rawEndOf: number[] = []
  let rawOff = 0
  for (const [i, o] of touching.entries()) {
    rawStartOf.push(rawOff)
    rawOff += o.text.length
    rawEndOf.push(rawOff)
    for (let k = foldMap(o.text).units.length; k > 0; k--) objAt.push(i)
  }
  let best: { at: number; score: number } | null = null
  let at = indexOfUnits(eng.units, target.units)
  while (at >= 0) {
    const endU = at + target.units.length - 1
    const set = touching.slice(objAt[at]!, objAt[endU]! + 1)
    const b = matchBounds(set)
    const score = overlapArea(b, edit.rect) / Math.max((b[2] - b[0]) * (b[3] - b[1]), 1e-6)
    if (score > 0 && (!best || score > best.score)) best = { at, score }
    const next = indexOfUnits(eng.units.slice(at + 1), target.units)
    at = next < 0 ? -1 : at + 1 + next
  }
  if (!best) return null
  const endU = best.at + target.units.length - 1
  const set = touching.slice(objAt[best.at]!, objAt[endU]! + 1)
  // Raw text the matched objects contribute, and the matched stretch inside it
  const setStart = rawStartOf[objAt[best.at]!]!
  const setEnd = rawEndOf[objAt[endU]!]!
  const setRaw = joined.slice(setStart, setEnd)
  const rawStart = eng.idx[best.at]!
  // End of the last matched unit, not the next unit's start: the gap between them is
  // boundary whitespace belonging to the untouched suffix (see the container path)
  const rawEnd = eng.end[endU]!
  const whole = norm(setRaw) === norm(edit.oldText)
  if (whole) {
    return {
      matches: set,
      newText:
        edit.newText === ''
          ? ''
          : edit.lineLeading !== undefined
            ? mergeEngineCodepoints(setRaw, edit.oldText, edit.newText)
            : spliceIntoEngine(setRaw, edit.oldText, edit.newText),
      whole: true,
    }
  }
  // Paragraph rebuilds position their lines at the block corner (origin/lineLeading);
  // apply strips those overrides from fragment matches, which would rebuild the block
  // at the anchor and drag the edge objects' surrounding text into the block run.
  // Fail closed like before instead of degrading the layout silently.
  if (edit.origin !== undefined || edit.lineLeading !== undefined) return null
  const fragment = spliceIntoEngine(joined.slice(rawStart, rawEnd), edit.oldText, edit.newText)
  return {
    matches: set,
    newText: joined.slice(setStart, rawStart) + fragment + joined.slice(rawEnd, setEnd),
    whole: false,
  }
}

/**
 * Resolve an edit to concrete text objects. Primary rule: objects sitting mostly inside
 * the edit's rect whose joined text equals oldText — tried in content-stream order, then
 * x-order, then per x-layer (stacked runs from an earlier overflowing edit). Fallback for
 * granularity mismatches (pdf.js splits one PDF text object into several spans at column
 * gaps / kerning breaks): a single object containing most of the rect gets the fragment
 * spliced into its text — the whole object is then rewritten, so surrounding text
 * survives with the new content. Last resort is matchByText: a text-first search over
 * objects merely touching the rect, rescuing rects that clip edge objects below the
 * containment threshold. All comparisons are NFKC-folded (norm) and the replacement
 * is spliced against the engine's own text to keep unedited codepoints.
 */
function matchEdit(objects: PageTextObj[], edit: TextEditInput): PlannedMatch | { reason: string } {
  // A replacement with no printable characters deletes the matched run: normalize
  // whitespace-only forms to '' so the splice helpers (which treat '\n' structurally)
  // never see them and the apply path removes the objects without drawing anything back
  if (edit.newText.trim() === '' && edit.newText !== '') edit = { ...edit, newText: '' }
  const oldKey = norm(edit.oldText)
  const matches = objects.filter((o) => overlapRatio(o.bounds, edit.rect) >= 0.5)
  const candidates: PageTextObj[][] = [matches]
  if (matches.length > 1) candidates.push(...xLayers(matches))
  for (const set of candidates) {
    if (set.length === 0) continue
    const streamJoin = set.map((o) => o.text).join('')
    for (const joined of [streamJoin, joinX(set), joinRows(set)]) {
      if (norm(joined) === oldKey) {
        return {
          matches: set,
          // Paragraph rebuilds keep newText's line/space structure (engine text has
          // none at line seams); single-run edits keep the engine's raw text
          newText:
            edit.newText === ''
              ? ''
              : edit.lineLeading !== undefined
                ? mergeEngineCodepoints(joined, edit.oldText, edit.newText)
                : spliceIntoEngine(joined, edit.oldText, edit.newText),
          whole: true,
        }
      }
    }
  }
  const container = objects
    .map((o) => ({ o, cover: rectCoverage(o.bounds, edit.rect) }))
    .filter((c) => c.cover >= 0.5)
    .sort((a, b) => b.cover - a.cover)[0]?.o
  if (container) {
    const raw = container.text
    const eng = foldMap(raw)
    const target = foldMap(edit.oldText)
    const at = indexOfUnits(eng.units, target.units)
    if (at >= 0) {
      const rawStart = eng.idx[at]!
      // Slice to the END of the last matched unit, not the next unit's start: the gap
      // between them is boundary whitespace that belongs to the untouched suffix. Folding
      // it into the fragment let a style-only edit (oldText === newText) eat the space —
      // spliceIntoEngine's retained suffix is empty then, so the gap survived nowhere.
      const rawEnd = eng.end[at + target.units.length - 1]!
      const fragment = spliceIntoEngine(raw.slice(rawStart, rawEnd), edit.oldText, edit.newText)
      return {
        matches: [container],
        newText: raw.slice(0, rawStart) + fragment + raw.slice(rawEnd),
        whole: false,
      }
    }
  }
  const rescued = matchByText(objects, edit)
  if (rescued) return rescued
  return { reason: 'the edited text could not be located on the page' }
}

const readHeapU32 = (m: Pdfium, ptr: number) =>
  m.HEAPU8[ptr]! +
  m.HEAPU8[ptr + 1]! * 0x100 +
  m.HEAPU8[ptr + 2]! * 0x10000 +
  m.HEAPU8[ptr + 3]! * 0x1000000

/** Decoded font-program bytes of an embedded font, null when not embedded/readable */
function embeddedFontData(m: Pdfium, font: number): Buffer | null {
  if (!font || !m._FPDFFont_GetIsEmbedded(font)) return null
  const lenPtr = m._malloc(4)
  try {
    if (!m._FPDFFont_GetFontData(font, 0, 0, lenPtr)) return null
    const len = readHeapU32(m, lenPtr)
    if (!len) return null
    const buf = m._malloc(len)
    try {
      if (!m._FPDFFont_GetFontData(font, buf, len, lenPtr)) return null
      return Buffer.from(m.HEAPU8.subarray(buf, buf + len))
    } finally {
      m._free(buf)
    }
  } finally {
    m._free(lenPtr)
  }
}

/** UTF-8 string via a two-pass FPDFFont_Get*Name call (returned length includes the NUL) */
function fontString(m: Pdfium, read: (buf: number, len: number) => number): string {
  const len = read(0, 0)
  if (len <= 1) return ''
  const buf = m._malloc(len)
  try {
    read(buf, len)
    return Buffer.from(m.HEAPU8.subarray(buf, buf + len - 1)).toString()
  } finally {
    m._free(buf)
  }
}

/** Font bytes for a rebuilt run, best fidelity first. Explicit choice: that face when its
    cmap covers every replacement char (chars it can't map would subset into .notdef boxes).
    No choice ("keep original"): the run's own embedded font when its subset already holds
    every glyph, else the installed font with the same PostScript/family name — a subset
    physically lacks glyphs the document never drew, so new chars need the full face
    (Acrobat resolves the same way). The fallback face otherwise. */
async function rebuildFontBytes(
  m: Pdfium,
  font: number,
  edit: TextEditInput,
  newText: string,
): Promise<Buffer> {
  const drawn = newText.replace(/\n/g, '')
  const style: EditFontStyle =
    edit.newBold && edit.newItalic
      ? 'bolditalic'
      : edit.newBold
        ? 'bold'
        : edit.newItalic
          ? 'italic'
          : 'regular'
  if (edit.newFont) {
    // Style variant first, base face as the degrade (never fail the edit over style)
    for (const s of style === 'regular' ? (['regular'] as const) : ([style, 'regular'] as const)) {
      const chosen = loadEditFont(edit.newFont, s)
      if (chosen && fontCoversText(chosen, drawn)) return subsetTtf(chosen, drawn)
    }
  } else if (font) {
    if (style !== 'regular') {
      // Style override on "keep original": the embedded subset is the base face, so
      // ask the installed index for the family's matching variant — appending the
      // style word to the PS name merges with any tokens it already carries
      const ps = fontString(m, (b, l) => m._FPDFFont_GetBaseFontName(font, b, l)).replace(
        /^[A-Z]{6}\+/,
        '',
      )
      const family = fontString(m, (b, l) => m._FPDFFont_GetFamilyName(font, b, l))
      if (ps || family) {
        // The combined query merges inherited and requested tokens; when the combined
        // face is missing they tie, and the ranker may hand back the original face —
        // a silent no-op for the toggle. Detect that (byte-equal to the plain-PS
        // lookup) and retry with the requested style alone, so the user's latest
        // toggle wins over the run's inherited style.
        let sys = findSystemFont(`${ps}-${style}`, family)
        const original = findSystemFont(ps, family)
        // A run whose face already carries the requested style makes the combined
        // lookup legitimately return that face — that is a hit, not a no-op
        const psTokens = (ps.toLowerCase().match(/bolditalic|bold|italic|oblique/g) ?? []).flatMap(
          (t) => (t === 'bolditalic' ? ['bold', 'italic'] : t === 'oblique' ? ['italic'] : [t]),
        )
        const wanted = style === 'bolditalic' ? ['bold', 'italic'] : [style]
        const alreadyStyled = wanted.every((t) => psTokens.includes(t))
        if (!alreadyStyled && sys && original && sys.equals(original)) {
          const stripped = ps.replace(/bolditalic|bold|italic|oblique/gi, '')
          const alt = findSystemFont(`${stripped}-${style}`, family)
          sys = alt && !alt.equals(original) ? alt : null
        }
        if (sys && fontCoversText(sys, drawn)) {
          try {
            return identityCffCharset(await subsetTtf(sys, drawn))
          } catch {
            /* charset not rewritable: degrade to the base face below */
          }
        }
      }
    }
    const embedded = embeddedFontData(m, font)
    // Already a subset: re-subsetting a GID-remapped face buys nothing and risks breakage.
    // Bare-CFF font programs (FontFile3 without an sfnt wrapper) have no cmap and fail
    // the coverage check, falling through to the name lookup.
    if (embedded && fontCoversText(embedded, drawn)) {
      try {
        return identityCffCharset(embedded)
      } catch {
        /* charset not rewritable: try the installed face instead */
      }
    }
    const ps = fontString(m, (b, l) => m._FPDFFont_GetBaseFontName(font, b, l)).replace(
      /^[A-Z]{6}\+/,
      '',
    )
    const family = fontString(m, (b, l) => m._FPDFFont_GetFamilyName(font, b, l))
    if (ps || family) {
      const sys = findSystemFont(ps, family)
      if (sys && fontCoversText(sys, drawn)) {
        try {
          return identityCffCharset(await subsetTtf(sys, drawn))
        } catch {
          /* charset not rewritable: fall back */
        }
      }
    }
  }
  const fallback = loadFallbackFont()
  // Chars beyond even the fallback face (emoji, rare CJK extensions) would embed as
  // missing glyphs and then fail read-back verification, blocking the whole save —
  // reject this one edit instead (it is reported as skipped, everything else saves)
  if (!fontCoversText(fallback, drawn)) {
    throw new Error('the replacement contains characters no available font can draw')
  }
  return subsetTtf(fallback, drawn)
}

/** Extra leading between stacked lines, multiple of the font size (matches the preview CSS) */
const LINE_GAP = 1.2

type Rgb = readonly [number, number, number]

/** Selection-level style of one planned char; every field absent = inherit the
    whole-edit override (newColor / newFont / newFontSize / newBold / newItalic),
    which in turn inherits the original run. */
export interface RunStyle {
  color?: Rgb
  font?: string
  size?: number
  bold?: boolean
  italic?: boolean
}

/** Identity key for segmentation — style objects are compared by value, not reference
    (overlaying preserved colors under user runs builds fresh objects per char) */
const styleKeyOf = (s: RunStyle | null): string =>
  s
    ? [
        s.color ? s.color.join(',') : '',
        s.font ?? '',
        s.size ?? '',
        s.bold === undefined ? '' : s.bold ? 1 : 0,
        s.italic === undefined ? '' : s.italic ? 1 : 0,
      ].join('|')
    : ''

/** True when the style needs its own face or size — such chars can never ride on a
    kept (merely translated) object; only pure fill overrides can. */
const styledBeyondColor = (s: RunStyle | null): boolean =>
  !!s &&
  (s.font !== undefined || s.size !== undefined || s.bold !== undefined || s.italic !== undefined)

/** The edit's selection-level runs: styleRuns, or legacy colorRuns from older senders */
const editStyleRuns = (
  edit: Pick<TextEditInput, 'colorRuns' | 'styleRuns'>,
): NonNullable<TextEditInput['styleRuns']> | undefined => edit.styleRuns ?? edit.colorRuns

/** Per-code-unit segment styles for the engine's planned text, aligned from the user's
    newText through the NFKC fold: the splice swaps unedited codepoints back to engine
    variants and a fragment match wraps container text around the replacement, so raw
    offsets do not transfer. null = no run styling applies (or unalignable — the caller
    degrades to a uniform rebuild). */
export function plannedCharStyles(
  edit: Pick<TextEditInput, 'newText' | 'colorRuns' | 'styleRuns'>,
  plannedText: string,
): (RunStyle | null)[] | null {
  const runs = editStyleRuns(edit)
  if (!runs || runs.length === 0) return null
  const user = foldMap(edit.newText)
  const planned = foldMap(plannedText)
  const off = indexOfUnits(planned.units, user.units)
  if (off < 0) return null
  // One style object per run so same-run chars group by reference too
  const styles = runs.map((r): RunStyle => ({
    color: r.color,
    font: r.font,
    size: r.size,
    bold: r.bold,
    italic: r.italic,
  }))
  const styleAt = (i: number): RunStyle | null => {
    for (const [ri, r] of runs.entries()) if (i >= r.start && i < r.end) return styles[ri]!
    return null
  }
  const out: (RunStyle | null)[] = new Array<RunStyle | null>(plannedText.length).fill(null)
  for (let u = 0; u < user.units.length; u++) {
    const s = styleAt(user.idx[u]!)
    if (!s) continue
    const p = u + off
    for (let k = planned.idx[p]!; k < planned.end[p]!; k++) out[k] = s
  }
  // Whitespace carries no ink and no fold unit: attach it to the preceding char's
  // segment so a styled word keeps its trailing space
  for (let k = 1; k < out.length; k++) {
    if (out[k] === null && plannedText[k] !== '\n' && /\s/.test(plannedText[k]!)) {
      out[k] = out[k - 1]!
    }
  }
  return out.some((s) => s !== null) ? out : null
}

/** Per-code-unit fill colors of `targetText` inherited from the matched objects.
    The rebuild removes and repaints every matched object, so colors an earlier
    edit saved into the run (selection colors = one object per color) must ride
    along or the next edit silently flattens them to the anchor color. Alignment
    is by folded unit rank: the planned text keeps the engine's units for the
    unchanged head and tail, so common prefix/suffix units pair 1:1 with the
    engine's and inherit their object's color; the edited middle stays null (the
    base color, or the user's explicit runs overlaid by the caller). null when
    the matches all share one fill — nothing worth preserving. */
function matchCharColors(
  m: Pdfium,
  matches: PageTextObj[],
  targetText: string,
): (Rgb | null)[] | null {
  if (matches.length < 2) return null
  const colPtr = m._malloc(16)
  const byKey = new Map<string, Rgb>()
  const colorOf = new Map<PageTextObj, Rgb | null>()
  try {
    for (const t of matches) {
      if (!m._FPDFPageObj_GetFillColor(t.obj, colPtr, colPtr + 4, colPtr + 8, colPtr + 12)) {
        colorOf.set(t, null)
        continue
      }
      const rgb = [0, 4, 8].map((off) => m.HEAPU8[colPtr + off]!) as unknown as Rgb
      const key = rgb.join(',')
      // Canonical instance per RGB so segmentLine's reference grouping merges
      // adjacent same-color spans that came from different objects
      const c = byKey.get(key) ?? rgb
      byKey.set(key, c)
      colorOf.set(t, c)
    }
  } finally {
    m._free(colPtr)
  }
  if (byKey.size <= 1) return null
  const engUnits: string[] = []
  const engColors: (Rgb | null)[] = []
  for (const t of readingOrder(matches)) {
    const units = foldMap(t.text).units
    const c = colorOf.get(t) ?? null
    for (const u of units) {
      engUnits.push(u)
      engColors.push(c)
    }
  }
  const tgt = foldMap(targetText)
  let p = 0
  const maxP = Math.min(tgt.units.length, engUnits.length)
  while (p < maxP && tgt.units[p] === engUnits[p]) p++
  let s = 0
  const maxS = maxP - p
  while (s < maxS && tgt.units[tgt.units.length - 1 - s] === engUnits[engUnits.length - 1 - s]) s++
  const shift = engUnits.length - tgt.units.length
  const out: (Rgb | null)[] = new Array<Rgb | null>(targetText.length).fill(null)
  let any = false
  for (let u = 0; u < tgt.units.length; u++) {
    const c = u < p ? engColors[u]! : u >= tgt.units.length - s ? engColors[u + shift]! : null
    if (!c) continue
    for (let k = tgt.idx[u]!; k < tgt.end[u]!; k++) out[k] = c
    any = true
  }
  return any ? out : null
}

/** User selection styles overlaid on the run's own preserved colors (a styled range
    without an explicit color keeps the preserved fill underneath); whitespace joins
    the preceding char's segment (it carries no ink of its own). */
function overlayStyles(
  baseColors: (Rgb | null)[] | null,
  over: (RunStyle | null)[] | null,
  text: string,
): (RunStyle | null)[] | null {
  if (!baseColors && !over) return null
  const out: (RunStyle | null)[] = new Array<RunStyle | null>(text.length).fill(null)
  const colorOnly = new Map<Rgb, RunStyle>()
  for (let k = 0; k < text.length; k++) {
    const b = baseColors?.[k] ?? null
    const o = over?.[k] ?? null
    if (o) {
      out[k] = o.color === undefined && b ? { ...o, color: b } : o
    } else if (b) {
      let s = colorOnly.get(b)
      if (!s) colorOnly.set(b, (s = { color: b }))
      out[k] = s
    }
  }
  for (let k = 1; k < out.length; k++) {
    if (out[k] === null && text[k] !== '\n' && /\s/.test(text[k]!)) out[k] = out[k - 1]!
  }
  return out.some((s) => s !== null) ? out : null
}

interface LineSeg {
  text: string
  style: RunStyle | null
  /** Text-space x of the segment start in PDF pt (each preceding segment measured
      with its own face and size) */
  xPt: number
}

/** Split one rebuilt line at style boundaries, positioning each segment by the advance
    of the text before it. `advancePt` measures one codepoint in text-space pt with the
    face and size that will actually draw it — from the loaded PDFium font, not the raw
    font bytes: embedded subsets routinely strip cmap/hmtx, which made a bytes-parsing
    measurement fail and silently discard the selection styling. A char that still
    cannot be measured degrades to a single base-style segment — a uniform line beats
    guessed positions. */
function segmentLine(
  line: string,
  styles: (RunStyle | null)[] | null,
  advancePt: (cp: number, style: RunStyle | null) => number | null,
): LineSeg[] {
  const whole: LineSeg[] = [{ text: line, style: null, xPt: 0 }]
  if (!styles) return whole
  // Group per codepoint (advances count surrogate pairs once) by per-code-unit style
  const groups: { text: string; style: RunStyle | null; key: string }[] = []
  let cu = 0
  for (const ch of line) {
    const s = styles[cu] ?? null
    const key = styleKeyOf(s)
    const last = groups[groups.length - 1]
    if (last && last.key === key) last.text += ch
    else groups.push({ text: ch, style: s, key })
    cu += ch.length
  }
  if (groups.length <= 1) {
    return groups.length ? groups.map((g) => ({ text: g.text, style: g.style, xPt: 0 })) : whole
  }
  const segs: LineSeg[] = []
  let x = 0
  for (const g of groups) {
    segs.push({ text: g.text, style: g.style, xPt: x })
    for (const ch of g.text) {
      const a = advancePt(ch.codePointAt(0)!, g.style)
      if (a === null) return whole
      x += a
    }
  }
  return segs
}

/** One original text object the rebuild keeps (translated, not redrawn) */
interface KeepPlanObj {
  src: PageTextObj
  /** Code-unit range in newText this object's characters cover */
  startK: number
  endK: number
  /** Advance width in page units (sum of the chars' loose boxes) */
  advW: number
  /** Original baseline origin of the first char (page space) */
  origX: number
  origY: number
  /** Fill override for the whole object; null = keep its own fill */
  color: Rgb | null
}

/** Signal to abandon the object-preserving build and fall back to a full redraw */
class PreserveAbort extends Error {}

/** Detect whitespace edits at the boundaries of the user's change. The keep-plan
    alignment works on space-free folded units, so a space-only edit (deleting the
    gap in "phon e") changes no unit at all: every object would land in the common
    prefix/suffix, be kept verbatim, and be re-placed at its original spacing —
    silently undoing the edit (space chars the objects carry would survive too).
    Diff old/new with spaces kept; when the first (last) divergence sits on a space
    unit, report how many non-space units before (after) it may still be kept — one
    less than the full common run, so the glyph adjacent to the edited seam is
    redrawn, which closes/opens the gap and drops any removed space chars.
    Returns null when the texts are identical or no boundary touches whitespace. */
export function wsEditClamp(
  oldText: string,
  newText: string,
): { headNS: number | null; tailNS: number | null } | null {
  const o = foldMap(oldText, true)
  const n = foldMap(newText, true)
  let pw = 0
  const maxPw = Math.min(o.units.length, n.units.length)
  while (pw < maxPw && o.units[pw] === n.units[pw]) pw++
  if (pw === o.units.length && pw === n.units.length) return null
  let sw = 0
  const maxSw = maxPw - pw
  while (sw < maxSw && o.units[o.units.length - 1 - sw] === n.units[n.units.length - 1 - sw]) sw++
  // Boundary units come from the diverging window only. A pure insertion leaves one
  // side's window empty (pw + sw covers it entirely); indexing that side would read
  // a RETAINED unit — a space in the unchanged prefix/suffix then faked a
  // "whitespace edit" and clamped the seam-adjacent glyph into the redraw, which
  // fails outright when that glyph is undrawable (PUA icon runs).
  const winUnit = (fm: { units: string[] }, i: number): string | undefined =>
    pw < fm.units.length - sw ? fm.units[i] : undefined
  const headWs = winUnit(o, pw) === ' ' || winUnit(n, pw) === ' '
  const tailWs =
    winUnit(o, o.units.length - 1 - sw) === ' ' || winUnit(n, n.units.length - 1 - sw) === ' '
  if (!headWs && !tailWs) return null
  const nonSpace = (units: string[], from: number, to: number) => {
    let c = 0
    for (let i = from; i < to; i++) if (units[i] !== ' ') c++
    return c
  }
  return {
    headNS: headWs ? nonSpace(o.units, 0, pw) - 1 : null,
    tailNS: tailWs ? nonSpace(o.units, o.units.length - sw, o.units.length) - 1 : null,
  }
}

/**
 * Which matched objects can survive the rebuild untouched. A full redraw flattens
 * mixed weights/faces into one located font — unrecoverable for Chrome-print docs
 * whose per-glyph Type3 fonts carry no metadata to even detect bold from. So align
 * the objects against the rebuilt text by folded unit rank and keep every object
 * whose characters all sit in the unchanged head/tail and land on one output line;
 * kept objects are merely translated, only the edited middle is redrawn. Geometry
 * comes from the pre-edit text page (char → object, loose-box advances).
 */
function buildKeepPlan(
  m: Pdfium,
  textPage: number,
  matches: PageTextObj[],
  newText: string,
  charStyles: (RunStyle | null)[] | null,
  newColor: Rgb | undefined,
  edit: Pick<TextEditInput, 'oldText' | 'newText'>,
): KeepPlanObj[] | null {
  const charsOf = new Map<number, number[]>(matches.map((t) => [t.obj, []]))
  const total = m._FPDFText_CountChars(textPage)
  for (let i = 0; i < total; i++) {
    charsOf.get(m._FPDFText_GetTextObject(textPage, i))?.push(i)
  }
  const rectPtr = m._malloc(16)
  const xPtr = m._malloc(8)
  const yPtr = m._malloc(8)
  try {
    const ordered = readingOrder(matches)
    const engUnits: string[] = []
    const unitCount: number[] = []
    const geo: ({ origX: number; origY: number; advW: number } | null)[] = []
    for (const t of ordered) {
      const chars = charsOf.get(t.obj) ?? []
      let g: { origX: number; origY: number; advW: number } | null = null
      // The object's page chars must correspond 1:1 to its extracted codepoints,
      // or advance/origin attribution would silently drift. pdfium appends a
      // synthesized space to the extracted text when a layout gap follows the
      // object (the page char list carries only the real chars), so the trimmed
      // form is accepted too — critical for icon glyphs (PUA codepoints no
      // installed font could redraw), which otherwise land in the redraw text
      // and fail the whole edit on font coverage.
      const cpLen = [...t.text].length
      const trimmedCpLen = [...t.text.replace(/\s+$/, '')].length
      if (chars.length > 0 && (chars.length === cpLen || chars.length === trimmedCpLen)) {
        let advW = 0
        for (const ci of chars) {
          if (!m._FPDFText_GetLooseCharBox(textPage, ci, rectPtr)) {
            advW = NaN
            break
          }
          advW += m.HEAPF32[(rectPtr >> 2) + 2]! - m.HEAPF32[rectPtr >> 2]!
        }
        if (Number.isFinite(advW) && m._FPDFText_GetCharOrigin(textPage, chars[0]!, xPtr, yPtr)) {
          g = { origX: m.HEAPF64[xPtr >> 3]!, origY: m.HEAPF64[yPtr >> 3]!, advW }
        }
      }
      geo.push(g)
      const units = foldMap(t.text).units
      unitCount.push(units.length)
      engUnits.push(...units)
    }
    const tgt = foldMap(newText)
    let p = 0
    const maxP = Math.min(tgt.units.length, engUnits.length)
    while (p < maxP && tgt.units[p] === engUnits[p]) p++
    // Whitespace-only boundary edits change no fold unit, so p/s alone would keep
    // (and re-place at original spacing) the very objects whose gap the user edited.
    // Clamp the kept regions past the units adjacent to each whitespace-edited seam;
    // fragment matches align the user's units into the container's fold first (off).
    // The head clamp must land before the suffix scan: a space-only edit leaves the
    // unit sequences identical, p covers everything, and s would be capped at 0.
    const ws = wsEditClamp(edit.oldText, edit.newText)
    const oldUnits = ws ? foldMap(edit.oldText).units : []
    const off = !ws
      ? -1
      : engUnits.length === oldUnits.length
        ? 0
        : indexOfUnits(engUnits, oldUnits)
    if (ws && off >= 0 && ws.headNS !== null) p = Math.min(p, Math.max(0, off + ws.headNS))
    let s = 0
    const maxS = maxP - p
    while (s < maxS && tgt.units[tgt.units.length - 1 - s] === engUnits[engUnits.length - 1 - s])
      s++
    if (ws && off >= 0 && ws.tailNS !== null) {
      const extra = engUnits.length - off - oldUnits.length
      s = Math.min(s, Math.max(0, extra + ws.tailNS))
    }
    const shift = engUnits.length - tgt.units.length
    const out: KeepPlanObj[] = []
    let a = 0
    let prevEnd = -1
    for (const [oi, t] of ordered.entries()) {
      const b = a + unitCount[oi]!
      const g = geo[oi]
      const inPrefix = b <= p
      const inSuffix = a >= engUnits.length - s
      if (g && b > a && (inPrefix || inSuffix)) {
        const ua = inPrefix ? a : a - shift
        const startK = tgt.idx[ua]!
        const endK = tgt.end[(inPrefix ? b : b - shift) - 1]!
        if (startK >= prevEnd && endK > startK && !newText.slice(startK, endK).includes('\n')) {
          // Fill override must be uniform across the object, or it needs a split.
          // Face/size overrides can never apply to a kept object at all — a range
          // styled beyond color always redraws (color = undefined skips the keep).
          let color: Rgb | null | undefined = newColor ?? null
          if (!newColor && charStyles) {
            const colorKey = (k: number) => {
              const c = (charStyles[k] ?? null)?.color
              return c ? c.join(',') : ''
            }
            const first = charStyles[startK] ?? null
            color = first?.color ?? null
            for (let k = startK; k < endK; k++) {
              if (styledBeyondColor(charStyles[k] ?? null) || colorKey(k) !== colorKey(startK)) {
                color = undefined
                break
              }
            }
          }
          if (color !== undefined) {
            out.push({ src: t, startK, endK, advW: g.advW, origX: g.origX, origY: g.origY, color })
            prevEnd = endK
          }
        }
      }
      a = b
    }
    return out.length > 0 ? out : null
  } finally {
    for (const ptr of [rectPtr, xPtr, yPtr]) m._free(ptr)
  }
}

async function rebuildRun(
  m: Pdfium,
  doc: number,
  page: number,
  edit: TextEditInput,
  matches: PageTextObj[],
  newText: string,
  textPage: number,
): Promise<boolean> {
  const anchor = matches.reduce((a, b) => (b.bounds[0] < a.bounds[0] ? b : a))
  const matPtr = m._malloc(24)
  const sizePtr = m._malloc(4)
  const colPtr = m._malloc(16)
  const widthPtr = m._malloc(4)
  try {
    m._FPDFPageObj_GetMatrix(anchor.obj, matPtr)
    let matrix = Array.from(m.HEAPF32.subarray(matPtr >> 2, (matPtr >> 2) + 6))
    let fontSize = edit.fontSize
    if (m._FPDFTextObj_GetFontSize(anchor.obj, sizePtr)) fontSize = m.HEAPF32[sizePtr >> 2]!
    if (edit.newFontSize !== undefined && edit.newFontSize > 0) fontSize = edit.newFontSize
    if (edit.origin) {
      // Paragraph rebuild: the renderer measured wrap width, leading and x offsets
      // in PDF user space, so write in user space too — identity matrix at the
      // origin and the renderer's effective size. The anchor's own matrix may carry
      // scale (Tf ~ 1 conventions), which would re-scale the user-space leading and
      // offsets and spread lines apart.
      matrix = [1, 0, 0, 1, edit.origin[0], edit.origin[1]]
      fontSize =
        edit.newFontSize !== undefined && edit.newFontSize > 0 ? edit.newFontSize : edit.fontSize
    }
    const hasColor = m._FPDFPageObj_GetFillColor(
      anchor.obj,
      colPtr,
      colPtr + 4,
      colPtr + 8,
      colPtr + 12,
    )
    const color: readonly [number, number, number, number] = edit.newColor
      ? [edit.newColor[0], edit.newColor[1], edit.newColor[2], 255]
      : hasColor
        ? ([0, 4, 8, 12].map((off) => m.HEAPU8[colPtr + off]!) as [number, number, number, number])
        : [0, 0, 0, 255]
    // Some producers (Chrome print) leave the object's fill alpha reading as 0 while
    // the glyphs render opaque (the RGB part reads fine). Inheriting that raw 0 drew
    // the rebuilt run fully transparent — visible text must never become invisible.
    if (color[3] === 0) (color as [number, number, number, number])[3] = 255

    // Colors already in the matched run survive underneath the user's selection
    // styles; an explicit whole-edit newColor means "repaint uniformly" and wins.
    const keepColors = edit.newColor ? null : matchCharColors(m, matches, newText)
    const charStyles = overlayStyles(keepColors, plannedCharStyles(edit, newText), newText)

    // Object-preserving mode. Whole-edit face/size/style overrides mean "restyle
    // everything" and take the full-redraw path; so does a rotated/skewed matrix
    // (the cursor walk below only models horizontal text). Selection-level style
    // runs are fine: buildKeepPlan redraws any object their ranges touch.
    const styleOverride =
      edit.newFontSize !== undefined || edit.newFont !== undefined || edit.newBold || edit.newItalic
    const axisAligned =
      Math.abs(matrix[1]!) < 1e-4 && Math.abs(matrix[2]!) < 1e-4 && matrix[0]! > 0 && matrix[3]! > 0
    const keeps =
      !styleOverride && axisAligned
        ? buildKeepPlan(m, textPage, matches, newText, charStyles, edit.newColor, edit)
        : null
    // Subset the rebuild font to the non-kept text only: requiring coverage for
    // kept glyphs (e.g. Type3 PUA codepoints) would fail edits that never touch them
    const redrawOnly = (): string => {
      if (!keeps) return newText
      const parts: string[] = []
      let k = 0
      for (const kp of [...keeps].sort((a, b) => a.startK - b.startK)) {
        parts.push(newText.slice(k, kp.startK))
        k = kp.endK
      }
      parts.push(newText.slice(k))
      return parts.join('')
    }

    let font = 0
    let anyCff = false
    const loadFont = async (styleEdit: TextEditInput, text: string): Promise<number> => {
      const fontBytes = await rebuildFontBytes(m, anchor.font, styleEdit, text.trim() ? text : 'x')
      const fontPtr = m._malloc(fontBytes.length)
      m.HEAPU8.set(fontBytes, fontPtr)
      const f = m._FPDFText_LoadFont(
        doc,
        fontPtr,
        fontBytes.length,
        isTruetype(fontBytes) ? FPDF_FONT_TRUETYPE : FPDF_FONT_TYPE1,
        1,
      )
      m._free(fontPtr)
      if (!f) throw new Error('FPDFText_LoadFont failed')
      anyCff = anyCff || !isTruetype(fontBytes)
      return f
    }
    const loadRebuild = async (text: string) => {
      font = await loadFont(edit, text)
    }
    await loadRebuild(keeps ? redrawOnly() : newText)

    // Effective face/size of a segment style: explicit run fields over the whole-edit
    // overrides (which rebuildFontBytes resolves against the original run)
    const faceOf = (s: RunStyle | null) => ({
      font: s?.font ?? edit.newFont,
      bold: s?.bold ?? !!edit.newBold,
      italic: s?.italic ?? !!edit.newItalic,
    })
    const faceKeyOf = (s: RunStyle | null) => {
      const f = faceOf(s)
      return `${f.font ?? ''}|${f.bold ? 1 : 0}|${f.italic ? 1 : 0}`
    }
    const baseFaceKey = faceKeyOf(null)
    const sizeOf = (s: RunStyle | null) => (s?.size !== undefined && s.size > 0 ? s.size : fontSize)

    // One extra font per distinct non-base face among the styled chars, subset to
    // exactly the text that face draws (styled ranges are never kept, so the set is
    // known up front and survives a PreserveAbort re-load of the base font)
    const styleFonts = new Map<string, number>()
    if (charStyles) {
      const byFace = new Map<string, { style: RunStyle; text: string }>()
      for (let k = 0; k < newText.length; k++) {
        const s = charStyles[k] ?? null
        if (!s) continue
        const fk = faceKeyOf(s)
        if (fk === baseFaceKey) continue
        const e = byFace.get(fk)
        if (e) e.text += newText[k]!
        else byFace.set(fk, { style: s, text: newText[k]! })
      }
      for (const [fk, { style, text }] of byFace) {
        const face = faceOf(style)
        styleFonts.set(
          fk,
          await loadFont(
            { ...edit, newFont: face.font, newBold: face.bold, newItalic: face.italic },
            text,
          ),
        )
      }
    }
    const fontFor = (s: RunStyle | null): number => styleFonts.get(faceKeyOf(s)) ?? font

    // Advance of one codepoint in text-space pt, measured with the face and size that
    // will draw it (same unicode→charcode mapping FPDFText_SetText uses)
    const advancePt = (cp: number, s: RunStyle | null): number | null =>
      m._FPDFFont_GetGlyphWidth(fontFor(s), cp, 1, widthPtr)
        ? m.HEAPF32[widthPtr >> 2]! * sizeOf(s)
        : null

    const lineHeight = edit.lineLeading ?? fontSize * LINE_GAP
    const [baseX, baseY] = edit.origin ?? [matrix[4]!, matrix[5]!]
    const newObjs: number[] = []
    const moves: { obj: number; dx: number; dy: number; color: Rgb | null }[] = []
    // Per new object: the kept object preceding it in text order (0 = none), so
    // redrawn fragments can be inserted next to their kept neighbors — otherwise
    // stream-order extraction (copy/paste, pdftotext) reads jumbled text
    const segAnchors: number[] = []
    let lastKeptObj = 0

    const makeSeg = (text: string, x: number, y: number, style: RunStyle | null) => {
      const newObj = m._FPDFPageObj_CreateTextObj(doc, fontFor(style), sizeOf(style))
      const textPtr = utf16Ptr(m, text)
      const ok = m._FPDFText_SetText(newObj, textPtr)
      m._free(textPtr)
      if (!ok) {
        m._FPDFPageObj_Destroy(newObj)
        throw new Error('FPDFText_SetText failed on rebuilt object')
      }
      const lineMatrix = [...matrix]
      lineMatrix[4] = x
      lineMatrix[5] = y
      m.HEAPF32.set(lineMatrix, matPtr >> 2)
      m._FPDFPageObj_SetMatrix(newObj, matPtr)
      const c = style?.color ? [style.color[0], style.color[1], style.color[2], 255] : color
      m._FPDFPageObj_SetFillColor(newObj, c[0]!, c[1]!, c[2]!, c[3]!)
      newObjs.push(newObj)
      segAnchors.push(lastKeptObj)
    }

    // Full redraw: one text object per line (several when selection styles split
    // it), stepping down one leading per line from the anchor baseline or origin
    const buildRedraw = () => {
      let lineStart = 0
      for (const [lineIdx, line] of newText.split('\n').entries()) {
        const lineStyles = charStyles ? charStyles.slice(lineStart, lineStart + line.length) : null
        lineStart += line.length + 1
        if (!line) continue
        const drop = lineIdx * lineHeight
        for (const seg of segmentLine(line, lineStyles, advancePt)) {
          if (!seg.text) continue
          // Segment offset is a text-space advance: map it through the matrix's x axis
          const segX = seg.xPt
          makeSeg(
            seg.text,
            baseX + (edit.lineXOffsets?.[lineIdx] ?? 0) + matrix[0]! * segX - matrix[2]! * drop,
            baseY + matrix[1]! * segX - matrix[3]! * drop,
            seg.style,
          )
        }
      }
    }

    // Preserving build: walk each output line with a page-space cursor — kept
    // objects become translations to the cursor, edited stretches are drawn with
    // the rebuild font. Consecutive kept objects carry their ORIGINAL spacing, so
    // untouched (even justified) text keeps its exact layout.
    const buildPreserved = (plan: KeepPlanObj[]) => {
      const sorted = [...plan].sort((a, b) => a.startK - b.startK)
      let ki = 0
      let lineStart = 0
      // Advances/leadings are text-space (Tf) values; map through the matrix's
      // axis scales to walk in page space (scale 1 for origin edits)
      const xScale = matrix[0]!
      const advOf = (ch: string, s: RunStyle | null): number => {
        const a = advancePt(ch.codePointAt(0)!, s)
        if (a !== null) return a * xScale
        if (/\s/.test(ch)) return sizeOf(s) * xScale * 0.28
        throw new PreserveAbort()
      }
      for (const [lineIdx, line] of newText.split('\n').entries()) {
        const lineEnd = lineStart + line.length
        const baseline = baseY - lineIdx * lineHeight * matrix[3]!
        let cursor = baseX + (edit.lineXOffsets?.[lineIdx] ?? 0)
        let prev: { keep: KeepPlanObj; placedX: number } | null = null
        let k = lineStart
        while (k < lineEnd) {
          const keep = ki < sorted.length && sorted[ki]!.startK === k ? sorted[ki]! : null
          if (keep) {
            if (
              prev &&
              Math.abs(keep.origY - prev.keep.origY) < 0.5 &&
              keep.origX > prev.keep.origX &&
              /^\s*$/.test(newText.slice(prev.keep.endK, keep.startK))
            ) {
              cursor = prev.placedX + (keep.origX - prev.keep.origX)
            }
            moves.push({
              obj: keep.src.obj,
              dx: cursor - keep.origX,
              dy: baseline - keep.origY,
              color: keep.color,
            })
            lastKeptObj = keep.src.obj
            prev = { keep, placedX: cursor }
            cursor += keep.advW
            k = keep.endK
            ki++
            continue
          }
          const runEnd = ki < sorted.length ? Math.min(sorted[ki]!.startK, lineEnd) : lineEnd
          if (runEnd <= k) throw new PreserveAbort()
          const runText = newText.slice(k, runEnd)
          if (runText.trim()) {
            const runStyles = charStyles ? charStyles.slice(k, runEnd) : null
            for (const seg of segmentLine(runText, runStyles, advancePt)) {
              if (!seg.text.trim()) continue
              makeSeg(seg.text, cursor + seg.xPt * xScale, baseline, seg.style)
            }
            prev = null
          }
          let cu = k
          for (const ch of runText) {
            cursor += advOf(ch, charStyles?.[cu] ?? null)
            cu += ch.length
          }
          k = runEnd
        }
        lineStart = lineEnd + 1
      }
      if (ki !== sorted.length) throw new PreserveAbort()
    }

    // All new objects are built before anything is removed or moved, so a failure
    // here leaves the page untouched (the edit is then skipped, not half-applied).
    try {
      if (keeps) {
        try {
          buildPreserved(keeps)
        } catch (err) {
          if (!(err instanceof PreserveAbort)) throw err
          for (const o of newObjs) m._FPDFPageObj_Destroy(o)
          newObjs.length = 0
          moves.length = 0
          segAnchors.length = 0
          lastKeptObj = 0
          // The redraw-only subset lacks the kept glyphs; the fallback draws them
          await loadRebuild(newText)
          buildRedraw()
        }
      } else {
        buildRedraw()
      }
    } catch (err) {
      for (const o of newObjs) m._FPDFPageObj_Destroy(o)
      throw err
    }

    const kept = new Set(moves.map((v) => v.obj))
    for (const v of moves) {
      m._FPDFPageObj_GetMatrix(v.obj, matPtr)
      m.HEAPF32[(matPtr >> 2) + 4] += v.dx
      m.HEAPF32[(matPtr >> 2) + 5] += v.dy
      m._FPDFPageObj_SetMatrix(v.obj, matPtr)
      if (v.color) m._FPDFPageObj_SetFillColor(v.obj, v.color[0]!, v.color[1]!, v.color[2]!, 255)
    }
    const removed = matches.filter((t) => !kept.has(t.obj))
    for (const t of removed) {
      m._FPDFPage_RemoveObject(page, t.obj)
      m._FPDFPageObj_Destroy(t.obj)
    }
    if (kept.size > 0 && newObjs.length > 0 && m._FPDFPage_InsertObjectAtIndex) {
      // Keep the content stream in reading order: each redrawn fragment goes right
      // after its kept predecessor (anchorless fragments lead the run). Indices are
      // post-removal; groups insert back-to-front so earlier anchors stay valid.
      const idxOf = new Map<number, number>()
      const count = m._FPDFPage_CountObjects(page)
      for (let i = 0; i < count; i++) idxOf.set(m._FPDFPage_GetObject(page, i), i)
      const runStart = Math.min(
        ...matches.filter((t) => kept.has(t.obj)).map((t) => idxOf.get(t.obj) ?? 0),
      )
      const groups: { at: number; objs: number[] }[] = []
      for (const [i, newObj] of newObjs.entries()) {
        const anchor = segAnchors[i]!
        const at = anchor && idxOf.has(anchor) ? idxOf.get(anchor)! + 1 : runStart
        const g = groups[groups.length - 1]
        if (g && g.at === at) g.objs.push(newObj)
        else groups.push({ at, objs: [newObj] })
      }
      groups.sort((a, b) => b.at - a.at)
      for (const g of groups) {
        for (const [j, newObj] of g.objs.entries()) {
          m._FPDFPage_InsertObjectAtIndex(page, newObj, g.at + j)
        }
      }
    } else {
      // Full redraw: insert where the removed run began — the run's minimum index
      // is the only position still valid after the removals shift everything else
      const insertAt = removed.length > 0 ? Math.min(...removed.map((t) => t.index)) : -1
      for (const [i, newObj] of newObjs.entries()) {
        if (insertAt >= 0 && m._FPDFPage_InsertObjectAtIndex) {
          m._FPDFPage_InsertObjectAtIndex(page, newObj, insertAt + i)
        } else {
          m._FPDFPage_InsertObject(page, newObj)
        }
      }
    }
    return anyCff
  } finally {
    for (const p of [matPtr, sizePtr, colPtr, widthPtr]) m._free(p)
  }
}

/** PDFium's FPDF_FONT_TYPE1 load path files a CFF-flavored sfnt under /FontFile — the
    Type1-program slot — which poppler/mupdf/Acrobat reject (blank glyphs). Relabel such
    programs as /FontFile3 with Subtype /OpenType (PDF 1.6), the correct slot for them.
    Only descriptors whose embedded program actually starts with the OTTO tag are touched,
    so genuine Type1 fonts from the original document pass through untouched. */
async function relabelOpenTypeFontFiles(bytes: Uint8Array): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes)
  let touched = false
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFDict)) continue
    if (obj.get(PDFName.of('Type')) !== PDFName.of('FontDescriptor')) continue
    const ff = obj.get(PDFName.of('FontFile'))
    if (!ff) continue
    const stream = doc.context.lookup(ff)
    if (!(stream instanceof PDFRawStream)) continue
    let program: Uint8Array
    try {
      program = decodePDFRawStream(stream).decode()
    } catch {
      continue
    }
    if (Buffer.from(program.subarray(0, 4)).toString('latin1') !== 'OTTO') continue
    obj.delete(PDFName.of('FontFile'))
    obj.set(PDFName.of('FontFile3'), ff)
    stream.dict.set(PDFName.of('Subtype'), PDFName.of('OpenType'))
    touched = true
  }
  return touched ? doc.save({ useObjectStreams: false }) : bytes
}

export function saveDoc(m: Pdfium, doc: number): Uint8Array {
  const writer = m._PDFiumExt_OpenFileWriter()
  try {
    if (!m._PDFiumExt_SaveAsCopy(doc, writer)) throw new Error('PDFium SaveAsCopy failed')
    const size = m._PDFiumExt_GetFileWriterSize(writer)
    const buf = m._malloc(size)
    m._PDFiumExt_GetFileWriterData(writer, buf, size)
    const out = Uint8Array.from(m.HEAPU8.subarray(buf, buf + size))
    m._free(buf)
    return out
  } finally {
    m._PDFiumExt_CloseFileWriter(writer)
  }
}

/** The wasm instance is a shared singleton with global heap state; edits must not interleave */
let applyChain: Promise<unknown> = Promise.resolve()

/** Serialize a pdfium operation behind every previously queued one (shared with image-edit) */
export function chainPdfium<T>(fn: () => Promise<T>): Promise<T> {
  const run = applyChain.then(fn)
  applyChain = run.catch(() => undefined)
  return run
}

const errMsg = (err: unknown) => (err instanceof Error ? err.message : String(err))

export interface TextEditsResult {
  bytes: Uint8Array
  /** Edits that could not be applied; everything else is in `bytes`. One stale edit
      must not block the whole save (it used to make every save fail until removed). */
  skipped: TextEditFailure[]
}

export interface TextInsertsResult {
  bytes: Uint8Array
  skipped: TextInsertFailure[]
}

/** Rewrite page content streams per the edits and return the new document bytes plus
    the edits that no longer match the document (skipped, reported to the caller). */
export function applyTextEdits(
  bytes: Uint8Array,
  edits: TextEditInput[],
): Promise<TextEditsResult> {
  return chainPdfium(() => applyTextEditsInner(bytes, edits))
}

/** Text-space axes that keep inserted glyphs upright after the page's final /Rotate. */
export function textInsertAxes(rotate = 0): readonly [number, number, number, number] {
  switch (((rotate % 360) + 360) % 360) {
    case 90:
      return [0, 1, -1, 0]
    case 180:
      return [-1, 0, 0, -1]
    case 270:
      return [0, -1, 1, 0]
    default:
      return [1, 0, 0, 1]
  }
}

/** Insert new searchable text objects without matching/removing existing page content. */
export function applyTextInserts(
  bytes: Uint8Array,
  inserts: TextInsertInput[],
): Promise<TextInsertsResult> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    const skipped: TextInsertFailure[] = []
    return withDocument(m, bytes, async (doc) => {
      const pageCount = m._FPDF_GetPageCount(doc)
      let appliedTotal = 0
      let embeddedCff = false
      const byPage = new Map<number, { input: TextInsertInput; editIndex: number }[]>()
      inserts.forEach((input, editIndex) => {
        if (input.pageIndex < 0 || input.pageIndex >= pageCount) {
          skipped.push({ editIndex, pageIndex: input.pageIndex, reason: 'page does not exist' })
          return
        }
        byPage.set(input.pageIndex, [...(byPage.get(input.pageIndex) ?? []), { input, editIndex }])
      })
      for (const [pageIndex, pageInserts] of byPage) {
        const page = m._FPDF_LoadPage(doc, pageIndex)
        if (!page) throw new Error(`could not load page ${pageIndex + 1}`)
        let applied = 0
        try {
          for (const { input, editIndex } of pageInserts) {
            const text = input.text.trim()
            if (!text) {
              skipped.push({ editIndex, pageIndex, reason: 'empty inserted text' })
              continue
            }
            const pseudoEdit: TextEditInput = {
              pageIndex,
              rect: [input.origin[0], input.origin[1], input.origin[0], input.origin[1]],
              oldText: '',
              newText: input.text,
              fontSize: input.fontSize,
              newFontSize: input.fontSize,
              newColor: input.color,
              newFont: input.font,
              newBold: input.bold,
              newItalic: input.italic,
            }
            const created: number[] = []
            let font = 0
            try {
              const fontBytes = await rebuildFontBytes(m, 0, pseudoEdit, input.text)
              const fontPtr = m._malloc(fontBytes.length)
              m.HEAPU8.set(fontBytes, fontPtr)
              font = m._FPDFText_LoadFont(
                doc,
                fontPtr,
                fontBytes.length,
                isTruetype(fontBytes) ? FPDF_FONT_TRUETYPE : FPDF_FONT_TYPE1,
                1,
              )
              m._free(fontPtr)
              if (!font) throw new Error('FPDFText_LoadFont failed')
              const matrixPtr = m._malloc(24)
              try {
                const leading = input.lineLeading ?? input.fontSize * LINE_GAP
                const [a, b, c, d] = textInsertAxes(input.rotate)
                for (const [lineIndex, line] of input.text.split('\n').entries()) {
                  if (!line) continue
                  const obj = m._FPDFPageObj_CreateTextObj(doc, font, input.fontSize)
                  const textPtr = utf16Ptr(m, line)
                  const ok = m._FPDFText_SetText(obj, textPtr)
                  m._free(textPtr)
                  if (!ok) {
                    m._FPDFPageObj_Destroy(obj)
                    throw new Error('FPDFText_SetText failed on inserted object')
                  }
                  const offset = input.lineXOffsets?.[lineIndex] ?? 0
                  const drop = lineIndex * leading
                  m.HEAPF32.set(
                    [
                      a,
                      b,
                      c,
                      d,
                      input.origin[0] + a * offset - c * drop,
                      input.origin[1] + b * offset - d * drop,
                    ],
                    matrixPtr >> 2,
                  )
                  m._FPDFPageObj_SetMatrix(obj, matrixPtr)
                  m._FPDFPageObj_SetFillColor(
                    obj,
                    input.color[0],
                    input.color[1],
                    input.color[2],
                    255,
                  )
                  created.push(obj)
                }
              } finally {
                m._free(matrixPtr)
              }
              for (const obj of created) m._FPDFPage_InsertObject(page, obj)
              embeddedCff = !isTruetype(fontBytes) || embeddedCff
              applied++
            } catch (err) {
              for (const obj of created) m._FPDFPageObj_Destroy(obj)
              skipped.push({ editIndex, pageIndex, reason: errMsg(err) })
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
      if (appliedTotal === 0) return { bytes, skipped }
      const saved = saveDoc(m, doc)
      return { bytes: embeddedCff ? await relabelOpenTypeFontFiles(saved) : saved, skipped }
    })
  })
}

/** Dry-run matching for pending edits (no mutation). Lets the renderer reject a bad edit
    when it is made (not at save) and cover the matched run's real ink in the preview. */
export function validateTextEdits(
  bytes: Uint8Array,
  edits: TextEditInput[],
): Promise<TextEditValidation[]> {
  return chainPdfium(() => validateTextEditsInner(bytes, edits))
}

export async function withDocument<T>(
  m: Pdfium,
  bytes: Uint8Array,
  fn: (doc: number) => Promise<T>,
): Promise<T> {
  const docPtr = m._malloc(bytes.length)
  m.HEAPU8.set(bytes, docPtr)
  const doc = m._FPDF_LoadMemDocument(docPtr, bytes.length, 0)
  if (!doc) {
    m._free(docPtr)
    throw new Error('PDFium could not load the document')
  }
  try {
    return await fn(doc)
  } finally {
    m._FPDF_CloseDocument(doc)
    m._free(docPtr)
  }
}

function groupByPage(
  m: Pdfium,
  doc: number,
  edits: TextEditInput[],
  outOfRange: (edit: TextEditInput) => void,
): Map<number, TextEditInput[]> {
  const pageCount = m._FPDF_GetPageCount(doc)
  const byPage = new Map<number, TextEditInput[]>()
  for (const e of edits) {
    if (e.pageIndex < 0 || e.pageIndex >= pageCount) {
      outOfRange(e)
      continue
    }
    byPage.set(e.pageIndex, [...(byPage.get(e.pageIndex) ?? []), e])
  }
  return byPage
}

async function applyTextEditsInner(
  bytes: Uint8Array,
  edits: TextEditInput[],
): Promise<TextEditsResult> {
  const m = await loadPdfium()
  const skipped: TextEditFailure[] = []
  const skip = (edit: TextEditInput, reason: string) =>
    skipped.push({ pageIndex: edit.pageIndex, oldText: edit.oldText, reason })

  return withDocument(m, bytes, async (doc) => {
    const byPage = groupByPage(m, doc, edits, (e) => skip(e, 'page does not exist'))
    let appliedTotal = 0
    let embeddedCff = false
    for (const [pageIndex, pageEdits] of byPage) {
      const page = m._FPDF_LoadPage(doc, pageIndex)
      if (!page) throw new Error(`could not load page ${pageIndex + 1}`)
      const textPage = m._FPDFText_LoadPage(page)
      try {
        const objects = collectTextObjects(m, page, textPage)
        // Two edits resolving to the same object would double-remove it; first claim wins
        const claimed = new Set<number>()
        const planned: {
          edit: TextEditInput
          matches: PageTextObj[]
          newText: string
          whole: boolean
        }[] = []
        for (const edit of pageEdits) {
          const res = matchEdit(objects, edit)
          if ('reason' in res) {
            skip(edit, res.reason)
          } else if (res.matches.some((t) => claimed.has(t.obj))) {
            skip(edit, 'overlaps another pending text edit')
          } else {
            for (const t of res.matches) claimed.add(t.obj)
            planned.push({ edit, ...res })
          }
        }
        // Descending object order keeps pending InsertObjectAtIndex targets valid
        planned.sort((a, b) => b.matches[0]!.index - a.matches[0]!.index)
        let applied = 0
        for (const { edit, matches, newText, whole } of planned) {
          // Pure move: translate the matched objects in page space and keep every
          // glyph as it is — no font resolution, no rebuild. Only a whole match may
          // move (a fragment's container carries surrounding text that must stay put).
          if (edit.translate) {
            if (!whole) {
              skip(edit, 'the text block cannot be moved as one unit')
              continue
            }
            const [dx, dy] = edit.translate
            for (const t of matches) m._FPDFPageObj_Transform(t.obj, 1, 0, 0, 1, dx, dy)
            applied++
            continue
          }
          // Deletion: an empty (or whitespace-only) planned replacement removes the
          // matched objects outright — no font resolution, nothing to rebuild
          if (newText.trim() === '') {
            for (const t of matches) {
              m._FPDFPage_RemoveObject(page, t.obj)
              m._FPDFPageObj_Destroy(t.obj)
            }
            applied++
            continue
          }
          // A fragment match rewrites its whole container run; the paragraph position
          // overrides would drag the container's surrounding text to the block corner
          const eff = whole
            ? edit
            : { ...edit, origin: undefined, lineLeading: undefined, lineXOffsets: undefined }
          try {
            if (canReuseFont(eff, newText, matches, objects)) {
              const textPtr = utf16Ptr(m, newText)
              const ok = m._FPDFText_SetText(matches[0]!.obj, textPtr)
              m._free(textPtr)
              if (!ok) throw new Error('FPDFText_SetText failed')
            } else {
              embeddedCff =
                (await rebuildRun(m, doc, page, eff, matches, newText, textPage)) || embeddedCff
            }
            applied++
          } catch (err) {
            skip(edit, errMsg(err))
          }
        }
        if (applied > 0 && !m._FPDFPage_GenerateContent(page)) {
          throw new Error(`could not regenerate page ${pageIndex + 1}`)
        }
        appliedTotal += applied
      } finally {
        m._FPDFText_ClosePage(textPage)
        m._FPDF_ClosePage(page)
      }
    }
    // Nothing applied → return the input bytes untouched (skip the pdfium round-trip)
    if (appliedTotal === 0) return { bytes, skipped }
    const saved = saveDoc(m, doc)
    return { bytes: embeddedCff ? await relabelOpenTypeFontFiles(saved) : saved, skipped }
  })
}

/** Union of the matched objects' ink bounds */
function matchBounds(matches: PageTextObj[]): [number, number, number, number] {
  return matches
    .slice(1)
    .reduce(
      (u, t) => [
        Math.min(u[0], t.bounds[0]),
        Math.min(u[1], t.bounds[1]),
        Math.max(u[2], t.bounds[2]),
        Math.max(u[3], t.bounds[3]),
      ],
      [...matches[0]!.bounds] as [number, number, number, number],
    )
}

async function validateTextEditsInner(
  bytes: Uint8Array,
  edits: TextEditInput[],
): Promise<TextEditValidation[]> {
  const m = await loadPdfium()
  return withDocument(m, bytes, async (doc) => {
    const results: TextEditValidation[] = edits.map(() => ({ reason: null }))
    const indexOf = new Map(edits.map((e, i) => [e, i]))
    const byPage = groupByPage(m, doc, edits, (e) => {
      results[indexOf.get(e)!] = { reason: 'page does not exist' }
    })
    for (const [pageIndex, pageEdits] of byPage) {
      const page = m._FPDF_LoadPage(doc, pageIndex)
      if (!page) {
        for (const e of pageEdits)
          results[indexOf.get(e)!] = { reason: `could not load page ${pageIndex + 1}` }
        continue
      }
      const textPage = m._FPDFText_LoadPage(page)
      const colPtr = m._malloc(16)
      try {
        const objects = collectTextObjects(m, page, textPage)
        for (const e of pageEdits) {
          const res = matchEdit(objects, e)
          if ('reason' in res) {
            results[indexOf.get(e)!] = { reason: res.reason }
            continue
          }
          // Same gate the apply path enforces: a move must own its objects outright
          if (e.translate && !res.whole) {
            results[indexOf.get(e)!] = { reason: 'the text block cannot be moved as one unit' }
            continue
          }
          // Fragment edits resolve to their whole container object; covering all of it
          // would blank the surrounding text the preview doesn't redraw. Keep the edit
          // rect's x-span there and take only the ink's vertical extent.
          const b = matchBounds(res.matches)
          const whole = norm(res.matches.map((t) => t.text).join('')) === norm(e.oldText)
          // Colors an earlier edit saved into the run, in oldText offsets: whole
          // matches share oldText's folded units, so ranks map straight across.
          // Editors seed their selection-color state from these so the colors
          // show while editing and the commit repaints them explicitly.
          const kept = whole ? matchCharColors(m, res.matches, e.oldText) : null
          const runs: { start: number; end: number; c: Rgb }[] = []
          if (kept) {
            for (let k = 0; k < kept.length; k++) {
              const c = kept[k]
              if (!c) continue
              const last = runs[runs.length - 1]
              if (last && last.end === k && last.c === c) last.end = k + 1
              else runs.push({ start: k, end: k + 1, c })
            }
          }
          // Base ink of the run (first object in reading order), so the editor can
          // display the document's actual color even when the run is uniform
          const first = readingOrder(res.matches)[0]!
          const baseColor = m._FPDFPageObj_GetFillColor(
            first.obj,
            colPtr,
            colPtr + 4,
            colPtr + 8,
            colPtr + 12,
          )
            ? ([m.HEAPU8[colPtr]!, m.HEAPU8[colPtr + 4]!, m.HEAPU8[colPtr + 8]!] as [
                number,
                number,
                number,
              ])
            : undefined
          results[indexOf.get(e)!] = {
            reason: null,
            bounds: whole ? b : [e.rect[0], b[1], e.rect[2], b[3]],
            colorRuns:
              runs.length > 0
                ? runs.map((r) => ({
                    start: r.start,
                    end: r.end,
                    color: [r.c[0], r.c[1], r.c[2]] as [number, number, number],
                  }))
                : undefined,
            baseColor,
          }
        }
      } finally {
        m._free(colPtr)
        m._FPDFText_ClosePage(textPage)
        m._FPDF_ClosePage(page)
      }
    }
    return results
  })
}

/**
 * Read-back verification: confirm every applied edit's replacement text is actually
 * extractable from the produced bytes. Runs on the final output (after the pdf-lib
 * round-trip), before it reaches disk — a failure aborts the save with the original
 * file untouched and the edits still pending, instead of silently losing work.
 * `pageIndex` here is the edit's page in the *final* document (caller remaps).
 */
export function verifyTextEdits(
  bytes: Uint8Array,
  edits: { pageIndex: number; newText: string }[],
): Promise<{ pageIndex: number; reason: string }[]> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async (doc) => {
      const failures: { pageIndex: number; reason: string }[] = []
      const pageCount = m._FPDF_GetPageCount(doc)
      const byPage = new Map<number, string[]>()
      for (const e of edits) {
        if (e.pageIndex < 0 || e.pageIndex >= pageCount) {
          failures.push({ pageIndex: e.pageIndex, reason: 'page missing from saved output' })
          continue
        }
        byPage.set(e.pageIndex, [...(byPage.get(e.pageIndex) ?? []), e.newText])
      }
      for (const [pageIndex, texts] of byPage) {
        const page = m._FPDF_LoadPage(doc, pageIndex)
        if (!page) {
          for (const _ of texts)
            failures.push({ pageIndex, reason: 'page unreadable in saved output' })
          continue
        }
        const textPage = m._FPDFText_LoadPage(page)
        try {
          // NFC on both sides: ToUnicode CMaps canonicalize singleton codepoints (e.g.
          // compatibility ideograph U+F900 extracts as U+8C48) — the glyph on the page is
          // right, only the reverse mapping differs, and that must not abort the save
          const canon = (s: string) => norm(s.normalize('NFC'))
          const objects = collectTextObjects(m, page, textPage)
          const pageText = canon(objects.map((o) => o.text).join(''))
          // The object-preserving rebuild can leave a reflowed line contiguous
          // only visually, not in stream order — accept either
          const visualText = canon(joinRows(objects))
          for (const newText of texts) {
            // Every non-empty line must be extractable (rebuilt runs are one object per line)
            const missing = newText
              .split('\n')
              .map(canon)
              .filter((l) => l.length > 0 && !pageText.includes(l) && !visualText.includes(l))
            if (missing.length > 0) {
              const snippet = missing[0]!.slice(0, 20)
              failures.push({
                pageIndex,
                reason: `replacement text missing from saved output ("${snippet}")`,
              })
            }
          }
        } finally {
          m._FPDFText_ClosePage(textPage)
          m._FPDF_ClosePage(page)
        }
      }
      return failures
    })
  })
}
