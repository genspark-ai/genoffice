import { readFile, rename, rm, writeFile } from 'node:fs/promises'
import {
  PDFArray,
  PDFBool,
  PDFDict,
  PDFDocument,
  PDFDropdown,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFOptionList,
  PDFRef,
  PDFString,
  degrees,
  rgb,
} from 'pdf-lib'
import type { PDFPage } from 'pdf-lib'
import { VISUAL_SIGNATURE_CONTENT_PREFIX } from '../shared/ipc'
import type {
  DrawingInput,
  FormValueInput,
  ImageEditFailure,
  MarkupInput,
  MetadataInput,
  NoteReplyTarget,
  SavePdfRequest,
  StaticFormFillRecord,
  TextEditFailure,
  TextInsertFailure,
} from '../shared/ipc'

const num = (v: number) => Math.round(v * 100) / 100
const STATIC_FORM_FILLS_KEY = PDFName.of('GenOfficeStaticFormFills')

function validStaticFormFill(value: unknown): value is StaticFormFillRecord {
  if (!value || typeof value !== 'object') return false
  const record = value as Partial<StaticFormFillRecord>
  return (
    typeof record.id === 'string' &&
    (record.kind === 'text' || record.kind === 'check' || record.kind === 'cross') &&
    Number.isInteger(record.pageIndex) &&
    Array.isArray(record.rect) &&
    record.rect.length === 4 &&
    record.rect.every((entry) => typeof entry === 'number' && Number.isFinite(entry))
  )
}

export async function readStaticFormFills(bytes: Uint8Array): Promise<StaticFormFillRecord[]> {
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  const value = pdfDoc.catalog.get(STATIC_FORM_FILLS_KEY)
  if (!(value instanceof PDFHexString)) return []
  try {
    const parsed: unknown = JSON.parse(value.decodeText())
    return Array.isArray(parsed) ? parsed.filter(validStaticFormFill) : []
  } catch {
    return []
  }
}

function resultingStaticFormFills(
  request: SavePdfRequest,
  pageCount: number,
): StaticFormFillRecord[] | undefined {
  if (request.staticFormFills === undefined) return undefined
  const deleted = new Set(request.deletedPages ?? [])
  const remaining =
    request.pageOrder?.filter((pageIndex) => !deleted.has(pageIndex)) ??
    Array.from({ length: pageCount }, (_, pageIndex) => pageIndex).filter(
      (pageIndex) => !deleted.has(pageIndex),
    )
  const newPageIndex = new Map(remaining.map((oldPageIndex, index) => [oldPageIndex, index]))
  return request.staticFormFills.flatMap((record) => {
    const pageIndex = newPageIndex.get(record.pageIndex)
    return pageIndex === undefined ? [] : [{ ...record, pageIndex }]
  })
}

function setVisualSignatureMetadata(annot: PDFDict, fieldName: string | undefined): void {
  if (!fieldName) return
  annot.set(PDFName.of('GenOfficeFormField'), PDFHexString.fromText(fieldName))
  annot.set(
    PDFName.of('Contents'),
    PDFHexString.fromText(`${VISUAL_SIGNATURE_CONTENT_PREFIX}${fieldName}`),
  )
}

const quadBounds = (q: number[]) => {
  const xs = [q[0]!, q[2]!, q[4]!, q[6]!]
  const ys = [q[1]!, q[3]!, q[5]!, q[7]!]
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)] as const
}

/**
 * Hand-written appearance stream (/AP /N), so viewers don't self-draw from QuadPoints —
 * Acrobat/Preview/pdfjs all render from AP for consistent results.
 * Highlight uses Multiply blending to mimic a highlighter; underline/strikeout are stroked
 * segments drawn along the "visual bottom edge" (pageRot is the page's final /Rotate;
 * at 90/270 line height runs along the x axis).
 */
function markupAppearance(
  pdfDoc: PDFDocument,
  m: MarkupInput,
  rect: number[],
  pageRot: number,
): ReturnType<typeof pdfDoc.context.stream> {
  const [r, g, b] = m.color
  const ops: string[] = []
  if (m.type === 'highlight') {
    ops.push('/GsM gs', `${r} ${g} ${b} rg`)
    for (const q of m.quads) {
      const [x1, y1, x2, y2] = quadBounds(q)
      ops.push(`${num(x1)} ${num(y1)} ${num(x2 - x1)} ${num(y2 - y1)} re f`)
    }
  } else {
    const t = m.type === 'underline' ? 0.08 : 0.46
    ops.push(`${r} ${g} ${b} RG`)
    for (const q of m.quads) {
      const [x1, y1, x2, y2] = quadBounds(q)
      const h = pageRot % 180 === 0 ? y2 - y1 : x2 - x1
      ops.push(`${Math.max(0.8, num(h * 0.06))} w`)
      if (pageRot === 90) {
        const x = x2 - h * t
        ops.push(`${num(x)} ${num(y1)} m ${num(x)} ${num(y2)} l S`)
      } else if (pageRot === 270) {
        const x = x1 + h * t
        ops.push(`${num(x)} ${num(y1)} m ${num(x)} ${num(y2)} l S`)
      } else {
        const y = pageRot === 180 ? y2 - h * t : y1 + h * t
        ops.push(`${num(x1)} ${num(y)} m ${num(x2)} ${num(y)} l S`)
      }
    }
  }
  return pdfDoc.context.stream(ops.join('\n'), {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: rect,
    Resources:
      m.type === 'highlight'
        ? { ExtGState: { GsM: { Type: 'ExtGState', BM: 'Multiply', ca: 1 } } }
        : {},
  })
}

const SUBTYPE: Record<MarkupInput['type'], string> = {
  highlight: 'Highlight',
  underline: 'Underline',
  strikeout: 'StrikeOut',
}

function addMarkup(pdfDoc: PDFDocument, page: PDFPage, m: MarkupInput): void {
  const xs = m.quads.flatMap((q) => [q[0]!, q[2]!, q[4]!, q[6]!])
  const ys = m.quads.flatMap((q) => [q[1]!, q[3]!, q[5]!, q[7]!])
  const rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
  const pageRot = ((page.getRotation().angle % 360) + 360) % 360
  const apRef = pdfDoc.context.register(markupAppearance(pdfDoc, m, rect, pageRot))
  const annot = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: SUBTYPE[m.type],
    Rect: rect,
    QuadPoints: m.quads.flat(),
    C: m.color,
    F: 4, // print
    T: 'GenOffice',
    P: page.ref,
    AP: { N: apRef },
  })
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
}

function appendAnnot(pdfDoc: PDFDocument, page: PDFPage, annotRef: PDFRef): void {
  const existing = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (existing) {
    existing.push(annotRef)
  } else {
    page.node.set(PDFName.of('Annots'), pdfDoc.context.obj([annotRef]))
  }
}

/** 4-segment Bezier approximation of an ellipse */
function ellipseOps(x1: number, y1: number, x2: number, y2: number): string[] {
  const k = 0.5522847
  const cx = (x1 + x2) / 2
  const cy = (y1 + y2) / 2
  const rx = Math.abs(x2 - x1) / 2
  const ry = Math.abs(y2 - y1) / 2
  const n = num
  return [
    `${n(cx + rx)} ${n(cy)} m`,
    `${n(cx + rx)} ${n(cy + k * ry)} ${n(cx + k * rx)} ${n(cy + ry)} ${n(cx)} ${n(cy + ry)} c`,
    `${n(cx - k * rx)} ${n(cy + ry)} ${n(cx - rx)} ${n(cy + k * ry)} ${n(cx - rx)} ${n(cy)} c`,
    `${n(cx - rx)} ${n(cy - k * ry)} ${n(cx - k * rx)} ${n(cy - ry)} ${n(cx)} ${n(cy - ry)} c`,
    `${n(cx + k * rx)} ${n(cy - ry)} ${n(cx + rx)} ${n(cy - k * ry)} ${n(cx + rx)} ${n(cy)} c`,
    'S',
  ]
}

/**
 * Image signature/stamp: a Stamp annotation whose appearance stream draws the embedded PNG.
 * The image is counter-rotated against the page's final /Rotate (viewers rotate annotation
 * appearances with the page), so it displays upright — matching the renderer preview.
 */
async function addImageStamp(
  pdfDoc: PDFDocument,
  page: PDFPage,
  d: Extract<DrawingInput, { kind: 'image' }>,
): Promise<void> {
  const png = await pdfDoc.embedPng(d.image)
  const [x1, y1, x2, y2] = d.rect
  const rw = x2 - x1
  const rh = y2 - y1
  const rot = ((page.getRotation().angle % 360) + 360) % 360
  // cm matrix mapping the image unit square into the BBox, pre-counter-rotated for the page
  const cm =
    rot === 90
      ? `0 ${num(rh)} ${num(-rw)} 0 ${num(rw)} 0`
      : rot === 180
        ? `${num(-rw)} 0 0 ${num(-rh)} ${num(rw)} ${num(rh)}`
        : rot === 270
          ? `0 ${num(-rh)} ${num(rw)} 0 0 ${num(rh)}`
          : `${num(rw)} 0 0 ${num(rh)} 0 0`
  const ap = pdfDoc.context.stream(`q ${cm} cm /Im0 Do Q`, {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: [0, 0, num(rw), num(rh)],
    Resources: { XObject: { Im0: png.ref } },
  })
  const annot = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: 'Stamp',
    Rect: [num(x1), num(y1), num(x2), num(y2)],
    F: 4,
    P: page.ref,
    AP: { N: pdfDoc.context.register(ap) },
  })
  annot.set(PDFName.of('T'), PDFHexString.fromText('GenOffice'))
  setVisualSignatureMetadata(annot, d.formFieldName)
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
}

/** Epoch ms → PDF date string, e.g. D:20260812175959+08'00' */
function pdfDateString(ms: number): string {
  const d = new Date(ms)
  const p = (n: number) => String(n).padStart(2, '0')
  const off = -d.getTimezoneOffset()
  const sign = off >= 0 ? '+' : '-'
  const abs = Math.abs(off)
  return (
    `D:${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}` +
    `${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}` +
    `${sign}${p(Math.floor(abs / 60))}'${p(abs % 60)}'`
  )
}

const NOTE_RECT_TOL = 2

/** Top-left anchors within tolerance — the only note identity that survives pdf.js
    resizing AP-less Text annot rects to its default icon size */
const noteRectsClose = (a: readonly number[], b: readonly number[]): boolean =>
  Math.abs(Math.min(a[0]!, a[2]!) - Math.min(b[0]!, b[2]!)) <= NOTE_RECT_TOL &&
  Math.abs(Math.max(a[1]!, a[3]!) - Math.max(b[1]!, b[3]!)) <= NOTE_RECT_TOL

/** Locate the saved Text annotation a reply points at (/IRT target). Earlier pdfium
    stages may have renumbered objects, so candidates match by rect + contents; the
    object-number hint only breaks ties between identical-looking notes. */
function findNoteAnnotRef(
  pdfDoc: PDFDocument,
  page: PDFPage,
  target: NoteReplyTarget,
): PDFRef | null {
  const annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) return null
  const matches: PDFRef[] = []
  for (let i = 0; i < annots.size(); i++) {
    const ref = annots.get(i)
    if (!(ref instanceof PDFRef)) continue
    const dict = pdfDoc.context.lookupMaybe(ref, PDFDict)
    if (!dict || dict.lookupMaybe(PDFName.of('Subtype'), PDFName) !== PDFName.of('Text')) continue
    const rectArr = dict.lookupMaybe(PDFName.of('Rect'), PDFArray)
    if (!rectArr || rectArr.size() !== 4) continue
    const rect: number[] = []
    for (let j = 0; j < rectArr.size(); j++) {
      const v = rectArr.lookupMaybe(j, PDFNumber)
      if (v) rect.push(v.asNumber())
    }
    if (rect.length !== 4 || !noteRectsClose(rect, target.rect)) continue
    const contents = dict.lookup(PDFName.of('Contents'))
    const text =
      contents instanceof PDFString || contents instanceof PDFHexString ? contents.decodeText() : ''
    if (text !== target.contents) continue
    if (ref.objectNumber === target.objNum) return ref
    matches.push(ref)
  }
  return matches[0] ?? null
}

/** Drawing annots: hand-written AP for Ink/Square/Circle/Line; notes are standard Text annots (viewer draws the icon) */
function addDrawing(
  pdfDoc: PDFDocument,
  page: PDFPage,
  d: DrawingInput,
  /** localId → registered ref of notes written earlier in this request (reply parenting) */
  noteRefs?: Map<string, PDFRef>,
): void {
  if (d.kind === 'image') return // handled by addImageStamp (needs async embed)
  const [r, g, b] = d.color

  if (d.kind === 'note') {
    const [x, y] = d.at
    const annot = pdfDoc.context.obj({
      Type: 'Annot',
      Subtype: 'Text',
      Rect: [num(x), num(y - 18), num(x + 20), num(y)],
      Name: 'Comment',
      C: d.color,
      F: 4,
      P: page.ref,
    })
    annot.set(PDFName.of('Contents'), PDFHexString.fromText(d.contents))
    annot.set(PDFName.of('T'), PDFHexString.fromText(d.author || 'GenOffice'))
    const when = pdfDateString(d.createdMs ?? Date.now())
    annot.set(PDFName.of('CreationDate'), PDFString.of(when))
    annot.set(PDFName.of('M'), PDFString.of(when))
    // Reply → standard /IRT chain (WPS/Acrobat comment threads). An unresolvable
    // parent degrades the reply to a root note instead of dropping the content.
    const parentRef = d.replyToLocalId
      ? (noteRefs?.get(d.replyToLocalId) ?? null)
      : d.replyToSaved
        ? findNoteAnnotRef(pdfDoc, page, d.replyToSaved)
        : null
    if (parentRef) {
      annot.set(PDFName.of('IRT'), parentRef)
      annot.set(PDFName.of('RT'), PDFName.of('R'))
    }
    const ref = pdfDoc.context.register(annot)
    if (d.localId) noteRefs?.set(d.localId, ref)
    appendAnnot(pdfDoc, page, ref)
    return
  }

  const ops: string[] = [`${num(d.width)} w 1 J 1 j ${r} ${g} ${b} RG`]
  let xs: number[] = []
  let ys: number[] = []
  let subtype: string

  if (d.kind === 'ink') {
    subtype = 'Ink'
    for (const path of d.paths) {
      if (path.length < 4) continue
      ops.push(`${num(path[0]!)} ${num(path[1]!)} m`)
      for (let i = 2; i < path.length; i += 2) ops.push(`${num(path[i]!)} ${num(path[i + 1]!)} l`)
      ops.push('S')
      for (let i = 0; i < path.length; i += 2) {
        xs.push(path[i]!)
        ys.push(path[i + 1]!)
      }
    }
  } else if (d.kind === 'rect' || d.kind === 'ellipse') {
    const [x1, y1, x2, y2] = d.rect
    subtype = d.kind === 'rect' ? 'Square' : 'Circle'
    if (d.kind === 'rect') ops.push(`${num(x1)} ${num(y1)} ${num(x2 - x1)} ${num(y2 - y1)} re S`)
    else ops.push(...ellipseOps(x1, y1, x2, y2))
    xs = [x1, x2]
    ys = [y1, y2]
  } else {
    const [fx, fy] = d.from
    const [tx, ty] = d.to
    subtype = 'Line'
    ops.push(`${num(fx)} ${num(fy)} m ${num(tx)} ${num(ty)} l S`)
    xs = [fx, tx]
    ys = [fy, ty]
    if (d.kind === 'arrow') {
      const ang = Math.atan2(ty - fy, tx - fx)
      const len = Math.max(9, d.width * 4.5)
      for (const off of [-0.45, 0.45]) {
        const hx = tx - len * Math.cos(ang + off)
        const hy = ty - len * Math.sin(ang + off)
        ops.push(`${num(tx)} ${num(ty)} m ${num(hx)} ${num(hy)} l S`)
        xs.push(hx)
        ys.push(hy)
      }
    }
  }

  const pad = d.width + 2
  const rect = [
    Math.min(...xs) - pad,
    Math.min(...ys) - pad,
    Math.max(...xs) + pad,
    Math.max(...ys) + pad,
  ]
  const ap = pdfDoc.context.stream(ops.join('\n'), { Type: 'XObject', Subtype: 'Form', BBox: rect })
  const annot = pdfDoc.context.obj({
    Type: 'Annot',
    Subtype: subtype,
    Rect: rect,
    C: d.color,
    F: 4,
    P: page.ref,
    BS: { W: d.width },
    AP: { N: pdfDoc.context.register(ap) },
  })
  if (d.kind === 'ink') annot.set(PDFName.of('InkList'), pdfDoc.context.obj(d.paths))
  if (d.kind === 'line' || d.kind === 'arrow') {
    annot.set(PDFName.of('L'), pdfDoc.context.obj([...d.from, ...d.to]))
  }
  annot.set(PDFName.of('T'), PDFHexString.fromText('GenOffice'))
  if (d.kind === 'ink') setVisualSignatureMetadata(annot, d.formFieldName)
  appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
}

function applyFormValues(pdfDoc: PDFDocument, values: FormValueInput[]): void {
  const form = pdfDoc.getForm()
  for (const v of values) {
    if (v.kind === 'text') {
      form.getTextField(v.name).setText(v.value ?? '')
    } else if (v.kind === 'radio') {
      const rg = form.getRadioGroup(v.name)
      if (v.value) rg.select(v.value)
      else rg.clear()
    } else if (v.kind === 'choice') {
      const f = form.getField(v.name)
      if (f instanceof PDFDropdown || f instanceof PDFOptionList) {
        if (v.value) f.select(v.value)
        else f.clear()
      }
    } else {
      const cb = form.getCheckBox(v.name)
      if (v.checked) cb.check()
      else cb.uncheck()
    }
  }
}

/** Extract the given pages (original indices) into bytes of a new PDF */
export async function extractPagesBytes(bytes: Uint8Array, pages: number[]): Promise<Uint8Array> {
  const src = await PDFDocument.load(bytes, { updateMetadata: false })
  const out = await PDFDocument.create()
  const valid = pages.filter((p) => p >= 0 && p < src.getPageCount())
  const copied = await out.copyPages(src, valid)
  for (const p of copied) out.addPage(p)
  return out.save({ useObjectStreams: false })
}

/** Insert all pages of another PDF after afterPageIndex (-1 = front); returns merged bytes and inserted page count */
export async function insertPdfBytes(
  bytes: Uint8Array,
  otherBytes: Uint8Array,
  afterPageIndex: number,
): Promise<{ merged: Uint8Array; count: number }> {
  const dst = await PDFDocument.load(bytes, { updateMetadata: false })
  const src = await PDFDocument.load(otherBytes, { updateMetadata: false })
  const copied = await dst.copyPages(src, src.getPageIndices())
  let at = Math.min(Math.max(afterPageIndex + 1, 0), dst.getPageCount())
  for (const p of copied) dst.insertPage(at++, p)
  return { merged: await dst.save({ useObjectStreams: false }), count: copied.length }
}

/** Insert a blank page after afterPageIndex (-1 = front), sized like the neighboring page */
export async function insertBlankPageBytes(
  bytes: Uint8Array,
  afterPageIndex: number,
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const at = Math.min(Math.max(afterPageIndex + 1, 0), doc.getPageCount())
  const ref = doc.getPage(Math.min(Math.max(afterPageIndex, 0), doc.getPageCount() - 1))
  const page = doc.insertPage(at, [ref.getWidth(), ref.getHeight()])
  // Match the neighbor's /Rotate too, or the blank page displays sideways next to it
  page.setRotation(ref.getRotation())
  return doc.save({ useObjectStreams: false })
}

/** Split into consecutive chunks of chunkSize pages, each becoming its own PDF (in page order) */
export async function splitPdfBytes(bytes: Uint8Array, chunkSize: number): Promise<Uint8Array[]> {
  const src = await PDFDocument.load(bytes, { updateMetadata: false })
  const total = src.getPageCount()
  const size = Math.max(1, Math.floor(chunkSize))
  const out: Uint8Array[] = []
  for (let start = 0; start < total; start += size) {
    const part = await PDFDocument.create()
    const count = Math.min(size, total - start)
    const copied = await part.copyPages(
      src,
      Array.from({ length: count }, (_, i) => start + i),
    )
    for (const p of copied) part.addPage(p)
    out.push(await part.save({ useObjectStreams: false }))
  }
  return out
}

export interface MergePagesOptions {
  /** Pages per sheet, 2-16 (WPS-style free count) */
  perSheet: number
  /** Fill order: horizontal = left→right then down, vertical = top→bottom then right */
  direction: 'horizontal' | 'vertical'
  /** Draw hairline separators on the internal cell boundaries */
  separator: boolean
}

/** Grid shape for an N-up sheet: 2-up is a side-by-side pair, otherwise near-square */
export function mergeGrid(perSheet: number): { cols: number; rows: number } {
  const n = Math.min(Math.max(Math.floor(perSheet), 2), 16)
  if (n === 2) return { cols: 2, rows: 1 }
  const cols = Math.ceil(Math.sqrt(n))
  return { cols, rows: Math.ceil(n / cols) }
}

/**
 * N-up imposition: place every perSheet consecutive pages onto one sheet, each
 * scaled to fit its cell and centered. The sheet matches the first page's size;
 * 2-up swaps width/height so two portrait pages sit side by side.
 */
export async function mergePagesBytes(
  bytes: Uint8Array,
  options: MergePagesOptions,
): Promise<Uint8Array> {
  const perSheet = Math.min(Math.max(Math.floor(options.perSheet), 2), 16)
  const src = await PDFDocument.load(bytes, { updateMetadata: false })
  const out = await PDFDocument.create()
  const total = src.getPageCount()
  const first = src.getPage(0)
  const { cols, rows } = mergeGrid(perSheet)
  const sheetW = perSheet === 2 ? first.getHeight() : first.getWidth()
  const sheetH = perSheet === 2 ? first.getWidth() : first.getHeight()
  // embedPages throws on pages without a content stream (e.g. our own inserted
  // blank pages) — give those an empty stream so they embed as empty cells
  for (const p of src.getPages()) {
    if (!p.node.Contents()) {
      p.node.set(PDFName.of('Contents'), src.context.register(src.context.stream('')))
    }
  }
  const embedded = await out.embedPages(src.getPages())
  const cellW = sheetW / cols
  const cellH = sheetH / rows
  for (let start = 0; start < total; start += perSheet) {
    const sheet = out.addPage([sheetW, sheetH])
    for (let i = 0; i < perSheet && start + i < total; i++) {
      const ep = embedded[start + i]!
      const scale = Math.min(cellW / ep.width, cellH / ep.height)
      const w = ep.width * scale
      const h = ep.height * scale
      const col = options.direction === 'vertical' ? Math.floor(i / rows) : i % cols
      const row = options.direction === 'vertical' ? i % rows : Math.floor(i / cols)
      sheet.drawPage(ep, {
        x: col * cellW + (cellW - w) / 2,
        // PDF y goes up: row 0 must land at the top of the sheet
        y: sheetH - (row + 1) * cellH + (cellH - h) / 2,
        width: w,
        height: h,
      })
    }
    if (options.separator) {
      // Document content (not UI chrome) — a fixed light gray like WPS's segment line
      const line = { thickness: 0.75, color: rgb(0.62, 0.62, 0.62) }
      for (let c = 1; c < cols; c++) {
        sheet.drawLine({
          start: { x: c * cellW, y: 0 },
          end: { x: c * cellW, y: sheetH },
          ...line,
        })
      }
      for (let r = 1; r < rows; r++) {
        sheet.drawLine({
          start: { x: 0, y: r * cellH },
          end: { x: sheetW, y: r * cellH },
          ...line,
        })
      }
    }
  }
  return out.save({ useObjectStreams: false })
}

/**
 * Replace the given pages (original indices) with all pages of another PDF,
 * inserted at the position of the first replaced page.
 */
export async function replacePagesBytes(
  bytes: Uint8Array,
  otherBytes: Uint8Array,
  pages: number[],
): Promise<{ merged: Uint8Array; removed: number; inserted: number }> {
  const dst = await PDFDocument.load(bytes, { updateMetadata: false })
  const src = await PDFDocument.load(otherBytes, { updateMetadata: false })
  const valid = [...new Set(pages.filter((p) => p >= 0 && p < dst.getPageCount()))].sort(
    (a, b) => a - b,
  )
  if (valid.length === 0) throw new Error('replacePages: no valid pages to replace')
  const at = valid[0]!
  for (const p of [...valid].reverse()) dst.removePage(p)
  const copied = await dst.copyPages(src, src.getPageIndices())
  let i = Math.min(at, dst.getPageCount())
  for (const p of copied) dst.insertPage(i++, p)
  return {
    merged: await dst.save({ useObjectStreams: false }),
    removed: valid.length,
    inserted: copied.length,
  }
}

/**
 * Resize every page to the target paper size (points): content and annotations
 * scale uniformly to fit. Centering is done by shifting the MediaBox origin
 * instead of translating the content, so annotations stay aligned with the
 * content they belong to. Pages displayed sideways (/Rotate 90 or 270) get the
 * swapped target so their displayed size matches the chosen paper.
 */
export async function setPageSizeBytes(
  bytes: Uint8Array,
  targetW: number,
  targetH: number,
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  for (const page of doc.getPages()) {
    const rot = ((page.getRotation().angle % 360) + 360) % 360
    const tw = rot === 90 || rot === 270 ? targetH : targetW
    const th = rot === 90 || rot === 270 ? targetW : targetH
    const { x, y, width: w, height: h } = page.getMediaBox()
    if (w === tw && h === th) continue
    const k = Math.min(tw / w, th / h)
    page.scaleContent(k, k)
    page.scaleAnnotations(k, k)
    // Both scale about the user-space origin; frame the scaled region centered
    const bx = x * k - (tw - w * k) / 2
    const by = y * k - (th - h * k) / 2
    page.setMediaBox(bx, by, tw, th)
    page.setCropBox(bx, by, tw, th)
  }
  return doc.save({ useObjectStreams: false })
}

/** Map a fractional rect of the displayed page (after /Rotate, y down from the top) to user space */
function displayFracToUserRect(
  rot: number,
  box: { x: number; y: number; width: number; height: number },
  l: number,
  t: number,
  r: number,
  b: number,
): { x: number; y: number; w: number; h: number } {
  const { x: x0, y: y0, width: W, height: H } = box
  if (rot === 90) {
    // display x runs along user +y, display y along user +x
    return { x: x0 + t * W, y: y0 + l * H, w: (b - t) * W, h: (r - l) * H }
  }
  if (rot === 180) {
    return { x: x0 + (1 - r) * W, y: y0 + t * H, w: (r - l) * W, h: (b - t) * H }
  }
  if (rot === 270) {
    return { x: x0 + (1 - b) * W, y: y0 + (1 - r) * H, w: (b - t) * W, h: (r - l) * H }
  }
  return { x: x0 + l * W, y: y0 + (1 - b) * H, w: (r - l) * W, h: (b - t) * H }
}

/**
 * Split every page into a perPage grid of pages (left→right, top→bottom as
 * displayed), the inverse of mergePagesBytes: each cell becomes its own page via
 * a copy with a tightened MediaBox/CropBox, so content is preserved losslessly.
 * The grid is laid out on the displayed page, then mapped through /Rotate.
 */
export async function splitPagesBytes(bytes: Uint8Array, perPage: 2 | 4 | 9): Promise<Uint8Array> {
  const src = await PDFDocument.load(bytes, { updateMetadata: false })
  const out = await PDFDocument.create()
  const { cols, rows } = mergeGrid(perPage)
  const total = src.getPageCount()
  for (let i = 0; i < total; i++) {
    const copies = await out.copyPages(
      src,
      Array.from({ length: perPage }, () => i),
    )
    for (let c = 0; c < perPage; c++) {
      const page = copies[c]!
      const rot = ((page.getRotation().angle % 360) + 360) % 360
      const col = c % cols
      const row = Math.floor(c / cols)
      const rect = displayFracToUserRect(
        rot,
        page.getCropBox(),
        col / cols,
        row / rows,
        (col + 1) / cols,
        (row + 1) / rows,
      )
      page.setMediaBox(rect.x, rect.y, rect.w, rect.h)
      page.setCropBox(rect.x, rect.y, rect.w, rect.h)
      out.addPage(page)
    }
  }
  return out.save({ useObjectStreams: false })
}

/** Crop rectangle as fractions of the displayed page (after /Rotate), y down from the top */
export interface CropFractionsRect {
  l: number
  t: number
  r: number
  b: number
}

/**
 * Shrink the CropBox of the given pages to the fractional rect. The fractions are
 * relative to the page as displayed (rotation applied, y down), so the same rect
 * lands on the same visual region regardless of each page's /Rotate value.
 */
export async function cropPagesBytes(
  bytes: Uint8Array,
  pages: number[],
  frac: CropFractionsRect,
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const l = Math.min(Math.max(frac.l, 0), 1)
  const t = Math.min(Math.max(frac.t, 0), 1)
  const r = Math.min(Math.max(frac.r, l), 1)
  const b = Math.min(Math.max(frac.b, t), 1)
  if (r - l <= 0 || b - t <= 0) throw new Error('cropPages: empty crop rect')
  for (const idx of pages) {
    if (idx < 0 || idx >= doc.getPageCount()) continue
    const page = doc.getPage(idx)
    const rot = ((page.getRotation().angle % 360) + 360) % 360
    const rect = displayFracToUserRect(rot, page.getCropBox(), l, t, r, b)
    page.setCropBox(rect.x, rect.y, rect.w, rect.h)
  }
  return doc.save({ useObjectStreams: false })
}

/** Append all pages of the other PDFs (in the given order) to the first; returns combined bytes and appended page count */
export async function mergePdfBytes(
  first: Uint8Array,
  others: Uint8Array[],
): Promise<{ merged: Uint8Array; appended: number }> {
  const dst = await PDFDocument.load(first, { updateMetadata: false })
  let appended = 0
  for (const otherBytes of others) {
    const src = await PDFDocument.load(otherBytes, { updateMetadata: false })
    const copied = await dst.copyPages(src, src.getPageIndices())
    for (const p of copied) dst.addPage(p)
    appended += copied.length
  }
  return { merged: await dst.save({ useObjectStreams: false }), appended }
}

function applyMetadata(pdfDoc: PDFDocument, meta: MetadataInput): void {
  if (meta.title !== undefined) pdfDoc.setTitle(meta.title)
  if (meta.author !== undefined) pdfDoc.setAuthor(meta.author)
  if (meta.subject !== undefined) pdfDoc.setSubject(meta.subject)
  if (meta.keywords !== undefined) {
    pdfDoc.setKeywords(
      meta.keywords
        .split(/[,，;；]/)
        .map((k) => k.trim())
        .filter(Boolean),
    )
  }
  pdfDoc.setModificationDate(new Date())
}

/**
 * Apply the request to the PDF at sourcePath and atomically write the result to targetPath
 * (temp file next to the target + rename, so a mid-write crash can't corrupt it).
 * The source file is only ever read: Save As (targetPath !== sourcePath) must never mutate
 * the original document, and a failed or cancelled save leaves both paths untouched.
 * In-place Save passes targetPath === sourcePath.
 * Returns the text edits that no longer matched the document and were skipped.
 */
export interface SavePdfSkips {
  skippedTextEdits: TextEditFailure[]
  skippedTextInserts: TextInsertFailure[]
  skippedImageEdits: ImageEditFailure[]
}

/** Original page index → index in the saved file (after this request's deletions/reorder);
    null = the page is gone from the output */
function finalPageIndex(request: SavePdfRequest, p: number): number | null {
  if (request.pageOrder) {
    const i = request.pageOrder.indexOf(p)
    return i >= 0 ? i : null
  }
  const del = request.deletedPages ?? []
  if (del.includes(p)) return null
  return p - del.filter((d) => d < p).length
}

/**
 * Read-back verification of applied content-stream edits against the final bytes.
 * Anything that fails here would have been silent data loss; the caller aborts the
 * save before the bytes reach disk, keeping the original file and the pending edits.
 */
async function verifyContentEdits(
  bytes: Uint8Array,
  request: SavePdfRequest,
  skips: SavePdfSkips,
): Promise<void> {
  const failures: { pageIndex: number; reason: string }[] = []
  const appliedText = (request.textEdits ?? []).filter(
    (e) =>
      !skips.skippedTextEdits.some((s) => s.pageIndex === e.pageIndex && s.oldText === e.oldText),
  )
  if (appliedText.length > 0) {
    const { verifyTextEdits } = await import('./text-edit')
    const remapped = appliedText.flatMap((e) => {
      const pageIndex = finalPageIndex(request, e.pageIndex)
      return pageIndex === null ? [] : [{ pageIndex, newText: e.newText }]
    })
    failures.push(...(await verifyTextEdits(bytes, remapped)))
  }
  const appliedInserts = (request.textInserts ?? []).filter(
    (_insert, editIndex) =>
      !skips.skippedTextInserts.some((skipped) => skipped.editIndex === editIndex),
  )
  if (appliedInserts.length > 0) {
    const { verifyTextEdits } = await import('./text-edit')
    const remapped = appliedInserts.flatMap((insert) => {
      const pageIndex = finalPageIndex(request, insert.pageIndex)
      return pageIndex === null ? [] : [{ pageIndex, newText: insert.text }]
    })
    failures.push(...(await verifyTextEdits(bytes, remapped)))
  }
  const appliedImages = (request.imageEdits ?? []).filter(
    (e, i) => e.kind !== 'deleteImage' && !skips.skippedImageEdits.some((s) => s.editIndex === i),
  )
  if (appliedImages.length > 0) {
    const { verifyImageEdits } = await import('./image-edit')
    const remapped = appliedImages.flatMap((e) => {
      const pageIndex = finalPageIndex(request, e.pageIndex)
      return pageIndex === null || e.kind === 'deleteImage' ? [] : [{ pageIndex, rect: e.rect }]
    })
    failures.push(...(await verifyImageEdits(bytes, remapped)))
  }
  if (failures.length > 0) {
    const pages = [...new Set(failures.map((f) => f.pageIndex + 1))].sort((a, b) => a - b)
    // "save-verify-failed pages=…" is parsed by the renderer to localize the notice
    throw new Error(
      `save-verify-failed pages=${pages.join(',')}: ${failures[0]!.reason}; the file was not written`,
    )
  }
}

export async function savePdfToPath(
  sourcePath: string,
  targetPath: string,
  request: SavePdfRequest,
): Promise<SavePdfSkips> {
  const { bytes, ...skips } = await applySaveRequest(
    new Uint8Array(await readFile(sourcePath)),
    request,
  )
  await verifyContentEdits(bytes, request, skips)
  const tmp = `${targetPath}.gensave-${process.pid}.tmp`
  try {
    await writeFile(tmp, bytes)
    await rename(tmp, targetPath)
  } catch (err) {
    await rm(tmp, { force: true })
    throw err
  }
  return skips
}

export interface AppliedSaveRequest {
  bytes: Uint8Array
  /** Text edits that could not be matched to the document; the rest of the request is in `bytes` */
  skippedTextEdits: TextEditFailure[]
  skippedTextInserts: TextInsertFailure[]
  /** Same, for content-stream image operations */
  skippedImageEdits: ImageEditFailure[]
}

/** Apply markups + form values + page ops, returning new bytes. Original objects are not reordered (pdf-lib keeps untouched objects). */
export async function applySaveRequest(
  bytes: Uint8Array,
  request: SavePdfRequest,
): Promise<AppliedSaveRequest> {
  let skippedTextEdits: TextEditFailure[] = []
  let skippedTextInserts: TextInsertFailure[] = []
  let skippedImageEdits: ImageEditFailure[] = []
  if (request.annotDeletes && request.annotDeletes.length > 0) {
    // First stage: the object numbers address the on-disk bytes; later pdfium
    // rewrites (text/image edits) may renumber objects
    const { applyAnnotDeletes } = await import('./annot-delete')
    bytes = await applyAnnotDeletes(bytes, request.annotDeletes)
  }
  if (request.textEdits && request.textEdits.length > 0) {
    // Content-stream rewrite must land before pdf-lib touches the bytes: everything
    // below annotates on top of whatever the pages now say
    const { applyTextEdits } = await import('./text-edit')
    const applied = await applyTextEdits(bytes, request.textEdits)
    bytes = applied.bytes
    skippedTextEdits = applied.skipped
  }
  if (request.textInserts && request.textInserts.length > 0) {
    const { applyTextInserts } = await import('./text-edit')
    const applied = await applyTextInserts(bytes, request.textInserts)
    bytes = applied.bytes
    skippedTextInserts = applied.skipped
  }
  if (request.imageEdits && request.imageEdits.length > 0) {
    const { applyImageEdits } = await import('./image-edit')
    const applied = await applyImageEdits(bytes, request.imageEdits)
    bytes = applied.bytes
    skippedImageEdits = applied.skipped
  }
  const pdfDoc = await PDFDocument.load(bytes, { updateMetadata: false })
  if (request.formValues.length > 0) applyFormValues(pdfDoc, request.formValues)
  const pages = pdfDoc.getPages()
  // Apply rotations first so markup appearances draw lines for the page's final orientation
  for (const r of request.rotations ?? []) {
    const page = pages[r.pageIndex]
    if (page) page.setRotation(degrees((page.getRotation().angle + r.delta) % 360))
  }
  for (const m of request.markups) {
    const page = pages[m.pageIndex]
    if (page) addMarkup(pdfDoc, page, m)
  }
  const noteRefs = new Map<string, PDFRef>()
  for (const d of request.drawings ?? []) {
    const page = pages[d.pageIndex]
    if (!page) continue
    if (d.kind === 'image') await addImageStamp(pdfDoc, page, d)
    else addDrawing(pdfDoc, page, d, noteRefs)
  }
  // Note content edits go after the drawings: replies added above locate their /IRT
  // parent by its old contents, which an earlier in-place rewrite would break. An
  // unmatched edit is a silent no-op (same degradation as an unresolvable reply).
  for (const e of request.noteEdits ?? []) {
    const page = pages[e.pageIndex]
    if (!page) continue
    const ref = findNoteAnnotRef(pdfDoc, page, {
      objNum: e.objNum,
      rect: e.rect,
      contents: e.oldContents,
    })
    const dict = ref ? pdfDoc.context.lookupMaybe(ref, PDFDict) : null
    if (!dict) continue
    dict.set(PDFName.of('Contents'), PDFHexString.fromText(e.contents))
    dict.set(PDFName.of('M'), PDFString.of(pdfDateString(Date.now())))
  }
  for (const s of request.stamps ?? []) {
    const page = pages[s.pageIndex]
    if (!page) continue
    const png = await pdfDoc.embedPng(s.image)
    const [x1, y1, x2, y2] = s.rect
    page.drawImage(png, {
      x: x1,
      y: y1,
      width: x2 - x1,
      height: y2 - y1,
      opacity: s.opacity ?? 1,
    })
  }
  if (request.metadata) applyMetadata(pdfDoc, request.metadata)
  // Deletions go last, in descending order; earlier ops all address original page indices
  for (const idx of [...(request.deletedPages ?? [])].sort((a, b) => b - a)) {
    if (idx >= 0 && idx < pdfDoc.getPageCount() && pdfDoc.getPageCount() > 1) pdfDoc.removePage(idx)
  }
  // Reorder last: pageOrder gives the new order of remaining-after-delete pages by original index.
  // pdf-lib's removePage never invalidates its page cache, so getPages() here would return the
  // stale pre-deletion list — derive the surviving pages from the pre-deletion snapshot instead.
  const order = request.pageOrder
  if (order && order.length > 0) {
    const deletedSet = new Set(request.deletedPages ?? [])
    const target = order
      .filter((o) => !deletedSet.has(o))
      .map((o) => pages[o])
      .filter((p) => p !== undefined)
    if (target.length === pdfDoc.getPageCount()) {
      while (pdfDoc.getPageCount() > 0) pdfDoc.removePage(0)
      for (const p of target) pdfDoc.addPage(p)
    }
  }
  const staticFormFills = resultingStaticFormFills(request, pages.length)
  if (staticFormFills !== undefined) {
    if (staticFormFills.length === 0) pdfDoc.catalog.delete(STATIC_FORM_FILLS_KEY)
    else
      pdfDoc.catalog.set(
        STATIC_FORM_FILLS_KEY,
        PDFHexString.fromText(JSON.stringify(staticFormFills)),
      )
  }
  try {
    return {
      bytes: await pdfDoc.save({ useObjectStreams: false }),
      skippedTextEdits,
      skippedTextInserts,
      skippedImageEdits,
    }
  } catch (err) {
    // Form values beyond WinAnsi (e.g. CJK) make pdf-lib's appearance generation fail:
    // skip it and set NeedAppearances so viewers rebuild them (Acrobat/pdfjs both support this)
    if (request.formValues.length === 0) throw err
    pdfDoc.getForm().acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
    return {
      bytes: await pdfDoc.save({ useObjectStreams: false, updateFieldAppearances: false }),
      skippedTextEdits,
      skippedTextInserts,
      skippedImageEdits,
    }
  }
}
