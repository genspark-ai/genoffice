import { readFile, rename, rm, writeFile } from 'node:fs/promises'
import {
  PDFArray,
  PDFBool,
  PDFDocument,
  PDFDropdown,
  PDFHexString,
  PDFName,
  PDFOptionList,
  degrees,
} from 'pdf-lib'
import type { PDFPage, PDFRef } from 'pdf-lib'
import type {
  DrawingInput,
  FormValueInput,
  MarkupInput,
  MetadataInput,
  SavePdfRequest,
} from '../shared/ipc'

const num = (v: number) => Math.round(v * 100) / 100

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

/** Drawing annots: hand-written AP for Ink/Square/Circle/Line; notes are standard Text annots (viewer draws the icon) */
function addDrawing(pdfDoc: PDFDocument, page: PDFPage, d: DrawingInput): void {
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
    annot.set(PDFName.of('T'), PDFHexString.fromText('GenOffice'))
    appendAnnot(pdfDoc, page, pdfDoc.context.register(annot))
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
 */
export async function savePdfToPath(
  sourcePath: string,
  targetPath: string,
  request: SavePdfRequest,
): Promise<void> {
  const bytes = await applySaveRequest(new Uint8Array(await readFile(sourcePath)), request)
  const tmp = `${targetPath}.gensave-${process.pid}.tmp`
  try {
    await writeFile(tmp, bytes)
    await rename(tmp, targetPath)
  } catch (err) {
    await rm(tmp, { force: true })
    throw err
  }
}

/** Apply markups + form values + page ops, returning new bytes. Original objects are not reordered (pdf-lib keeps untouched objects). */
export async function applySaveRequest(
  bytes: Uint8Array,
  request: SavePdfRequest,
): Promise<Uint8Array> {
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
  for (const d of request.drawings ?? []) {
    const page = pages[d.pageIndex]
    if (page) addDrawing(pdfDoc, page, d)
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
  try {
    return await pdfDoc.save({ useObjectStreams: false })
  } catch (err) {
    // Form values beyond WinAnsi (e.g. CJK) make pdf-lib's appearance generation fail:
    // skip it and set NeedAppearances so viewers rebuild them (Acrobat/pdfjs both support this)
    if (request.formValues.length === 0) throw err
    pdfDoc.getForm().acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
    return await pdfDoc.save({ useObjectStreams: false, updateFieldAppearances: false })
  }
}
