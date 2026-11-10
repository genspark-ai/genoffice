import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFRawStream,
  PDFRef,
  PDFStream,
  decodePDFRawStream,
} from 'pdf-lib'
import type { Pdfium } from './text-edit'

/**
 * Text drawn inside Form XObjects (FPDI/TCPDF templates, pdf-lib embedded pages,
 * print-to-PDF wrappers) is invisible to the page-level object list. PDFium's
 * public API can remove an object from a form but not add one, so such text is
 * edited by lifting it out: the matched text object (and, for nested templates,
 * each form between it and the page) is removed from its parent and re-inserted
 * at page level right after that parent. Objects inside a form are parsed in the
 * form's own space (the parent's CTM at the Do is kept on the form object only),
 * so the lift concatenates the parent's matrix onto the lifted object; after that
 * the ordinary page-owned edit path applies unchanged. PDFium regenerates each
 * dirtied form stream on FPDFPage_GenerateContent.
 */

const FPDF_PAGEOBJ_FORM = 5

/** Form objects from the page's top-level form down to the object's immediate parent */
export type FormChain = readonly number[]

/** PDF matrix [a, b, c, d, e, f] */
export type Mat = readonly [number, number, number, number, number, number]

const IDENTITY: Mat = [1, 0, 0, 1, 0, 0]

/** Apply `a`, then `b` (pdfium's Concat order) */
export const concatMat = (a: Mat, b: Mat): Mat => [
  a[0] * b[0] + a[1] * b[2],
  a[0] * b[1] + a[1] * b[3],
  a[2] * b[0] + a[3] * b[2],
  a[2] * b[1] + a[3] * b[3],
  a[4] * b[0] + a[5] * b[2] + b[4],
  a[4] * b[1] + a[5] * b[3] + b[5],
]

export function transformRect(r: readonly [number, number, number, number], m: Mat) {
  const xs: number[] = []
  const ys: number[] = []
  for (const [x, y] of [
    [r[0], r[1]],
    [r[2], r[1]],
    [r[0], r[3]],
    [r[2], r[3]],
  ] as const) {
    xs.push(m[0] * x + m[2] * y + m[4])
    ys.push(m[1] * x + m[3] * y + m[5])
  }
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)] as [
    number,
    number,
    number,
    number,
  ]
}

function objMatrix(m: Pdfium, obj: number): Mat {
  const p = m._malloc(24)
  try {
    if (!m._FPDFPageObj_GetMatrix(obj, p)) return IDENTITY
    return Array.from(m.HEAPF32.subarray(p >> 2, (p >> 2) + 6)) as unknown as Mat
  } finally {
    m._free(p)
  }
}

export interface FormTextVisit {
  obj: number
  /** Empty for page-owned text */
  chain: FormChain
  /** Index of the top-level ancestor (or the object itself) in the page object list */
  topIndex: number
  /** Ordinals of each chain form among its parent's form children (Do order) */
  ordinals: readonly number[]
  /** Maps the object's own coordinate space to page space */
  toPage: Mat
}

/** Visit every page object of `objType`, descending into Form XObjects */
export function walkPageText(
  m: Pdfium,
  page: number,
  objType: number,
  visit: (v: FormTextVisit) => void,
): void {
  const walkForm = (
    form: number,
    chain: number[],
    ordinals: number[],
    topIndex: number,
    toPage: Mat,
    depth: number,
  ) => {
    if (depth > 8) return
    const n = m._FPDFFormObj_CountObjects(form)
    let formOrdinal = 0
    for (let i = 0; i < n; i++) {
      const o = m._FPDFFormObj_GetObject(form, i)
      const t = m._FPDFPageObj_GetType(o)
      if (t === objType) visit({ obj: o, chain, topIndex, ordinals, toPage })
      else if (t === FPDF_PAGEOBJ_FORM) {
        walkForm(
          o,
          [...chain, o],
          [...ordinals, formOrdinal],
          topIndex,
          concatMat(objMatrix(m, o), toPage),
          depth + 1,
        )
        formOrdinal++
      }
    }
  }
  const count = m._FPDFPage_CountObjects(page)
  let formOrdinal = 0
  for (let i = 0; i < count; i++) {
    const o = m._FPDFPage_GetObject(page, i)
    const t = m._FPDFPageObj_GetType(o)
    if (t === objType) visit({ obj: o, chain: [], topIndex: i, ordinals: [], toPage: IDENTITY })
    else if (t === FPDF_PAGEOBJ_FORM) {
      walkForm(o, [o], [formOrdinal], i, objMatrix(m, o), 1)
      formOrdinal++
    }
  }
}

export const SHARED_TEMPLATE_REASON = 'the text lives in a template that is drawn more than once'
export const SHARED_TEMPLATE_IMAGE_REASON =
  'the image lives in a template that is drawn more than once'

interface FormNode {
  /** How many times this form is painted across the document (through all parents) */
  draws: number
  children: FormNode[]
}

/** Per page, the top-level Form XObjects in content-stream Do order (null = unparsable) */
export type FormDrawTree = (FormNode[] | null)[]

const nameOf = (s: string) => PDFName.of(s)

function decodeStream(s: PDFStream): string | null {
  try {
    const bytes = s instanceof PDFRawStream ? decodePDFRawStream(s).decode() : s.getContents()
    return Buffer.from(bytes).toString('latin1')
  } catch {
    return null
  }
}

/** XObject names painted by a content stream, in order; false when it cannot be read */
function doNames(content: string | null): string[] | false {
  if (content === null) return false
  // Strip inline image data so its bytes cannot spell a stray " Do"
  const code = content.replace(/\bBI\b[\s\S]*?\bEI\b/g, ' ')
  return [...code.matchAll(/\/([^\s/[\]<>(){}%]+)\s+Do\b/g)].map((mm) => mm[1]!)
}

/** Rank each form by its total paint count; tree mirrors pdfium's form nesting per page */
export async function formDrawTree(bytes: Uint8Array): Promise<FormDrawTree> {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false })
  const ctx = doc.context
  const draws = new Map<string, number>()
  // Form refs per form in Do order; false when its content cannot be read
  const kids = new Map<string, PDFRef[] | false>()

  const formKids = (resources: PDFDict | undefined, content: string | null): PDFRef[] | false => {
    const names = doNames(content)
    if (names === false) return false
    const xobjs = resources?.lookupMaybe(nameOf('XObject'), PDFDict)
    const out: PDFRef[] = []
    for (const n of names) {
      const ref = xobjs?.get(nameOf(n))
      if (!(ref instanceof PDFRef)) continue
      const s = ctx.lookup(ref)
      if (s instanceof PDFStream && s.dict.get(nameOf('Subtype')) === nameOf('Form')) out.push(ref)
    }
    return out
  }
  const kidsOfForm = (ref: PDFRef): PDFRef[] | false => {
    const key = ref.toString()
    const cached = kids.get(key)
    if (cached !== undefined) return cached
    const s = ctx.lookup(ref) as PDFStream
    const res = s.dict.lookupMaybe(nameOf('Resources'), PDFDict)
    const k = formKids(res, decodeStream(s))
    kids.set(key, k)
    return k
  }
  const paint = (ref: PDFRef, times: number, path: Set<string>) => {
    const key = ref.toString()
    draws.set(key, (draws.get(key) ?? 0) + times)
    if (path.has(key)) return
    const k = kidsOfForm(ref)
    if (!k) return
    const next = new Set(path).add(key)
    for (const c of k) paint(c, times, next)
  }

  const pageKids: (PDFRef[] | false)[] = doc.getPages().map((p) => {
    const contents = p.node.Contents()
    let content: string | null = ''
    const parts =
      contents instanceof PDFArray
        ? contents.asArray().map((r) => ctx.lookup(r))
        : contents
          ? [contents]
          : []
    for (const part of parts) {
      const dec = part instanceof PDFStream ? decodeStream(part) : null
      if (dec === null) {
        content = null
        break
      }
      content += `${dec}\n`
    }
    return formKids(p.node.Resources(), content)
  })
  for (const k of pageKids) if (k) for (const r of k) paint(r, 1, new Set())

  const build = (ref: PDFRef, path: Set<string>): FormNode => {
    const key = ref.toString()
    const k = kidsOfForm(ref)
    const next = new Set(path).add(key)
    return {
      draws: k === false ? Infinity : (draws.get(key) ?? 1),
      children: k && !path.has(key) ? k.map((c) => build(c, next)) : [],
    }
  }
  return pageKids.map((k) => (k ? k.map((r) => build(r, new Set())) : null))
}

/** True when every form on the chain is painted exactly once, so editing inside it
    cannot change another page (or another copy on this page). Any mismatch between
    the parsed tree and pdfium's nesting is treated as shared. */
export function chainIsExclusive(
  tree: FormDrawTree,
  pageIndex: number,
  ordinals: readonly number[],
): boolean {
  let level = tree[pageIndex] ?? null
  for (const k of ordinals) {
    const node = level?.[k]
    if (!node || node.draws !== 1) return false
    level = node.children
  }
  return true
}

/** Move `obj` (and each form between it and the page) out to page level, keeping
    stream order and geometry: each lifted object lands right after its former
    parent and takes on the parent's (by then page-space) matrix. */
export function liftOutOfForms(m: Pdfium, page: number, obj: number, chain: FormChain): void {
  const indexOf = (target: number): number => {
    const n = m._FPDFPage_CountObjects(page)
    for (let i = 0; i < n; i++) if (m._FPDFPage_GetObject(page, i) === target) return i
    return -1
  }
  const lift = (child: number, parent: number) => {
    if (!m._FPDFFormObj_RemoveObject(parent, child)) {
      throw new Error('could not detach the object from its template')
    }
    const at = indexOf(parent)
    if (at >= 0 && m._FPDFPage_InsertObjectAtIndex)
      m._FPDFPage_InsertObjectAtIndex(page, child, at + 1)
    else m._FPDFPage_InsertObject(page, child)
    const [a, b, c, d, e, f] = objMatrix(m, parent)
    m._FPDFPageObj_Transform(child, a, b, c, d, e, f)
  }
  // Top-down: once a nested form sits on the page, its own children can follow
  for (let i = 1; i < chain.length; i++) {
    if (indexOf(chain[i]!) < 0) lift(chain[i]!, chain[i - 1]!)
  }
  const parent = chain[chain.length - 1]
  if (parent !== undefined && indexOf(obj) < 0) lift(obj, parent)
}
