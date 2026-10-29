import { describe, expect, it } from 'vitest'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFRawStream,
  decodePDFRawStream,
  degrees,
} from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { applyAnnotDeletes } from '../src/main/annot-delete'
import { applySaveRequest } from '../src/main/save-pdf'
import { planEditOps, type OpContext } from '../src/renderer/edit-ops'
import type { AnnotDeleteInput, SavePdfRequest } from '../src/shared/ipc'

const request = (over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

async function makePdf(rot = 0): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  doc.addPage([612, 792]).setRotation(degrees(rot))
  return doc.save({ useObjectStreams: false })
}

function pageAnnots(doc: PDFDocument, pageIndex: number): PDFDict[] {
  const annots = doc.getPage(pageIndex).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) return []
  return Array.from({ length: annots.size() }, (_, i) => annots.lookup(i, PDFDict))
}

const subtypeOf = (annot: PDFDict) => annot.lookup(PDFName.of('Subtype'), PDFName).decodeText()

const numberArray = (annot: PDFDict, key: string): number[] => {
  const arr = annot.lookup(PDFName.of(key), PDFArray)
  return Array.from({ length: arr.size() }, (_, i) => arr.lookup(i, PDFNumber).asNumber())
}

const annotObjNum = (doc: PDFDocument, pageIndex: number, index: number): number => {
  const annots = doc.getPage(pageIndex).node.lookup(PDFName.of('Annots'), PDFArray)
  return (annots.get(index) as { objectNumber: number }).objectNumber
}

const toDelete = (subtype: AnnotDeleteInput['subtype'], rect: number[]): AnnotDeleteInput => ({
  pageIndex: 0,
  objNum: 0,
  subtype,
  rect: rect as [number, number, number, number],
})

// ── edit-op validation (planned failures, same shape as edit-ops.test.ts) ──

const ctx = (): OpContext => ({
  readOnly: false,
  pageCount: 3,
  deleted: new Set(),
  claimedImages: new Set(),
})

let n = 0
const newId = () => `id${++n}`

const plan = (drawing: Record<string, unknown>) =>
  planEditOps([{ op: 'addDrawing', drawing }], ctx(), newId)

describe('squiggly markups', () => {
  const SQUIG_QUADS = [[100, 716, 200, 716, 100, 698, 200, 698]]

  it('writes a Squiggly annotation pdf.js reads back with quads and color', async () => {
    const bytes = await makePdf()
    const { bytes: saved } = await applySaveRequest(
      bytes,
      request({
        markups: [
          { pageIndex: 0, type: 'squiggly', color: [0.13, 0.65, 0.35], quads: SQUIG_QUADS },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Squiggly'])
    // Same identity fields as every other quad markup (print flag, author, color)
    expect(annots[0]!.lookup(PDFName.of('F'), PDFNumber).asNumber()).toBe(4)
    expect(numberArray(annots[0]!, 'C')).toEqual([0.13, 0.65, 0.35])

    const loadingTask = getDocument({ data: saved.slice() })
    try {
      const pdfJsDoc = await loadingTask.promise
      const annos = await (await pdfJsDoc.getPage(1)).getAnnotations()
      const squiggles = annos.filter((a) => a.subtype === 'Squiggly')
      expect(squiggles).toHaveLength(1)
      const a = squiggles[0]!
      // pdf.js normalizes each quad to [minX,maxY,maxX,maxY,minX,minY,maxX,minY]
      expect(Array.from(a.quadPoints as Float32Array)).toEqual(SQUIG_QUADS[0])
      const expected = [0.13, 0.65, 0.35].map((c) => c * 255)
      for (let i = 0; i < 3; i++) {
        expect(Math.abs((a.color as number[])[i]! - expected[i]!)).toBeLessThanOrEqual(1)
      }
    } finally {
      await loadingTask.destroy()
    }
  })

  it('keeps a hand-written appearance stream and is deletable like the other markups', async () => {
    const bytes = await makePdf()
    const { bytes: saved } = await applySaveRequest(
      bytes,
      request({
        markups: [
          { pageIndex: 0, type: 'squiggly', color: [0.13, 0.65, 0.35], quads: SQUIG_QUADS },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots).toHaveLength(1)
    expect(annots[0]!.has(PDFName.of('AP'))).toBe(true)

    const remaining = await applyAnnotDeletes(saved, [
      {
        ...toDelete('squiggly', [100, 698, 200, 716]),
        objNum: annotObjNum(out, 0, 0),
      },
    ])
    expect(pageAnnots(await PDFDocument.load(remaining), 0)).toEqual([])
  })

  it('keeps the wave inside the quad on rotated pages (90 and 180)', async () => {
    // The wave bulges from the visual bottom edge into the text; on a sideways or
    // upside-down page a wrong perpendicular pushes the curve out of the /AP BBox,
    // where viewers clip it. Every on-curve point (m starts and v endpoints) must
    // stay inside the quad bounds.
    for (const rot of [90, 180]) {
      const bytes = await makePdf(rot)
      const quad =
        rot === 90
          ? [100, 600, 112, 600, 100, 400, 112, 400]
          : [100, 600, 200, 600, 100, 588, 200, 588]
      const { bytes: saved } = await applySaveRequest(
        bytes,
        request({
          markups: [{ pageIndex: 0, type: 'squiggly', color: [0.13, 0.65, 0.35], quads: [quad] }],
        }),
      )
      const out = await PDFDocument.load(saved)
      const apRef = pageAnnots(out, 0)[0]!.lookup(PDFName.of('AP'), PDFDict).lookup(PDFName.of('N'))
      const ops = Buffer.from(
        decodePDFRawStream(out.context.lookup(apRef) as PDFRawStream).decode(),
      ).toString('latin1')
      // canonical quad [x1,yMax,x2,yMax,x1,yMin,x2,yMin] → bounds x1..x2, yMin..yMax
      const [x1, y1, x2, y2] = [quad[0]!, quad[5]!, quad[2]!, quad[1]!]
      const inside = (v: number, lo: number, hi: number) => v >= lo - 0.5 && v <= hi + 0.5
      for (const line of ops.split('\n')) {
        const parts = line.trim().split(/\s+/)
        if (parts.at(-1) === 'm') {
          expect(inside(Number(parts[0]), x1, x2), `${rot}: m x ${line}`).toBe(true)
          expect(inside(Number(parts[1]), y1, y2), `${rot}: m y ${line}`).toBe(true)
        } else if (parts.at(-1) === 'v') {
          // control points may leave the box (the curve does not); endpoints may not
          expect(inside(Number(parts[2]), x1, x2), `${rot}: v x ${line}`).toBe(true)
          expect(inside(Number(parts[3]), y1, y2), `${rot}: v y ${line}`).toBe(true)
        }
      }
    }
  })
})

describe('caret drawings', () => {
  it('writes a Caret annotation whose rect is centered on the point', async () => {
    const bytes = await makePdf()
    const { bytes: saved } = await applySaveRequest(
      bytes,
      request({
        drawings: [{ kind: 'caret', pageIndex: 0, color: [0.86, 0.22, 0.18], point: [150, 400] }],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Caret'])
    expect(numberArray(annots[0]!, 'Rect')).toEqual([140, 390, 160, 410])
    expect(annots[0]!.lookup(PDFName.of('F'), PDFNumber).asNumber()).toBe(4)
    // the appearance stream is the wedge a viewer draws: pdf.js has no
    // default appearance for a Caret, so without it the mark is in the file
    // and nothing draws it
    expect(annots[0]!.has(PDFName.of('AP'))).toBe(true)

    const loadingTask = getDocument({ data: saved.slice() })
    try {
      const pdfJsDoc = await loadingTask.promise
      const annos = await (await pdfJsDoc.getPage(1)).getAnnotations()
      expect(annos.filter((a) => a.subtype === 'Caret')).toHaveLength(1)
    } finally {
      await loadingTask.destroy()
    }
  })
})

describe('polygon drawings', () => {
  const VERTICES: [number, number][] = [
    [50, 100],
    [200, 100],
    [250, 200],
    [150, 300],
    [20, 180],
  ]

  it('writes a stroke-only Polygon annotation whose vertices round-trip exact', async () => {
    const bytes = await makePdf()
    const { bytes: saved } = await applySaveRequest(
      bytes,
      request({
        drawings: [
          { kind: 'polygon', pageIndex: 0, color: [0.17, 0.4, 1], width: 2, vertices: VERTICES },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Polygon'])
    // Rect is the vertex bbox, Vertices the exact flat corner list
    expect(numberArray(annots[0]!, 'Rect')).toEqual([20, 100, 250, 300])
    expect(numberArray(annots[0]!, 'Vertices')).toEqual(VERTICES.flat())
    expect(numberArray(annots[0]!, 'C')).toEqual([0.17, 0.4, 1])
    // Stroke-only: no interior color. The appearance stream is the shape's
    // own, so a viewer that does not synthesise one still draws the mark.
    expect(annots[0]!.has(PDFName.of('IC'))).toBe(false)
    expect(annots[0]!.has(PDFName.of('AP'))).toBe(true)

    const loadingTask = getDocument({ data: saved.slice() })
    try {
      const pdfJsDoc = await loadingTask.promise
      const annos = await (await pdfJsDoc.getPage(1)).getAnnotations()
      expect(annos.filter((a) => a.subtype === 'Polygon')).toHaveLength(1)
    } finally {
      await loadingTask.destroy()
    }
  })

  it('deletes a saved polygon by subtype through the annot-delete path', async () => {
    const bytes = await makePdf()
    const { bytes: saved } = await applySaveRequest(
      bytes,
      request({
        drawings: [
          { kind: 'polygon', pageIndex: 0, color: [0, 0, 0], width: 1, vertices: VERTICES },
          { kind: 'caret', pageIndex: 0, color: [0, 0, 0], point: [300, 500] },
        ],
      }),
    )
    const doc = await PDFDocument.load(saved)
    const remaining = await applyAnnotDeletes(saved, [
      {
        ...toDelete('polygon', [20, 100, 250, 300]),
        objNum: annotObjNum(doc, 0, 0),
      },
    ])
    expect(pageAnnots(await PDFDocument.load(remaining), 0).map(subtypeOf)).toEqual(['Caret'])
  })
})

describe('caret/polygon edit-op validation', () => {
  it('accepts a caret with a finite point and a polygon with 3+ vertices', () => {
    const caret = plan({
      kind: 'caret',
      pageIndex: 0,
      color: [0, 0, 0],
      point: [150, 400],
    })
    expect(caret.failures).toEqual([])
    const polygon = plan({
      kind: 'polygon',
      pageIndex: 0,
      color: [0, 0, 0],
      width: 1,
      vertices: [
        [0, 0],
        [10, 0],
        [5, 8],
      ],
    })
    expect(polygon.failures).toEqual([])
  })

  it('rejects a polygon with fewer than 3 vertices as a planned failure', () => {
    const result = plan({
      kind: 'polygon',
      pageIndex: 0,
      color: [0, 0, 0],
      width: 1,
      vertices: [
        [0, 0],
        [10, 10],
      ],
    })
    expect(result.failures).toHaveLength(1)
    expect(result.failures[0]!.error).toMatch(/at least 3 vertices/)
  })

  it('rejects a caret with a NaN or out-of-range point as a planned failure', () => {
    const nan = plan({ kind: 'caret', pageIndex: 0, color: [0, 0, 0], point: [NaN, 400] })
    expect(nan.failures).toHaveLength(1)
    expect(nan.failures[0]!.error).toMatch(/\[x, y\]/)
    const wild = plan({ kind: 'caret', pageIndex: 0, color: [0, 0, 0], point: [20000, 400] })
    expect(wild.failures).toHaveLength(1)
    expect(wild.failures[0]!.error).toMatch(/PDF coordinate range/)
  })
})
