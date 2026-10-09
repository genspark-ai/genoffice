import { createHash } from 'node:crypto'
import { mkdtempSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  PDFArray,
  PDFContentStream,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFString,
  PDFRawStream,
  PDFRef,
  decodePDFRawStream,
  degrees,
  rgb,
} from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import {
  applySaveRequest,
  cropPagesBytes,
  extractPagesBytes,
  insertBlankPageBytes,
  insertPdfBytes,
  mergeGrid,
  mergePagesBytes,
  mergePdfBytes,
  readStaticFormFills,
  replacePagesBytes,
  savePdfToPath,
  setPageSizeBytes,
  splitPagesBytes,
  splitPdfBytes,
} from '../src/main/save-pdf'
import { VISUAL_SIGNATURE_CONTENT_PREFIX } from '../src/shared/ipc'
import { PDFJS_ANNOT_TEXT } from '../src/renderer/note-threads'
import type { SavePdfRequest } from '../src/shared/ipc'

/** pdf.js AnnotationType.POPUP — a note's popup must never surface as its own entry */

/** 1x1 red pixel PNG */
const TINY_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

async function makePdf(sizes: [number, number][]): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  for (const size of sizes) doc.addPage(size)
  return doc.save({ useObjectStreams: false })
}

/** applySaveRequest with the skipped-text-edit channel unwrapped (none expected here) */
async function apply(bytes: Uint8Array, req: SavePdfRequest): Promise<Uint8Array> {
  const result = await applySaveRequest(bytes, req)
  expect(result.skippedTextEdits).toEqual([])
  return result.bytes
}

const request = (over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

function pageAnnots(doc: PDFDocument, pageIndex: number): PDFDict[] {
  const annots = doc.getPage(pageIndex).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) return []
  return Array.from({ length: annots.size() }, (_, i) => annots.lookup(i, PDFDict))
}

const subtypeOf = (annot: PDFDict) => annot.lookup(PDFName.of('Subtype'), PDFName).decodeText()

describe('extractPagesBytes', () => {
  it('extracts the requested pages in the given order', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const out = await PDFDocument.load(await extractPagesBytes(bytes, [2, 0]))
    expect(out.getPageCount()).toBe(2)
    expect(out.getPage(0).getWidth()).toBe(300)
    expect(out.getPage(1).getWidth()).toBe(100)
  })

  it('silently drops out-of-range page indices', async () => {
    const bytes = await makePdf([[100, 100]])
    const out = await PDFDocument.load(await extractPagesBytes(bytes, [-1, 0, 5]))
    expect(out.getPageCount()).toBe(1)
  })
})

describe('insertPdfBytes', () => {
  it('inserts all pages at the front when afterPageIndex is -1', async () => {
    const dst = await makePdf([[100, 100]])
    const src = await makePdf([
      [200, 200],
      [300, 300],
    ])
    const { merged, count } = await insertPdfBytes(dst, src, -1)
    expect(count).toBe(2)
    const out = await PDFDocument.load(merged)
    expect(out.getPageCount()).toBe(3)
    expect(out.getPage(0).getWidth()).toBe(200)
    expect(out.getPage(2).getWidth()).toBe(100)
  })

  it('clamps an out-of-range afterPageIndex to the document end', async () => {
    const dst = await makePdf([[100, 100]])
    const src = await makePdf([[200, 200]])
    const { merged } = await insertPdfBytes(dst, src, 99)
    const out = await PDFDocument.load(merged)
    expect(out.getPage(1).getWidth()).toBe(200)
  })
})

describe('insertBlankPageBytes', () => {
  it('inserts a blank page after the given page, sized like it', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 300],
    ])
    const out = await PDFDocument.load(await insertBlankPageBytes(bytes, 1))
    expect(out.getPageCount()).toBe(3)
    expect(out.getPage(2).getWidth()).toBe(200)
    expect(out.getPage(2).getHeight()).toBe(300)
  })

  it('inserts at the front sized like the first page when afterPageIndex is -1', async () => {
    const bytes = await makePdf([[150, 250]])
    const out = await PDFDocument.load(await insertBlankPageBytes(bytes, -1))
    expect(out.getPageCount()).toBe(2)
    expect(out.getPage(0).getWidth()).toBe(150)
    expect(out.getPage(0).getHeight()).toBe(250)
  })

  it('copies the neighbor /Rotate so the displayed orientation matches', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([100, 200]).setRotation(degrees(90))
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await PDFDocument.load(await insertBlankPageBytes(bytes, 0))
    expect(out.getPage(1).getRotation().angle).toBe(90)
  })
})

describe('splitPdfBytes', () => {
  it('splits into chunks of the given size, keeping page order', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
      [400, 400],
      [500, 500],
    ])
    const parts = await splitPdfBytes(bytes, 2)
    expect(parts.length).toBe(3)
    const first = await PDFDocument.load(parts[0]!)
    expect(first.getPageCount()).toBe(2)
    expect(first.getPage(0).getWidth()).toBe(100)
    const last = await PDFDocument.load(parts[2]!)
    expect(last.getPageCount()).toBe(1)
    expect(last.getPage(0).getWidth()).toBe(500)
  })

  it('clamps chunk size to at least one page per file', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
    ])
    const parts = await splitPdfBytes(bytes, 0)
    expect(parts.length).toBe(2)
  })
})

const mergeOpts = (perSheet: number) =>
  ({ perSheet, direction: 'horizontal', separator: false }) as const

function decodedPageContent(page: ReturnType<PDFDocument['getPage']>): string {
  const contents = page.node.Contents()
  if (contents instanceof PDFContentStream) {
    return Buffer.from(contents.getUnencodedContents()).toString('latin1')
  }
  if (contents instanceof PDFRawStream) {
    return Buffer.from(decodePDFRawStream(contents).decode()).toString('latin1')
  }
  if (contents instanceof PDFArray) {
    return Array.from({ length: contents.size() }, (_, index) => {
      const value = contents.lookup(index)
      if (value instanceof PDFContentStream) {
        return Buffer.from(value.getUnencodedContents()).toString('latin1')
      }
      return value instanceof PDFRawStream
        ? Buffer.from(decodePDFRawStream(value).decode()).toString('latin1')
        : ''
    }).join('')
  }
  return ''
}

type Matrix = [number, number, number, number, number, number]

function multiplyMatrices(left: Matrix, right: Matrix): Matrix {
  return [
    left[0] * right[0] + left[2] * right[1],
    left[1] * right[0] + left[3] * right[1],
    left[0] * right[2] + left[2] * right[3],
    left[1] * right[2] + left[3] * right[3],
    left[0] * right[4] + left[2] * right[5] + left[4],
    left[1] * right[4] + left[3] * right[5] + left[5],
  ]
}

function composedMatrix(content: string): Matrix {
  const matrices = [
    ...content.matchAll(
      /(-?(?:\d+\.?\d*|\.\d+))\s+(-?(?:\d+\.?\d*|\.\d+))\s+(-?(?:\d+\.?\d*|\.\d+))\s+(-?(?:\d+\.?\d*|\.\d+))\s+(-?(?:\d+\.?\d*|\.\d+))\s+(-?(?:\d+\.?\d*|\.\d+))\s+cm/g,
    ),
  ].map((match) => match.slice(1, 7).map(Number) as Matrix)
  return matrices.reduce(multiplyMatrices, [1, 0, 0, 1, 0, 0])
}

function transformedBounds(matrix: Matrix, rect: [number, number, number, number]) {
  const points = [
    [rect[0], rect[1]],
    [rect[0] + rect[2], rect[1]],
    [rect[0], rect[1] + rect[3]],
    [rect[0] + rect[2], rect[1] + rect[3]],
  ].map(([x, y]) => [
    matrix[0] * x + matrix[2] * y + matrix[4],
    matrix[1] * x + matrix[3] * y + matrix[5],
  ])
  return {
    minX: Math.min(...points.map(([x]) => x)),
    minY: Math.min(...points.map(([, y]) => y)),
    maxX: Math.max(...points.map(([x]) => x)),
    maxY: Math.max(...points.map(([, y]) => y)),
  }
}

describe('mergeGrid', () => {
  it('is a pair for 2 and a near-square grid otherwise', () => {
    expect(mergeGrid(2)).toEqual({ cols: 2, rows: 1 })
    expect(mergeGrid(4)).toEqual({ cols: 2, rows: 2 })
    expect(mergeGrid(6)).toEqual({ cols: 3, rows: 2 })
    expect(mergeGrid(9)).toEqual({ cols: 3, rows: 3 })
    expect(mergeGrid(16)).toEqual({ cols: 4, rows: 4 })
  })
})

describe('mergePagesBytes', () => {
  it('puts every perSheet pages onto one sheet sized like the first page', async () => {
    const bytes = await makePdf([
      [100, 200],
      [100, 200],
      [100, 200],
      [100, 200],
      [100, 200],
    ])
    const out = await PDFDocument.load(await mergePagesBytes(bytes, mergeOpts(4)))
    // 5 pages at 4 per sheet → 2 sheets, keeping the source page size
    expect(out.getPageCount()).toBe(2)
    expect(out.getPage(0).getWidth()).toBe(100)
    expect(out.getPage(0).getHeight()).toBe(200)
  })

  it('supports arbitrary counts like WPS, e.g. 6 per sheet', async () => {
    const bytes = await makePdf(Array.from({ length: 7 }, () => [100, 200] as [number, number]))
    const out = await PDFDocument.load(await mergePagesBytes(bytes, mergeOpts(6)))
    expect(out.getPageCount()).toBe(2)
  })

  it('swaps sheet width/height for 2-up so portrait pages sit side by side', async () => {
    const bytes = await makePdf([
      [100, 200],
      [100, 200],
      [100, 200],
    ])
    const out = await PDFDocument.load(await mergePagesBytes(bytes, mergeOpts(2)))
    expect(out.getPageCount()).toBe(2)
    expect(out.getPage(0).getWidth()).toBe(200)
    expect(out.getPage(0).getHeight()).toBe(100)
  })

  it('preserves source-page rotation when imposing pages', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([100, 200])
    page.setRotation(degrees(90))
    page.drawRectangle({ x: 10, y: 20, width: 30, height: 40, color: rgb(1, 0, 0) })
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await PDFDocument.load(await mergePagesBytes(bytes, mergeOpts(2)))
    expect(out.getPage(0).getWidth()).toBe(100)
    expect(out.getPage(0).getHeight()).toBe(200)
    const matrix = composedMatrix(decodedPageContent(out.getPage(0)))
    expect(matrix[0]).toBeCloseTo(0)
    expect(matrix[1]).toBeCloseTo(-0.25)
    expect(matrix[2]).toBeCloseTo(0.25)
    expect(matrix[3]).toBeCloseTo(0)
    expect(matrix[4]).toBeCloseTo(0)
    expect(matrix[5]).toBeCloseTo(112.5)
    const bounds = transformedBounds(matrix, [10, 20, 30, 40])
    expect(bounds.minX).toBeCloseTo(5)
    expect(bounds.minY).toBeCloseTo(102.5)
    expect(bounds.maxX).toBeCloseTo(15)
    expect(bounds.maxY).toBeCloseTo(110)
  })

  it('draws the embedded pages onto each sheet', async () => {
    const bytes = await makePdf([
      [100, 100],
      [100, 100],
      [100, 100],
      [100, 100],
    ])
    const out = await PDFDocument.load(await mergePagesBytes(bytes, mergeOpts(4)))
    expect(out.getPageCount()).toBe(1)
    const xobjects = out.getPage(0).node.Resources()?.lookupMaybe(PDFName.of('XObject'), PDFDict)
    expect(xobjects?.keys().length).toBe(4)
  })

  it('separator option adds line-drawing operators to the sheet content', async () => {
    const pages: [number, number][] = [
      [100, 100],
      [100, 100],
      [100, 100],
      [100, 100],
    ]
    const plain = await mergePagesBytes(await makePdf(pages), mergeOpts(4))
    const lined = await mergePagesBytes(await makePdf(pages), {
      ...mergeOpts(4),
      separator: true,
    })
    // The separator variant carries extra stroke content
    expect(lined.length).toBeGreaterThan(plain.length)
  })
})

describe('mergePdfBytes', () => {
  it('appends all pages of the other PDFs in order', async () => {
    const first = await makePdf([[100, 100]])
    const second = await makePdf([
      [200, 200],
      [300, 300],
    ])
    const third = await makePdf([[400, 400]])
    const { merged, appended } = await mergePdfBytes(first, [second, third])
    expect(appended).toBe(3)
    const out = await PDFDocument.load(merged)
    expect(out.getPageCount()).toBe(4)
    expect(out.getPage(0).getWidth()).toBe(100)
    expect(out.getPage(1).getWidth()).toBe(200)
    expect(out.getPage(3).getWidth()).toBe(400)
  })
})

describe('replacePagesBytes', () => {
  it('swaps the given pages for the other PDF at the first replaced position', async () => {
    const base = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const other = await makePdf([
      [500, 500],
      [600, 600],
    ])
    const { merged, removed, inserted } = await replacePagesBytes(base, other, [1])
    expect(removed).toBe(1)
    expect(inserted).toBe(2)
    const out = await PDFDocument.load(merged)
    expect(out.getPages().map((p) => p.getWidth())).toEqual([100, 500, 600, 300])
  })

  it('handles non-contiguous pages, inserting at the earliest one', async () => {
    const base = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const other = await makePdf([[500, 500]])
    const { merged } = await replacePagesBytes(base, other, [2, 0])
    const out = await PDFDocument.load(merged)
    expect(out.getPages().map((p) => p.getWidth())).toEqual([500, 200])
  })
})

describe('setPageSizeBytes', () => {
  it('resizes every page to the target size', async () => {
    const bytes = await makePdf([
      [100, 200],
      [400, 100],
    ])
    const out = await PDFDocument.load(await setPageSizeBytes(bytes, 595, 842))
    for (const page of out.getPages()) {
      expect(page.getWidth()).toBe(595)
      expect(page.getHeight()).toBe(842)
    }
  })

  it('swaps the target for sideways pages so the displayed size matches', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([100, 200]).setRotation(degrees(90))
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await PDFDocument.load(await setPageSizeBytes(bytes, 595, 842))
    expect(out.getPage(0).getWidth()).toBe(842)
    expect(out.getPage(0).getHeight()).toBe(595)
  })
})

describe('splitPagesBytes', () => {
  it('splits each page into side-by-side halves for 2', async () => {
    const bytes = await makePdf([[200, 100]])
    const out = await PDFDocument.load(await splitPagesBytes(bytes, 2))
    expect(out.getPageCount()).toBe(2)
    const [a, b] = out.getPages()
    expect(a!.getMediaBox()).toMatchObject({ x: 0, width: 100, height: 100 })
    expect(b!.getMediaBox()).toMatchObject({ x: 100, width: 100, height: 100 })
  })

  it('splits into a 2x2 grid for 4, top-left cell first', async () => {
    const bytes = await makePdf([
      [200, 100],
      [200, 100],
    ])
    const out = await PDFDocument.load(await splitPagesBytes(bytes, 4))
    expect(out.getPageCount()).toBe(8)
    // First cell is the top-left quarter: x=0, y=50 (PDF y goes up)
    expect(out.getPage(0).getMediaBox()).toMatchObject({ x: 0, y: 50, width: 100, height: 50 })
    expect(out.getPage(3).getMediaBox()).toMatchObject({ x: 100, y: 0, width: 100, height: 50 })
  })

  it('lays the grid on the displayed page through /Rotate 90', async () => {
    // Portrait 100x200 rotated 90 displays landscape 200x100; the first of two
    // side-by-side halves is the displayed left half = the user-space bottom half
    const doc = await PDFDocument.create()
    doc.addPage([100, 200]).setRotation(degrees(90))
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await PDFDocument.load(await splitPagesBytes(bytes, 2))
    expect(out.getPageCount()).toBe(2)
    expect(out.getPage(0).getMediaBox()).toMatchObject({ x: 0, y: 0, width: 100, height: 100 })
    expect(out.getPage(1).getMediaBox()).toMatchObject({ x: 0, y: 100, width: 100, height: 100 })
  })
})

describe('cropPagesBytes', () => {
  it('shrinks the CropBox to the fractional rect (y measured from the top)', async () => {
    const bytes = await makePdf([
      [100, 200],
      [100, 200],
    ])
    const out = await PDFDocument.load(
      await cropPagesBytes(bytes, [0], { l: 0.1, t: 0.2, r: 0.6, b: 0.7 }),
    )
    const box = out.getPage(0).getCropBox()
    expect(box.x).toBeCloseTo(10)
    expect(box.y).toBeCloseTo(60)
    expect(box.width).toBeCloseTo(50)
    expect(box.height).toBeCloseTo(100)
    // Page 1 untouched
    expect(out.getPage(1).getCropBox()).toMatchObject({ x: 0, y: 0, width: 100, height: 200 })
  })

  it('maps the displayed rect through /Rotate 90', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([100, 200]).setRotation(degrees(90))
    const bytes = await doc.save({ useObjectStreams: false })
    const out = await PDFDocument.load(
      await cropPagesBytes(bytes, [0], { l: 0.1, t: 0.2, r: 0.6, b: 0.7 }),
    )
    expect(out.getPage(0).getCropBox()).toMatchObject({ x: 20, y: 20, width: 50, height: 100 })
  })
})

describe('static form fill metadata', () => {
  it('persists records and remaps their pages through delete and reorder', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await apply(
      bytes,
      request({
        staticFormFills: [
          { id: 'a', kind: 'text', pageIndex: 0, rect: [1, 2, 30, 12], text: 'Alice' },
          { id: 'b', kind: 'check', pageIndex: 2, rect: [5, 6, 15, 16] },
          { id: 'deleted', kind: 'cross', pageIndex: 1, rect: [5, 6, 15, 16] },
        ],
        deletedPages: [1],
        pageOrder: [2, 0],
      }),
    )

    expect(await readStaticFormFills(saved)).toEqual([
      { id: 'a', kind: 'text', pageIndex: 1, rect: [1, 2, 30, 12], text: 'Alice' },
      { id: 'b', kind: 'check', pageIndex: 0, rect: [5, 6, 15, 16] },
    ])
  })

  it('removes the catalog entry when the last fill is deleted', async () => {
    const withFill = await apply(
      await makePdf([[100, 100]]),
      request({
        staticFormFills: [{ id: 'a', kind: 'cross', pageIndex: 0, rect: [1, 2, 3, 4] }],
      }),
    )
    const cleared = await apply(withFill, request({ staticFormFills: [] }))

    expect(await readStaticFormFills(cleared)).toEqual([])
  })
})

describe('savePdfToPath', () => {
  const sha256 = (path: string) => createHash('sha256').update(readFileSync(path)).digest('hex')
  const highlight = {
    pageIndex: 0,
    type: 'highlight' as const,
    color: [1, 0.87, 0.35] as [number, number, number],
    quads: [[10, 100, 60, 100, 10, 88, 60, 88]],
  }

  it('Save As writes the edits to the target only and never mutates the source', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-'))
    const src = join(dir, 'original.pdf')
    const dst = join(dir, 'copy.pdf')
    writeFileSync(src, await makePdf([[612, 792]]))
    const srcHash = sha256(src)
    const srcInode = statSync(src).ino

    await savePdfToPath(src, dst, request({ path: src, targetPath: dst, markups: [highlight] }))

    // Source: same inode, same bytes
    expect(sha256(src)).toBe(srcHash)
    expect(statSync(src).ino).toBe(srcInode)
    // Target: valid PDF containing the new annotation
    const out = await PDFDocument.load(new Uint8Array(readFileSync(dst)))
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Highlight'])
    // No temp files left behind
    expect(readdirSync(dir).sort()).toEqual(['copy.pdf', 'original.pdf'])
  })

  it('in-place save (target === source) replaces the file atomically', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-'))
    const src = join(dir, 'doc.pdf')
    writeFileSync(src, await makePdf([[612, 792]]))

    await savePdfToPath(src, src, request({ path: src, markups: [highlight] }))

    const out = await PDFDocument.load(new Uint8Array(readFileSync(src)))
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Highlight'])
    expect(readdirSync(dir)).toEqual(['doc.pdf'])
  })

  it('skips markups with non-finite colors or quads instead of corrupting the file', async () => {
    const nanColor = { ...highlight, color: [NaN, 0, 0] as [number, number, number] }
    const outOfRange = { ...highlight, color: [2, 0, 0] as [number, number, number] }
    const nanQuads = { ...highlight, quads: [[10, NaN, 60, 100, 10, 88, 60, 88]] }
    const emptyQuads = { ...highlight, quads: [] as number[][] }
    const out = await PDFDocument.load(
      await apply(
        await makePdf([[612, 792]]),
        request({ markups: [highlight, nanColor, outOfRange, nanQuads, emptyQuads] }),
      ),
    )
    // only the valid highlight survives; the output still re-opens cleanly
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Highlight'])
  })

  it('a failed save leaves the source and target untouched and cleans up temp files', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-'))
    const src = join(dir, 'original.pdf')
    writeFileSync(src, await makePdf([[612, 792]]))
    const srcHash = sha256(src)

    // Apply failure (unknown form field): nothing may be written anywhere
    await expect(
      savePdfToPath(
        src,
        join(dir, 'copy.pdf'),
        request({ path: src, formValues: [{ name: 'missing', kind: 'text', value: 'x' }] }),
      ),
    ).rejects.toThrow()
    expect(sha256(src)).toBe(srcHash)
    expect(readdirSync(dir)).toEqual(['original.pdf'])

    // Write failure (target directory does not exist): source intact, temp cleaned up
    await expect(
      savePdfToPath(src, join(dir, 'no-such-dir', 'copy.pdf'), request({ path: src })),
    ).rejects.toThrow()
    expect(sha256(src)).toBe(srcHash)
    expect(readdirSync(dir)).toEqual(['original.pdf'])
  })
})

describe('applySaveRequest', () => {
  it('applies page rotation deltas on top of the existing rotation', async () => {
    const bytes = await makePdf([[100, 100]])
    const saved = await apply(bytes, request({ rotations: [{ pageIndex: 0, delta: 90 }] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPage(0).getRotation().angle).toBe(90)
  })

  it('writes markup annotations with an appearance stream', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        markups: [
          {
            pageIndex: 0,
            type: 'highlight',
            color: [1, 0.87, 0.35],
            quads: [[10, 100, 60, 100, 10, 88, 60, 88]],
          },
          {
            pageIndex: 0,
            type: 'underline',
            color: [0.17, 0.4, 1],
            quads: [[10, 80, 60, 80, 10, 68, 60, 68]],
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Highlight', 'Underline'])
    // Every markup gets a hand-written /AP /N appearance
    for (const a of annots) {
      expect(a.lookup(PDFName.of('AP'), PDFDict).has(PDFName.of('N'))).toBe(true)
    }
  })

  it('writes note and shape drawing annotations', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        drawings: [
          { kind: 'note', pageIndex: 0, color: [1, 0, 0], at: [50, 700], contents: 'hello note' },
          {
            kind: 'ink',
            pageIndex: 0,
            color: [0, 0, 1],
            width: 2,
            paths: [[10, 10, 20, 20, 30, 15]],
          },
          { kind: 'rect', pageIndex: 0, color: [0, 1, 0], width: 1, rect: [40, 40, 90, 80] },
          {
            kind: 'arrow',
            pageIndex: 0,
            color: [0, 0, 0],
            width: 2,
            from: [100, 100],
            to: [200, 150],
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Text', 'Ink', 'Square', 'Line'])
  })

  // A sticky note is the one annotation whose own entries are not enough for
  // macOS Preview: without /AP it has no icon to draw, so the note reads as absent.
  describe('sticky note interop (Preview / pdf.js)', () => {
    const note = (over: Record<string, unknown> = {}) => ({
      kind: 'note' as const,
      pageIndex: 0,
      color: [1, 0.9, 0.3] as [number, number, number],
      at: [120, 700] as [number, number],
      contents: 'hello note',
      ...over,
    })

    async function saveNoteWith(drawing: ReturnType<typeof note>) {
      const bytes = await makePdf([[612, 792]])
      const saved = await apply(bytes, request({ drawings: [drawing] }))
      const doc = await PDFDocument.load(saved)
      const annots = doc.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
      return { doc, annots, noteRef: annots.lookup(0) as PDFRef }
    }

    it('writes an appearance stream for the note icon', async () => {
      const { doc, noteRef } = await saveNoteWith(note())
      const dict = doc.context.lookup(noteRef) as PDFDict
      expect(dict.lookup(PDFName.of('AP'), PDFDict).has(PDFName.of('N'))).toBe(true)
    })

    it('gives the appearance form its own /Resources', async () => {
      // A Form XObject with no /Resources key at all draws dark grey in Quartz
      // and yellow in Poppler — the note appears in one viewer only. This is the
      // difference that silently shipped, so it is asserted directly.
      const { doc, noteRef } = await saveNoteWith(note())
      const dict = doc.context.lookup(noteRef) as PDFDict
      const apRef = dict.lookup(PDFName.of('AP'), PDFDict).lookup(PDFName.of('N')) as PDFRef
      const ap = doc.context.lookup(apRef) as PDFRawStream
      expect(ap.dict.lookup(PDFName.of('Subtype'), PDFName).decodeText()).toBe('Form')
      expect(ap.dict.has(PDFName.of('Resources'))).toBe(true)
    })

    it('still reads back as exactly one note in pdf.js', async () => {
      const bytes = await makePdf([[612, 792]])
      const saved = await apply(
        bytes,
        request({ drawings: [note(), note({ at: [300, 400], contents: 'second' })] }),
      )
      const loadingTask = getDocument({ data: saved.slice() })
      try {
        const pdfJsDoc = await loadingTask.promise
        const annos = await (await pdfJsDoc.getPage(1)).getAnnotations()
        const texts = annos.filter((a) => a.annotationType === PDFJS_ANNOT_TEXT)
        expect(texts.map((a) => a.contentsObj?.str)).toEqual(['hello note', 'second'])
      } finally {
        await loadingTask.destroy()
      }
    })
  })

  it('ignores markups and drawings addressing missing pages', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        markups: [
          { pageIndex: 9, type: 'highlight', color: [1, 1, 0], quads: [[0, 1, 1, 1, 0, 0, 1, 0]] },
        ],
        drawings: [{ kind: 'note', pageIndex: 9, color: [1, 0, 0], at: [0, 0], contents: 'x' }],
      }),
    )
    expect(pageAnnots(await PDFDocument.load(saved), 0)).toHaveLength(0)
  })

  it('writes an image drawing as a Stamp annotation with an image appearance', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        drawings: [{ kind: 'image', pageIndex: 0, image: TINY_PNG, rect: [100, 500, 300, 600] }],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Stamp'])
    const rect = annots[0]!.lookup(PDFName.of('Rect'), PDFArray)
    expect(String(rect)).toBe('[ 100 500 300 600 ]')
    expect(annots[0]!.lookup(PDFName.of('AP'), PDFDict).has(PDFName.of('N'))).toBe(true)
  })

  it('persists the AcroForm field association on visual signature annotations', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        drawings: [
          {
            kind: 'image',
            pageIndex: 0,
            image: TINY_PNG,
            rect: [100, 500, 300, 600],
            formFieldName: 'signature.image',
          },
          {
            kind: 'ink',
            pageIndex: 0,
            color: [0, 0, 0],
            width: 2,
            paths: [[10, 10, 20, 20]],
            formFieldName: 'signature.ink',
          },
        ],
      }),
    )
    const annots = pageAnnots(await PDFDocument.load(saved), 0)

    for (const [index, fieldName] of ['signature.image', 'signature.ink'].entries()) {
      expect(
        annots[index]!.lookup(PDFName.of('GenOfficeFormField'), PDFHexString).decodeText(),
      ).toBe(fieldName)
      expect(annots[index]!.lookup(PDFName.of('Contents'), PDFHexString).decodeText()).toBe(
        `${VISUAL_SIGNATURE_CONTENT_PREFIX}${fieldName}`,
      )
    }

    const loadingTask = getDocument({ data: saved.slice() })
    try {
      const pdfJsDoc = await loadingTask.promise
      const pdfJsAnnotations = await (await pdfJsDoc.getPage(1)).getAnnotations()
      expect(pdfJsAnnotations.map((annotation) => annotation.contentsObj?.str)).toEqual([
        `${VISUAL_SIGNATURE_CONTENT_PREFIX}signature.image`,
        `${VISUAL_SIGNATURE_CONTENT_PREFIX}signature.ink`,
      ])
    } finally {
      await loadingTask.destroy()
    }
  })

  it('counter-rotates the image appearance on rotated pages', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        rotations: [{ pageIndex: 0, delta: 90 }],
        drawings: [{ kind: 'image', pageIndex: 0, image: TINY_PNG, rect: [100, 500, 300, 600] }],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Stamp'])
  })

  it('embeds PNG stamps without failing', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        stamps: [{ pageIndex: 0, image: TINY_PNG, rect: [0, 0, 612, 792], opacity: 0.2 }],
      }),
    )
    expect((await PDFDocument.load(saved)).getPageCount()).toBe(1)
  })

  it('applies metadata and splits keywords on mixed separators', async () => {
    const bytes = await makePdf([[100, 100]])
    const saved = await apply(
      bytes,
      request({ metadata: { title: 'My Title', author: 'Me', keywords: 'a, b；c，d' } }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getTitle()).toBe('My Title')
    expect(out.getAuthor()).toBe('Me')
    expect(out.getKeywords()).toContain('a')
    expect(out.getKeywords()).toContain('d')
  })

  it('deletes pages by original index but never removes the last page', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await apply(bytes, request({ deletedPages: [0, 2] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(1)
    expect(out.getPage(0).getWidth()).toBe(200)

    const savedAll = await apply(bytes, request({ deletedPages: [0, 1, 2] }))
    expect((await PDFDocument.load(savedAll)).getPageCount()).toBe(1)
  })

  it('reorders pages by original index', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await apply(bytes, request({ pageOrder: [2, 0, 1] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPages().map((p) => p.getWidth())).toEqual([300, 100, 200])
  })

  it('applies the reorder when combined with deletions', async () => {
    // Regression: pdf-lib's removePage does not invalidate its page cache,
    // so a getPages() call after the deletion loop returns the stale
    // pre-deletion list. applySaveRequest must derive the remaining pages
    // from the pre-deletion snapshot instead of re-reading them.
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await apply(bytes, request({ deletedPages: [1], pageOrder: [2, 0] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(2)
    expect(out.getPages().map((p) => p.getWidth())).toEqual([300, 100])
  })

  it('skips the reorder when deleting every page leaves the guarded last page', async () => {
    // Deleting all pages keeps one via the last-page guard; the order list
    // then matches nothing alive, so the reorder must be skipped safely.
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
    ])
    const saved = await apply(bytes, request({ deletedPages: [0, 1], pageOrder: [1, 0] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(1)
  })

  it('fills text fields and checkboxes', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    const form = doc.getForm()
    form.createTextField('user.name').addToPage(page, { x: 20, y: 200, width: 200, height: 20 })
    form.createCheckBox('user.agree').addToPage(page, { x: 20, y: 150, width: 16, height: 16 })
    const bytes = await doc.save({ useObjectStreams: false })

    const saved = await apply(
      bytes,
      request({
        formValues: [
          { name: 'user.name', kind: 'text', value: 'Alice' },
          { name: 'user.agree', kind: 'checkbox', checked: true },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getForm().getTextField('user.name').getText()).toBe('Alice')
    expect(out.getForm().getCheckBox('user.agree').isChecked()).toBe(true)
  })

  it('fills radio groups and dropdowns', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    const form = doc.getForm()
    const radio = form.createRadioGroup('user.color')
    radio.addOptionToPage('red', page, { x: 20, y: 200, width: 16, height: 16 })
    radio.addOptionToPage('blue', page, { x: 50, y: 200, width: 16, height: 16 })
    const country = form.createDropdown('user.country')
    country.addOptions(['CN', 'US'])
    country.addToPage(page, { x: 20, y: 150, width: 100, height: 20 })
    const bytes = await doc.save({ useObjectStreams: false })

    const saved = await apply(
      bytes,
      request({
        formValues: [
          { name: 'user.color', kind: 'radio', value: 'blue' },
          { name: 'user.country', kind: 'choice', value: 'CN' },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)

    expect(out.getForm().getRadioGroup('user.color').getSelected()).toBe('blue')
    expect(out.getForm().getDropdown('user.country').getSelected()).toEqual(['CN'])
  })

  it('appends blank pages first so ordering, deletion and edits on them apply', async () => {
    const bytes = await makePdf([
      [300, 300],
      [300, 300],
    ])
    const saved = await apply(
      bytes,
      request({
        blankPages: [
          { pageIndex: 2, width: 400, height: 200 },
          { pageIndex: 3, width: 100, height: 100 },
        ],
        pageOrder: [2, 0, 1, 3],
        deletedPages: [3],
        formFields: [{ name: 'on_blank', kind: 'text', pageIndex: 2, rect: [10, 10, 110, 30] }],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(3)
    const first = out.getPage(0)
    expect([first.getWidth(), first.getHeight()]).toEqual([400, 200])
    const widget = out.getForm().getTextField('on_blank').acroField.getWidgets()[0]!
    expect(widget.P()).toBe(first.ref)
  })

  it('rejects a blank page whose index does not follow the document', async () => {
    const bytes = await makePdf([[300, 300]])
    await expect(
      apply(bytes, request({ blankPages: [{ pageIndex: 5, width: 100, height: 100 }] })),
    ).rejects.toThrow(/does not follow/)
  })

  it('creates authored text, checkbox and signature fields as real widgets', async () => {
    const bytes = await makePdf([[300, 300]])
    const saved = await apply(
      bytes,
      request({
        formFields: [
          {
            name: 'full_name',
            kind: 'text',
            pageIndex: 0,
            rect: [20, 200, 220, 222],
            required: true,
            maxLen: 40,
            value: 'Bob',
          },
          {
            name: 'agree',
            kind: 'checkbox',
            pageIndex: 0,
            rect: [20, 150, 36, 166],
            checked: true,
          },
          { name: 'sig', kind: 'signature', pageIndex: 0, rect: [20, 50, 170, 90], required: true },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const form = out.getForm()
    const name = form.getTextField('full_name')
    expect(name.getText()).toBe('Bob')
    expect(name.isRequired()).toBe(true)
    expect(name.getMaxLength()).toBe(40)
    expect(form.getCheckBox('agree').isChecked()).toBe(true)
    expect(form.getField('sig').acroField.dict.get(PDFName.of('FT'))).toBe(PDFName.of('Sig'))

    const loadingTask = getDocument({ data: saved.slice() })
    try {
      const annos = await (await (await loadingTask.promise).getPage(1)).getAnnotations()
      const widgets = annos.filter((a) => a.subtype === 'Widget')
      expect(widgets.map((a) => [a.fieldName, a.fieldType])).toEqual([
        ['full_name', 'Tx'],
        ['agree', 'Btn'],
        ['sig', 'Sig'],
      ])
      expect(widgets[0]!.rect.map(Math.round)).toEqual([20, 200, 220, 222])
      expect(widgets[2]!.required).toBe(true)
    } finally {
      await loadingTask.destroy()
    }
  })

  it('creates radio groups, dropdowns and date fields', async () => {
    const bytes = await makePdf([[300, 300]])
    const saved = await apply(
      bytes,
      request({
        formFields: [
          {
            name: 'color',
            kind: 'radio',
            pageIndex: 0,
            rect: [20, 250, 34, 264],
            exportValue: 'red',
          },
          {
            name: 'color',
            kind: 'radio',
            pageIndex: 0,
            rect: [60, 250, 74, 264],
            exportValue: 'blue',
            checked: true,
          },
          {
            name: 'country',
            kind: 'choice',
            pageIndex: 0,
            rect: [20, 200, 180, 222],
            options: ['CN', 'US', ''],
            value: 'US',
          },
          {
            name: 'when',
            kind: 'text',
            pageIndex: 0,
            rect: [20, 150, 130, 172],
            dateFormat: 'dd/mm/yyyy',
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const form = out.getForm()
    const radio = form.getRadioGroup('color')
    expect(radio.getOptions()).toEqual(['red', 'blue'])
    expect(radio.getSelected()).toBe('blue')
    const dd = form.getDropdown('country')
    expect(dd.getOptions()).toEqual(['CN', 'US'])
    expect(dd.getSelected()).toEqual(['US'])
    const when = form.getTextField('when')
    const widget = when.acroField.getWidgets()[0]!
    expect(when.acroField.dict.has(PDFName.of('AA'))).toBe(true)
    const aa = widget.dict.lookup(PDFName.of('AA'), PDFDict)
    const fmt = aa.lookup(PDFName.of('F'), PDFDict).lookup(PDFName.of('JS'), PDFString)
    expect(fmt.decodeText()).toBe('AFDate_FormatEx("dd/mm/yyyy");')

    // A non-radio field cannot reuse a radio group's name
    await expect(
      apply(
        saved,
        request({
          formFields: [{ name: 'color', kind: 'text', pageIndex: 0, rect: [0, 0, 50, 20] }],
        }),
      ),
    ).rejects.toThrow(/already exists/)
  })

  it('moves, re-flags and removes widgets already in the file', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 400])
    const form = doc.getForm()
    form.createTextField('name').addToPage(page, { x: 10, y: 300, width: 100, height: 20 })
    form.createCheckBox('agree').addToPage(page, { x: 10, y: 200, width: 14, height: 14 })
    const multi = form.createTextField('twice')
    multi.addToPage(page, { x: 10, y: 100, width: 50, height: 20 })
    multi.addToPage(page, { x: 100, y: 100, width: 50, height: 20 })
    const bytes = await doc.save({ useObjectStreams: false })

    const ids = async (data: Uint8Array) => {
      const task = getDocument({ data: data.slice() })
      try {
        const annos = await (await (await task.promise).getPage(1)).getAnnotations()
        return annos
          .filter((a) => a.subtype === 'Widget')
          .map((a) => ({
            id: a.id as string,
            name: a.fieldName as string,
            rect: a.rect as number[],
          }))
      } finally {
        await task.destroy()
      }
    }
    const before = await ids(bytes)
    const nameW = before.find((w) => w.name === 'name')!
    const agreeW = before.find((w) => w.name === 'agree')!
    const twiceW = before.filter((w) => w.name === 'twice')

    const saved = await apply(
      bytes,
      request({
        formWidgetEdits: [
          { widgetId: nameW.id, fieldName: 'name', rect: [20, 250, 220, 274], required: true },
          { widgetId: agreeW.id, fieldName: 'agree', remove: true },
          { widgetId: twiceW[0]!.id, fieldName: 'twice', remove: true },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const outForm = out.getForm()
    expect(outForm.getFieldMaybe('agree')).toBeUndefined()
    expect(outForm.getTextField('name').isRequired()).toBe(true)
    expect(outForm.getTextField('twice').acroField.getWidgets()).toHaveLength(1)
    const after = await ids(saved)
    expect(after.find((w) => w.name === 'name')!.rect.map(Math.round)).toEqual([20, 250, 220, 274])
    expect(after.filter((w) => w.name === 'twice')).toHaveLength(1)
    expect(after.some((w) => w.name === 'agree')).toBe(false)
  })

  it('keeps the authored rect on a rotated page and rejects duplicate names', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 400])
    page.setRotation(degrees(90))
    doc.getForm().createTextField('taken').addToPage(page, { x: 10, y: 10, width: 50, height: 20 })
    const bytes = await doc.save({ useObjectStreams: false })

    const saved = await apply(
      bytes,
      request({
        formFields: [{ name: 'rotated', kind: 'text', pageIndex: 0, rect: [30, 100, 130, 124] }],
      }),
    )
    const out = await PDFDocument.load(saved)
    const widget = out.getForm().getTextField('rotated').acroField.getWidgets()[0]!
    const r = widget.getRectangle()
    expect([r.x, r.y, r.x + r.width, r.y + r.height].map(Math.round)).toEqual([30, 100, 130, 124])
    expect(widget.getAppearanceCharacteristics()?.getRotation()).toBe(90)

    await expect(
      apply(
        bytes,
        request({
          formFields: [{ name: 'taken', kind: 'checkbox', pageIndex: 0, rect: [0, 0, 10, 10] }],
        }),
      ),
    ).rejects.toThrow(/already exists/)
  })

  it('falls back to NeedAppearances when form values cannot be WinAnsi-encoded', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    doc.getForm().createTextField('cjk').addToPage(page, { x: 20, y: 200, width: 200, height: 20 })
    const bytes = await doc.save({ useObjectStreams: false })

    const saved = await apply(
      bytes,
      request({ formValues: [{ name: 'cjk', kind: 'text', value: '中文测试' }] }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getForm().getTextField('cjk').getText()).toBe('中文测试')
    const needAppearances = out.getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))
    expect(String(needAppearances)).toBe('true')
  })
})
