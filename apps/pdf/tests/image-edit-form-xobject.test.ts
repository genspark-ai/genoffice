import { PDFDocument } from 'pdf-lib'
import { describe, expect, it, vi } from 'vitest'
import { SHARED_TEMPLATE_IMAGE_REASON } from '../src/main/form-text'
import { applyImageEdits, listPageImages } from '../src/main/image-edit'

vi.mock('electron', () => ({ nativeImage: {} }))

const FORM_X = 60
const FORM_Y = 300
const SCALE = 0.5
const IMG = { x: 100, y: 500, w: 200, h: 120 }

// 2x2 opaque PNG (base64; pdf-lib accepts strings, which sidesteps jsdom's realm check)
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGP4z8DwHwyBNBAAAFRNB/8GxGIYAAAAAElFTkSuQmCC'

/** A page whose only content is a Form XObject (an embedded page) holding one image */
async function makeTemplateDoc(opts: { nested?: boolean; copies?: number; pages?: number } = {}) {
  const src = await PDFDocument.create()
  const srcPage = src.addPage([595, 842])
  const png = await src.embedPng(PNG)
  srcPage.drawImage(png, { x: IMG.x, y: IMG.y, width: IMG.w, height: IMG.h })
  let inner = await PDFDocument.load(await src.save({ useObjectStreams: false }))
  if (opts.nested) {
    const mid = await PDFDocument.create()
    const midPage = mid.addPage([595, 842])
    midPage.drawPage(await mid.embedPage(inner.getPage(0)), { x: 0, y: 0 })
    inner = await PDFDocument.load(await mid.save({ useObjectStreams: false }))
  }
  const doc = await PDFDocument.create()
  const emb = await doc.embedPage(inner.getPage(0))
  for (let p = 0; p < (opts.pages ?? 1); p++) {
    const page = doc.addPage([595, 842])
    for (let c = 0; c < (opts.copies ?? 1); c++) {
      page.drawPage(emb, { x: FORM_X, y: FORM_Y - c * 250, xScale: SCALE, yScale: SCALE })
    }
  }
  return doc.save({ useObjectStreams: false })
}

const expectedRect = [
  FORM_X + IMG.x * SCALE,
  FORM_Y + IMG.y * SCALE,
  FORM_X + (IMG.x + IMG.w) * SCALE,
  FORM_Y + (IMG.y + IMG.h) * SCALE,
] as const

/** Count image XObjects actually painted on a page, read back by pdf.js */
async function paintedImages(bytes: Uint8Array, pageIndex = 0): Promise<number> {
  const { getDocument, OPS } = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await getDocument({ data: bytes.slice() }).promise
  try {
    const page = await doc.getPage(pageIndex + 1)
    const ops = await page.getOperatorList()
    return ops.fnArray.filter((f) => f === OPS.paintImageXObject).length
  } finally {
    await doc.loadingTask.destroy()
  }
}

const expectRect = (got: readonly number[]) => {
  for (let i = 0; i < 4; i++) expect(got[i]).toBeCloseTo(expectedRect[i]!, 1)
}

describe('images inside Form XObjects', () => {
  it('lists a template image with its page-space rect', async () => {
    const bytes = await makeTemplateDoc()
    const images = await listPageImages(bytes)
    expect(images).toHaveLength(1)
    expect(images[0]!.pageIndex).toBe(0)
    expectRect(images[0]!.rect)
  })

  it('lists an image inside a nested template', async () => {
    const images = await listPageImages(await makeTemplateDoc({ nested: true }))
    expect(images).toHaveLength(1)
    expectRect(images[0]!.rect)
  })

  it('deletes a template image', async () => {
    const bytes = await makeTemplateDoc()
    expect(await paintedImages(bytes)).toBe(1)
    const { bytes: out, skipped } = await applyImageEdits(bytes, [
      { kind: 'deleteImage', pageIndex: 0, oldRect: [...expectedRect] },
    ])
    expect(skipped).toEqual([])
    expect(await paintedImages(out)).toBe(0)
    expect(await listPageImages(out)).toEqual([])
  })

  it('moves a template image in page space', async () => {
    const bytes = await makeTemplateDoc()
    const rect = [...expectedRect] as [number, number, number, number]
    const moved: [number, number, number, number] = [
      rect[0] + 40,
      rect[1] - 30,
      rect[2] + 40,
      rect[3] - 30,
    ]
    const { bytes: out, skipped } = await applyImageEdits(bytes, [
      { kind: 'transformImage', pageIndex: 0, oldRect: rect, rect: moved },
    ])
    expect(skipped).toEqual([])
    const images = await listPageImages(out)
    expect(images).toHaveLength(1)
    for (let i = 0; i < 4; i++) expect(images[0]!.rect[i]).toBeCloseTo(moved[i]!, 1)
  })

  it('refuses to edit an image in a template painted more than once', async () => {
    for (const opts of [{ pages: 2 }, { copies: 2 }]) {
      const bytes = await makeTemplateDoc(opts)
      const res = await applyImageEdits(bytes, [
        { kind: 'deleteImage', pageIndex: 0, oldRect: [...expectedRect] },
      ])
      expect(res.skipped.map((s) => s.reason)).toEqual([SHARED_TEMPLATE_IMAGE_REASON])
      expect(res.bytes).toBe(bytes)
    }
  })
})
