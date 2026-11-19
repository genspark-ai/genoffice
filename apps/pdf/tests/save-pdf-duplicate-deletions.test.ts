// @vitest-environment node
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFRawStream,
  PDFString,
  rgb,
} from 'pdf-lib'
import type { PDFPage } from 'pdf-lib'
import { expect, it } from 'vitest'
import { applySaveRequest, readStaticFormFills, savePdfToPath } from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

const request = (deletedPages: number[], extra: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: '/synthetic/input.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  deletedPages,
  ...extra,
})
async function fixture() {
  const doc = await PDFDocument.create()
  for (const [index, width] of [100, 200, 300].entries()) {
    const page = doc.addPage([width, 400])
    page.node.set(PDFName.of('FixturePage'), PDFString.of(`original-${index}`))
    page.drawRectangle({
      x: 10 + index,
      y: 20,
      width: 25 + index,
      height: 30,
      color: rgb(index / 3, 0, 0),
    })
  }
  return doc.save({ useObjectStreams: false })
}
function contents(page: PDFPage): Uint8Array[] {
  const value = page.node.lookup(PDFName.of('Contents'))
  const refs = value instanceof PDFArray ? value.asArray() : [value]
  return refs.map((ref) => page.doc.context.lookup(ref, PDFRawStream).contents)
}

it.each([
  [
    [0, 0],
    [1, 2],
  ],
  [
    [1, 1],
    [0, 2],
  ],
  [
    [0, 0, 0],
    [1, 2],
  ],
  [[2, 0, 2, 0], [1]],
  [
    [2, 2],
    [0, 1],
  ],
  [[0, 1, 2, 0, 1, 2], [0]],
  [[], [0, 1, 2]],
])('applies duplicate original-page deletions %j once', async (deleted, survivors) => {
  const input = await fixture(),
    originalBytes = input.slice()
  const before = await PDFDocument.load(input)
  const req = request(deleted),
    originalDeleted = [...deleted]
  const result = await applySaveRequest(input, req)
  const output = await PDFDocument.load(result.bytes)
  expect(output.getPages().map((page) => page.getWidth())).toEqual(
    survivors.map((index) => [100, 200, 300][index]),
  )
  for (const [index, originalIndex] of survivors.entries()) {
    const page = output.getPage(index)
    expect(page.node.lookup(PDFName.of('FixturePage'), PDFString).decodeText()).toBe(
      `original-${originalIndex}`,
    )
    expect(contents(page)).toEqual(contents(before.getPage(originalIndex)))
  }
  expect(input).toEqual(originalBytes)
  expect(req.deletedPages).toEqual(originalDeleted)
  expect(result.skippedTextEdits).toEqual([])
  expect(result.skippedTextInserts).toEqual([])
  expect(result.skippedImageEdits).toEqual([])
})

it('keeps static form metadata and bookmark targets aligned with the surviving original pages', async () => {
  const result = await applySaveRequest(
    await fixture(),
    request([0, 0], {
      staticFormFills: [
        { id: 'deleted', kind: 'text', pageIndex: 0, rect: [1, 2, 3, 4], text: 'gone' },
        { id: 'kept', kind: 'text', pageIndex: 2, rect: [1, 2, 3, 4], text: 'kept' },
      ],
      outline: [
        { title: 'Deleted', pageIndex: 0 },
        { title: 'Kept', pageIndex: 2 },
      ],
    }),
  )
  const output = await PDFDocument.load(result.bytes)
  expect(output.getPages().map((page) => page.getWidth())).toEqual([200, 300])
  expect(await readStaticFormFills(result.bytes)).toEqual([
    { id: 'kept', kind: 'text', pageIndex: 1, rect: [1, 2, 3, 4], text: 'kept' },
  ])
  const first = output.catalog
    .lookup(PDFName.of('Outlines'), PDFDict)
    .lookup(PDFName.of('First'), PDFDict)
  expect(first.lookup(PDFName.of('Title'), PDFHexString).decodeText()).toBe('Kept')
  expect(first.lookup(PDFName.of('Dest'), PDFArray).get(0)).toEqual(output.getPage(1).ref)
})

it('combines duplicate deletions with an existing valid original-page reorder', async () => {
  const result = await applySaveRequest(await fixture(), request([0, 0], { pageOrder: [2, 1] }))
  const output = await PDFDocument.load(result.bytes)
  expect(output.getPages().map((page) => page.getWidth())).toEqual([300, 200])
})

it('maps content read-back verification through unique deletions before publishing a saved file', async () => {
  const home = mkdtempSync(join(tmpdir(), 'pdf-duplicate-delete-'))
  try {
    const input = await fixture(),
      source = join(home, 'input.pdf'),
      target = join(home, 'output.pdf')
    writeFileSync(source, input)
    await savePdfToPath(
      source,
      target,
      request([0, 0], {
        path: source,
        targetPath: target,
        textInserts: [
          {
            pageIndex: 2,
            origin: [10, 100],
            text: 'Fixture insert',
            fontSize: 12,
            color: [0, 0, 0],
          },
        ],
      }),
    )
    expect(
      (await PDFDocument.load(readFileSync(target))).getPages().map((page) => page.getWidth()),
    ).toEqual([200, 300])
    expect(readFileSync(source)).toEqual(Buffer.from(input))
  } finally {
    rmSync(home, { recursive: true, force: true })
  }
})
