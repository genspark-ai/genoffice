// @vitest-environment node
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { expect, it, vi } from 'vitest'
import { chainPdfium, applyTextEdits } from '../src/main/text-edit'
import { subsetTtf } from '../src/main/font-subset'

// Only the repository font fixture is used. Never fall back to an installed face.
vi.mock('../src/main/font-locate', () => ({
  findSystemFont: () => {
    throw new Error('System fonts are outside this test')
  },
  findFontCovering: () => {
    throw new Error('System fonts are outside this test')
  },
  isTruetype: (bytes: Buffer) => bytes.readUInt32BE(0) === 0x00010000,
}))
const fixtureFont = () =>
  readFileSync(resolve(__dirname, '../../../packages/ui/src/fonts/Carlito-Regular.ttf'))

it('serializes overlapping PDF jobs and returns their original result objects', async () => {
  const events: string[] = []
  let release!: () => void
  const gate = new Promise<void>((resolve) => {
    release = resolve
  })
  const firstValue = new Uint8Array([1, 2, 3]),
    secondValue = { bytes: new Uint8Array([4, 5]) }
  const first = chainPdfium(async () => {
    events.push('first-start')
    await gate
    events.push('first-end')
    return firstValue
  })
  const second = chainPdfium(async () => {
    events.push('second')
    return secondValue
  })
  await Promise.resolve()
  await Promise.resolve()
  expect(events).toEqual(['first-start'])
  release()
  expect(await first).toBe(firstValue)
  expect(await second).toBe(secondValue)
  expect(events).toEqual(['first-start', 'first-end', 'second'])
})

it('preserves PDF job rejections and lets already queued jobs recover', async () => {
  const failure = new Error('synthetic failure')
  const failed = chainPdfium(async () => {
    throw failure
  })
  const rejection = expect(failed).rejects.toBe(failure)
  const success = chainPdfium(async () => 'after failure')
  await rejection
  expect(await success).toBe('after failure')
})

it('returns a successful native PDF edit and allows a following queued job', async () => {
  const doc = await PDFDocument.create()
  const page = doc.addPage([300, 400]),
    font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('Alpha', { x: 50, y: 300, size: 14, font })
  const bytes = await doc.save({ useObjectStreams: false })
  const result = await applyTextEdits(bytes, [
    { pageIndex: 0, rect: [45, 294, 100, 320], oldText: 'Alpha', newText: 'Bravo', fontSize: 14 },
  ])
  expect(result.skipped).toEqual([])
  expect(result.bytes).not.toEqual(bytes)
  expect((await PDFDocument.load(result.bytes)).getPageCount()).toBe(1)
  expect(await chainPdfium(async () => 42)).toBe(42)
})

it('keeps font outputs byte-identical for concurrent calls and recovers after bad input', async () => {
  const font = fixtureFont()
  const expected = await subsetTtf(font, 'Synthetic ABC')
  const order: string[] = []
  const first = subsetTtf(font, 'Synthetic ABC').then((value) => {
    order.push('first')
    return value
  })
  const failed = subsetTtf(Buffer.from('invalid synthetic font'), 'A')
  const rejection = expect(failed).rejects.toThrow()
  const last = subsetTtf(font, 'Synthetic ABC').then((value) => {
    order.push('last')
    return value
  })
  await rejection
  expect(await first).toEqual(expected)
  expect(await last).toEqual(expected)
  expect(order).toEqual(['first', 'last'])
  expect(expected.readUInt32BE(0)).toBe(0x00010000)
})
