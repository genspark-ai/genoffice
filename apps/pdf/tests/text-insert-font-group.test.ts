// @vitest-environment node
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFRawStream,
  StandardFonts,
  decodePDFRawStream,
} from 'pdf-lib'
import { afterEach, expect, it, vi } from 'vitest'
import { applyTextInserts, canDrawText } from '../src/main/text-edit'
import { subsetTtf } from '../src/main/font-subset'
import { savePdfToPath } from '../src/main/save-pdf'
import type { TextInsertInput } from '../src/shared/ipc'

// Redirect selected-font reads to a public fixture; never read an installed font.
vi.mock('node:fs', async (importOriginal) => {
  const real = await importOriginal<typeof import('node:fs')>()
  return {
    ...real,
    readFileSync: ((path: unknown, ...args: unknown[]) => {
      const value = String(path)
      if (/(?:\/Library\/Fonts|\/usr\/share\/fonts|Windows[\\/]Fonts)/i.test(value)) {
        if (value.endsWith('/Arial.ttf'))
          return real.readFileSync(
            resolve(__dirname, '../../../packages/ui/src/fonts/Carlito-Regular.ttf'),
          )
        throw Object.assign(new Error('No installed font fixture'), { code: 'ENOENT' })
      }
      return (real.readFileSync as (...args: unknown[]) => unknown)(path, ...args)
    }) as typeof real.readFileSync,
  }
})
vi.mock('../src/main/font-locate', async () => {
  const { fontCoversText } = await import('../src/main/font-cmap')
  const fixture = () =>
    readFileSync(resolve(__dirname, '../../../packages/ui/src/fonts/Carlito-Regular.ttf'))
  return {
    findSystemFont: () => fixture(),
    findFontCovering: (text: string) => {
      const font = fixture()
      return fontCoversText(font, text) ? font : null
    },
    isTruetype: (bytes: Buffer) => bytes.readUInt32BE(0) === 0x00010000,
  }
})
vi.mock('../src/main/font-subset', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/main/font-subset')>()
  return { ...real, subsetTtf: vi.fn(real.subsetTtf) }
})
afterEach(() => vi.mocked(subsetTtf).mockClear())
const insert = (text: string, y = 700, over: Partial<TextInsertInput> = {}): TextInsertInput => ({
  pageIndex: 0,
  origin: [50, y],
  text,
  fontSize: 14,
  font: 'arial',
  color: [0, 0, 0],
  ...over,
})
async function fixture() {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  doc.addPage([595, 842]).drawText('Existing first page', { x: 50, y: 780, size: 14, font })
  doc.addPage([595, 842]).drawText('Untouched second page', { x: 50, y: 780, size: 14, font })
  return doc.save({ useObjectStreams: false })
}
async function texts(bytes: Uint8Array) {
  const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const task = getDocument({
    data: new Uint8Array(bytes),
    useSystemFonts: false,
    disableFontFace: true,
  })
  const doc = await task.promise
  try {
    const out: string[] = []
    for (let p = 1; p <= doc.numPages; p++) {
      const content = await (await doc.getPage(p)).getTextContent()
      out.push(content.items.map((i) => ('str' in i ? i.str : '')).join(''))
    }
    return out
  } finally {
    await task.destroy()
  }
}
async function content(bytes: Uint8Array, index: number) {
  const doc = await PDFDocument.load(bytes)
  const value = doc.getPage(index).node.Contents()
  const streams =
    value instanceof PDFArray
      ? value.asArray().map((ref) => doc.context.lookup(ref, PDFRawStream))
      : [value as PDFRawStream]
  return streams.map((stream) =>
    Buffer.from(decodePDFRawStream(stream).decode()).toString('base64'),
  )
}
async function embeddedFontCount(bytes: Uint8Array) {
  const doc = await PDFDocument.load(bytes)
  const fonts = doc.getPage(0).node.Resources()!.lookup(PDFName.of('Font'), PDFDict)
  return fonts
    .entries()
    .filter(
      ([, value]) =>
        doc.context.lookup(value, PDFDict).lookup(PDFName.of('Subtype'), PDFName).decodeText() ===
        'Type0',
    ).length
}

it('skips only the unsupported entry and shares one subset across valid entries', async () => {
  const bytes = await fixture(),
    original = bytes.slice()
  expect(canDrawText('VALID_LINE_1', 'arial')).toBe(true)
  expect(canDrawText('Hi 😀', 'arial')).toBe(false)
  const result = await applyTextInserts(bytes, [
    insert('VALID_LINE_1'),
    insert('VALID_LINE_2', 650),
    insert('Hi 😀', 600),
  ])
  expect(result.skipped.map((x) => x.editIndex)).toEqual([2])
  expect(result.skipped[0]!.reason).toContain('no available font')
  expect(await texts(result.bytes)).toEqual([
    'Existing first pageVALID_LINE_1VALID_LINE_2',
    'Untouched second page',
  ])
  expect(bytes).toEqual(original)
  expect(await content(result.bytes, 1)).toEqual(await content(bytes, 1))
  expect(await embeddedFontCount(result.bytes)).toBe(1)
  expect(subsetTtf).toHaveBeenCalledTimes(1)
})
it('saves and reads back valid entries while preserving the original source', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'pdf-insert-group-'))
  try {
    const bytes = await fixture(),
      src = join(dir, 'source.pdf'),
      dst = join(dir, 'saved.pdf')
    writeFileSync(src, bytes)
    const result = await savePdfToPath(src, dst, {
      path: src,
      targetPath: dst,
      markups: [],
      drawings: [],
      stamps: [],
      formValues: [],
      textInserts: [insert('VALID_LINE_1'), insert('VALID_LINE_2', 650), insert('Hi 😀', 600)],
    })
    expect(result.skippedTextInserts.map((x) => x.editIndex)).toEqual([2])
    expect(readFileSync(src)).toEqual(Buffer.from(bytes))
    const saved = readFileSync(dst)
    expect(await texts(saved)).toEqual([
      'Existing first pageVALID_LINE_1VALID_LINE_2',
      'Untouched second page',
    ])
    expect(await embeddedFontCount(saved)).toBe(1)
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
})
it('keeps the fast shared-font path for an entirely drawable group', async () => {
  const result = await applyTextInserts(await fixture(), [insert('ONE'), insert('TWO', 650)])
  expect(result.skipped).toEqual([])
  expect(subsetTtf).toHaveBeenCalledTimes(1)
  expect(await embeddedFontCount(result.bytes)).toBe(1)
})
it('leaves all-unsupported groups unchanged and reports each entry once', async () => {
  const bytes = await fixture()
  const result = await applyTextInserts(bytes, [insert('😀'), insert('😀', 650)])
  expect(result.skipped.map((x) => x.editIndex)).toEqual([0, 1])
  expect(result.bytes).toBe(bytes)
  expect(subsetTtf).not.toHaveBeenCalled()
})
it('does not retry or reinterpret unrelated subset failures as coverage failures', async () => {
  vi.mocked(subsetTtf).mockRejectedValueOnce(new Error('synthetic subset failure'))
  const bytes = await fixture()
  const result = await applyTextInserts(bytes, [insert('ONE'), insert('TWO', 650)])
  expect(result.bytes).toBe(bytes)
  expect(result.skipped.map((x) => [x.editIndex, x.reason])).toEqual([
    [0, 'synthetic subset failure'],
    [1, 'synthetic subset failure'],
  ])
  expect(subsetTtf).toHaveBeenCalledTimes(1)
})

it('reports later subset failure only for the remaining valid entries', async () => {
  vi.mocked(subsetTtf).mockRejectedValueOnce(new Error('synthetic retry failure'))
  const bytes = await fixture()
  const result = await applyTextInserts(bytes, [
    insert('ONE'),
    insert('TWO', 650),
    insert('Hi 😀', 600),
  ])
  expect(result.bytes).toBe(bytes)
  expect(result.skipped.map((x) => x.editIndex).sort()).toEqual([0, 1, 2])
  expect(result.skipped.filter((x) => x.editIndex !== 2).map((x) => x.reason)).toEqual([
    'synthetic retry failure',
    'synthetic retry failure',
  ])
  expect(result.skipped.find((x) => x.editIndex === 2)!.reason).toContain('no available font')
  expect(subsetTtf).toHaveBeenCalledTimes(1)
})
it('does not classify unrelated errors by their message', async () => {
  const reason = 'the replacement contains characters no available font can draw'
  vi.mocked(subsetTtf).mockRejectedValueOnce(new Error(reason))
  const result = await applyTextInserts(await fixture(), [insert('ONE'), insert('TWO', 650)])
  expect(result.skipped.map((x) => x.editIndex)).toEqual([0, 1])
  expect(subsetTtf).toHaveBeenCalledTimes(1)
})
it('keeps original indexes across invalid pages, multiple faces and multiline inserts', async () => {
  const bytes = await fixture()
  const result = await applyTextInserts(bytes, [
    insert('Hi 😀', 700, { pageIndex: 8 }),
    insert('ONE\nTWO'),
    insert('Hi 😀', 600),
    insert('THREE', 600, { pageIndex: 1, font: 'times' }),
  ])
  expect(result.skipped.map((x) => x.editIndex)).toEqual([0, 2])
  const read = await texts(result.bytes)
  expect(read[0]).toContain('ONETWO')
  expect(read[1]).toContain('THREE')
  expect(subsetTtf).toHaveBeenCalledTimes(2)
})
