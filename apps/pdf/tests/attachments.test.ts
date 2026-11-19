import { describe, expect, it } from 'vitest'
import { PDFDocument, PDFName, PDFRawStream, PDFString } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { applyAttachments, attachmentLabels, AttachmentsError } from '../src/main/pdf-attachments'
import { applySaveRequest } from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

const b64 = (text: string): string => Buffer.from(text, 'utf8').toString('base64')

const emptyRequest = (): SavePdfRequest => ({
  path: 'fixture.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
})

async function attachmentsOf(bytes: Uint8Array): Promise<Map<string, number>> {
  const doc = await getDocument({ data: bytes.slice() }).promise
  const all = (await doc.getAttachments()) ?? new Map()
  const out = new Map<string, number>()
  for (const [name, file] of all) {
    // the content id IS the name-tree key (the attachment's name)
    const content = file.content ?? (await doc.getAttachmentContent(name))
    out.set(name, content?.byteLength ?? 0)
  }
  return out
}

describe('embedded files through the save pipeline', () => {
  it('adds two attachments in one save and pdf.js lists them with their sizes', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const { bytes } = await applySaveRequest(await doc.save(), {
      ...emptyRequest(),
      attachments: {
        add: [
          { name: 'a.txt', bytesBase64: b64('first file') },
          { name: 'b.txt', bytesBase64: b64('second file with more bytes') },
        ],
      },
    })
    const listed = await attachmentsOf(bytes)
    expect(listed.get('a.txt')).toBe('first file'.length)
    expect(listed.get('b.txt')).toBe('second file with more bytes'.length)
  })

  it('removes a pre-existing attachment and rejects an unknown one', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    applyAttachments(doc, { add: [{ name: 'keep.pdf', bytesBase64: b64('kept') }] })
    const seeded = await doc.save()
    expect((await attachmentsOf(seeded)).has('keep.pdf')).toBe(true)

    const { bytes } = await applySaveRequest(seeded, {
      ...emptyRequest(),
      attachments: { remove: ['keep.pdf'] },
    })
    expect((await attachmentsOf(bytes)).size).toBe(0)
    await expect(
      applySaveRequest(seeded, { ...emptyRequest(), attachments: { remove: ['nope.pdf'] } }),
    ).rejects.toBeInstanceOf(AttachmentsError)
  })

  it('fails a save that reuses an existing name, leaving the tree untouched', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    applyAttachments(doc, { add: [{ name: 'dup.txt', bytesBase64: b64('original') }] })
    const seeded = await doc.save()
    await expect(
      applySaveRequest(seeded, {
        ...emptyRequest(),
        attachments: { add: [{ name: 'dup.txt', bytesBase64: b64('other') }] },
      }),
    ).rejects.toBeInstanceOf(AttachmentsError)
    expect(await attachmentsOf(seeded)).toBeInstanceOf(Map)
  })

  it('keys a seeded attachment by its name-tree key even when /F is a literal string', async () => {
    // Most producers write the key and /F as literal strings; keying by /UF
    // turned the label into "(report.xlsx)" (parentheses and all), which no
    // longer matched what pdf.js reports and made removing it throw.
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    const context = doc.context
    const streamDict = context.obj({ Type: 'EmbeddedFile', Length: 5 })
    const streamRef = context.nextRef()
    context.assign(streamRef, PDFRawStream.of(streamDict, new Uint8Array(Buffer.from('bytes'))))
    const spec = context.obj({
      Type: 'Filespec',
      F: PDFString.of('report.xlsx'),
      UF: PDFString.of('report.xlsx'),
      EF: context.obj({ F: streamRef, UF: streamRef }),
    })
    const names = context.obj({
      Names: context.obj([PDFString.of('report.xlsx'), context.register(spec)]),
    })
    doc.catalog.set(
      PDFName.of('Names'),
      context.register(context.obj({ EmbeddedFiles: context.register(names) })),
    )
    expect(attachmentLabels(doc)).toEqual(['report.xlsx'])
    const seeded = await doc.save()

    const { bytes } = await applySaveRequest(seeded, {
      ...emptyRequest(),
      attachments: { remove: ['report.xlsx'] },
    })
    expect((await attachmentsOf(bytes)).size).toBe(0)
  })
})

describe('the embedded-files tree helper', () => {
  it('labels a tree it built and refuses a /Kids tree', async () => {
    const doc = await PDFDocument.create()
    doc.addPage([300, 300])
    applyAttachments(doc, { add: [{ name: 'z.txt', bytesBase64: b64('zzz') }] })
    expect(attachmentLabels(doc)).toEqual(['z.txt'])

    const context = doc.context
    const kidsNode = context.obj({
      Kids: [context.register(context.obj({ Names: [] }))],
    })
    doc.catalog.set(PDFName.of('Names'), context.obj({ EmbeddedFiles: context.register(kidsNode) }))
    expect(() =>
      applyAttachments(doc, { add: [{ name: 'x.txt', bytesBase64: b64('x') }] }),
    ).toThrow(AttachmentsError)
  })
})
