import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRef, PDFString } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applySaveRequest } from '../src/main/save-pdf'
import type { DrawingInput, SavePdfRequest } from '../src/shared/ipc'

type NoteInput = Extract<DrawingInput, { kind: 'note' }>

const NOTE_RECT: [number, number, number, number] = [100, 682, 120, 700]

/** One-page PDF with a saved Text comment, as WPS/Acrobat would leave it */
async function fixtureWithComment(): Promise<{ bytes: Uint8Array; objNum: number }> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([595, 842])
  const annot = doc.context.obj({
    Type: 'Annot',
    Subtype: 'Text',
    Rect: [...NOTE_RECT],
    Name: 'Comment',
    F: 4,
    P: page.ref,
  })
  annot.set(PDFName.of('Contents'), PDFHexString.fromText('Original comment'))
  annot.set(PDFName.of('T'), PDFHexString.fromText('WPS User'))
  annot.set(PDFName.of('CreationDate'), PDFString.of("D:20260810120000+08'00'"))
  const ref = doc.context.register(annot)
  page.node.set(PDFName.of('Annots'), doc.context.obj([ref]))
  return { bytes: await doc.save({ useObjectStreams: false }), objNum: ref.objectNumber }
}

const request = (drawings: DrawingInput[]): SavePdfRequest => ({
  path: 'test.pdf',
  markups: [],
  drawings,
  formValues: [],
  stamps: [],
})

const note = (over: Partial<NoteInput>): NoteInput => ({
  kind: 'note',
  pageIndex: 0,
  color: [1, 0.78, 0.13],
  at: [200, 500],
  contents: 'note',
  ...over,
})

interface ParsedAnnot {
  ref: PDFRef
  dict: PDFDict
  contents: string
  author: string
  irt: PDFRef | null
}

/** All ref-backed Text annots of page 1 in the saved bytes */
async function textAnnots(bytes: Uint8Array): Promise<ParsedAnnot[]> {
  const doc = await PDFDocument.load(bytes)
  const page = doc.getPage(0)
  const annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) return []
  const out: ParsedAnnot[] = []
  for (let i = 0; i < annots.size(); i++) {
    const ref = annots.get(i)
    if (!(ref instanceof PDFRef)) continue
    const dict = doc.context.lookupMaybe(ref, PDFDict)
    if (!dict || dict.lookupMaybe(PDFName.of('Subtype'), PDFName) !== PDFName.of('Text')) continue
    const text = (key: string): string => {
      const v = dict.lookup(PDFName.of(key))
      return v instanceof PDFString || v instanceof PDFHexString ? v.decodeText() : ''
    }
    const irt = dict.get(PDFName.of('IRT'))
    out.push({
      ref,
      dict,
      contents: text('Contents'),
      author: text('T'),
      irt: irt instanceof PDFRef ? irt : null,
    })
  }
  return out
}

describe('note reply save', () => {
  it('writes author and timestamps on a new root note', async () => {
    const { bytes } = await fixtureWithComment()
    const { bytes: out } = await applySaveRequest(
      bytes,
      request([note({ contents: '你好，批注', author: 'hong', createdMs: 1786511400000 })]),
    )
    const added = (await textAnnots(out)).find((a) => a.contents === '你好，批注')!
    expect(added.author).toBe('hong')
    expect(added.irt).toBeNull()
    const created = added.dict.lookup(PDFName.of('CreationDate'))
    const modified = added.dict.lookup(PDFName.of('M'))
    expect(created instanceof PDFString && created.decodeText()).toMatch(/^D:20\d{12}[+-]/)
    expect(modified instanceof PDFString && modified.decodeText()).toMatch(/^D:20\d{12}[+-]/)
  })

  it('falls back to the GenOffice author when none is provided', async () => {
    const { bytes } = await fixtureWithComment()
    const { bytes: out } = await applySaveRequest(bytes, request([note({ contents: 'anon' })]))
    expect((await textAnnots(out)).find((a) => a.contents === 'anon')!.author).toBe('GenOffice')
  })

  it('links a reply to a saved comment via /IRT even with a stale object number', async () => {
    const { bytes } = await fixtureWithComment()
    const { bytes: out } = await applySaveRequest(
      bytes,
      request([
        note({
          contents: 'my reply',
          author: 'hong',
          // objNum hint deliberately wrong: rect + contents must still find the parent
          replyToSaved: { objNum: 9999, rect: NOTE_RECT, contents: 'Original comment' },
        }),
      ]),
    )
    const all = await textAnnots(out)
    const reply = all.find((a) => a.contents === 'my reply')!
    const parent = all.find((a) => a.contents === 'Original comment')!
    expect(reply.irt).toBe(parent.ref)
    expect(reply.dict.lookup(PDFName.of('RT'))).toBe(PDFName.of('R'))
  })

  it('parents a reply onto a note written in the same request via localId', async () => {
    const { bytes } = await fixtureWithComment()
    const { bytes: out } = await applySaveRequest(
      bytes,
      request([
        note({ contents: 'new root', localId: 'p1' }),
        note({ contents: 'new reply', localId: 'p2', replyToLocalId: 'p1' }),
      ]),
    )
    const all = await textAnnots(out)
    const root = all.find((a) => a.contents === 'new root')!
    const reply = all.find((a) => a.contents === 'new reply')!
    expect(reply.irt).toBe(root.ref)
    expect(root.irt).toBeNull()
  })

  it('degrades a reply to a root note when the saved parent cannot be found', async () => {
    const { bytes } = await fixtureWithComment()
    const { bytes: out } = await applySaveRequest(
      bytes,
      request([
        note({
          contents: 'stranded reply',
          replyToSaved: { objNum: 9999, rect: [1, 1, 9, 9], contents: 'gone' },
        }),
      ]),
    )
    const reply = (await textAnnots(out)).find((a) => a.contents === 'stranded reply')!
    expect(reply.irt).toBeNull()
    expect(reply.dict.get(PDFName.of('RT'))).toBeUndefined()
  })
})
