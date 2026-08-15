import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFString } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applyAnnotDeletes } from '../src/main/annot-delete'

const NOTE_RECT: [number, number, number, number] = [100, 682, 120, 700]

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
  annot.set(PDFName.of('Contents'), PDFHexString.fromText('Delete me'))
  const ref = doc.context.register(annot)
  page.node.set(PDFName.of('Annots'), doc.context.obj([ref]))
  return { bytes: await doc.save({ useObjectStreams: false }), objNum: ref.objectNumber }
}

/** Thread fixture: root + two /IRT replies, all sharing the root's rect (real thread shape) */
async function fixtureWithThread(): Promise<{ bytes: Uint8Array; objNums: number[] }> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([595, 842])
  const refs = []
  for (const contents of ['root note', 'first reply', 'second reply']) {
    const annot = doc.context.obj({
      Type: 'Annot',
      Subtype: 'Text',
      Rect: [...NOTE_RECT],
      Name: 'Comment',
      F: 4,
      P: page.ref,
    })
    annot.set(PDFName.of('Contents'), PDFHexString.fromText(contents))
    if (refs.length > 0) {
      annot.set(PDFName.of('IRT'), refs[0]!)
      annot.set(PDFName.of('RT'), PDFName.of('R'))
    }
    refs.push(doc.context.register(annot))
  }
  page.node.set(PDFName.of('Annots'), doc.context.obj(refs))
  return {
    bytes: await doc.save({ useObjectStreams: false }),
    objNums: refs.map((r) => r.objectNumber),
  }
}

async function pageAnnotCount(bytes: Uint8Array): Promise<number> {
  const doc = await PDFDocument.load(bytes)
  const annots = doc.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  return annots?.size() ?? 0
}

async function pageAnnotContents(bytes: Uint8Array): Promise<string[]> {
  const doc = await PDFDocument.load(bytes)
  const annots = doc.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) return []
  const out: string[] = []
  for (let i = 0; i < annots.size(); i++) {
    const v = annots.lookup(i, PDFDict).lookup(PDFName.of('Contents'))
    out.push(v instanceof PDFString || v instanceof PDFHexString ? v.decodeText() : '')
  }
  return out
}

describe('annot-delete for note comments', () => {
  it('removes a saved Text annotation by object number', async () => {
    const { bytes, objNum } = await fixtureWithComment()
    const out = await applyAnnotDeletes(bytes, [
      { pageIndex: 0, objNum, subtype: 'note', rect: NOTE_RECT },
    ])
    expect(await pageAnnotCount(out)).toBe(0)
  })

  it('recovers via the rect fallback when the object number is stale', async () => {
    const { bytes } = await fixtureWithComment()
    const out = await applyAnnotDeletes(bytes, [
      { pageIndex: 0, objNum: 9999, subtype: 'note', rect: NOTE_RECT },
    ])
    expect(await pageAnnotCount(out)).toBe(0)
  })

  it('leaves the file untouched when nothing matches', async () => {
    const { bytes } = await fixtureWithComment()
    const out = await applyAnnotDeletes(bytes, [
      { pageIndex: 0, objNum: 9999, subtype: 'note', rect: [1, 1, 9, 9] },
    ])
    expect(await pageAnnotCount(out)).toBe(1)
  })

  it('deletes only the addressed comment of a same-rect thread (contents identity)', async () => {
    const { bytes, objNums } = await fixtureWithThread()
    const out = await applyAnnotDeletes(bytes, [
      {
        pageIndex: 0,
        objNum: objNums[2]!,
        subtype: 'note',
        rect: NOTE_RECT,
        contents: 'second reply',
      },
    ])
    expect(await pageAnnotContents(out)).toEqual(['root note', 'first reply'])
  })

  it('contents identity holds even when the object number hint is stale', async () => {
    const { bytes } = await fixtureWithThread()
    const out = await applyAnnotDeletes(bytes, [
      { pageIndex: 0, objNum: 9999, subtype: 'note', rect: NOTE_RECT, contents: 'first reply' },
    ])
    expect(await pageAnnotContents(out)).toEqual(['root note', 'second reply'])
  })
})
