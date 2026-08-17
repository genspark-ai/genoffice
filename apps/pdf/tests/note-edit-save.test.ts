import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFRef, PDFString } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applySaveRequest } from '../src/main/save-pdf'
import type { NoteEditInput, SavePdfRequest } from '../src/shared/ipc'

const ROOT_RECT: [number, number, number, number] = [100, 682, 120, 700]
const REPLY_RECT: [number, number, number, number] = [100, 682, 120, 700]

/** One-page PDF with a saved comment thread: a root and one /IRT reply */
async function fixtureWithThread(): Promise<{
  bytes: Uint8Array
  rootObjNum: number
  replyObjNum: number
}> {
  const doc = await PDFDocument.create()
  const page = doc.addPage([595, 842])
  const root = doc.context.obj({
    Type: 'Annot',
    Subtype: 'Text',
    Rect: [...ROOT_RECT],
    Name: 'Comment',
    F: 4,
    P: page.ref,
  })
  root.set(PDFName.of('Contents'), PDFHexString.fromText('Original comment'))
  root.set(PDFName.of('T'), PDFHexString.fromText('WPS User'))
  root.set(PDFName.of('CreationDate'), PDFString.of("D:20260810120000+08'00'"))
  const rootRef = doc.context.register(root)
  const reply = doc.context.obj({
    Type: 'Annot',
    Subtype: 'Text',
    Rect: [...REPLY_RECT],
    Name: 'Comment',
    F: 4,
    P: page.ref,
    RT: 'R',
  })
  reply.set(PDFName.of('Contents'), PDFHexString.fromText('Saved reply'))
  reply.set(PDFName.of('T'), PDFHexString.fromText('Colleague'))
  reply.set(PDFName.of('IRT'), rootRef)
  const replyRef = doc.context.register(reply)
  page.node.set(PDFName.of('Annots'), doc.context.obj([rootRef, replyRef]))
  return {
    bytes: await doc.save({ useObjectStreams: false }),
    rootObjNum: rootRef.objectNumber,
    replyObjNum: replyRef.objectNumber,
  }
}

const request = (over: Partial<SavePdfRequest>): SavePdfRequest => ({
  path: 'test.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

const edit = (over: Partial<NoteEditInput>): NoteEditInput => ({
  pageIndex: 0,
  objNum: 0,
  rect: ROOT_RECT,
  oldContents: 'Original comment',
  contents: 'Rewritten comment',
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

describe('note edit save', () => {
  it('rewrites /Contents in place, stamps /M, and keeps saved replies attached', async () => {
    const { bytes, rootObjNum, replyObjNum } = await fixtureWithThread()
    const { bytes: out } = await applySaveRequest(
      bytes,
      request({ noteEdits: [edit({ objNum: rootObjNum })] }),
    )
    const all = await textAnnots(out)
    const root = all.find((a) => a.contents === 'Rewritten comment')!
    expect(root.ref.objectNumber).toBe(rootObjNum)
    expect(root.author).toBe('WPS User')
    const modified = root.dict.lookup(PDFName.of('M'))
    expect(modified instanceof PDFString && modified.decodeText()).toMatch(/^D:20\d{12}[+-]/)
    // The saved reply still points at the same (edited) parent object
    const reply = all.find((a) => a.contents === 'Saved reply')!
    expect(reply.ref.objectNumber).toBe(replyObjNum)
    expect(reply.irt?.objectNumber).toBe(rootObjNum)
  })

  it('edits a reply by its contents even though it shares the root rect', async () => {
    const { bytes, replyObjNum } = await fixtureWithThread()
    const { bytes: out } = await applySaveRequest(
      bytes,
      request({
        noteEdits: [
          edit({ objNum: replyObjNum, oldContents: 'Saved reply', contents: 'Reply v2' }),
        ],
      }),
    )
    const all = await textAnnots(out)
    expect(all.find((a) => a.contents === 'Reply v2')!.ref.objectNumber).toBe(replyObjNum)
    expect(all.find((a) => a.contents === 'Original comment')).toBeDefined()
  })

  it('matches by rect + old contents when the object number hint is stale', async () => {
    const { bytes } = await fixtureWithThread()
    const { bytes: out } = await applySaveRequest(
      bytes,
      request({ noteEdits: [edit({ objNum: 9999 })] }),
    )
    expect((await textAnnots(out)).some((a) => a.contents === 'Rewritten comment')).toBe(true)
  })

  it('is a no-op when the target cannot be found', async () => {
    const { bytes } = await fixtureWithThread()
    const { bytes: out } = await applySaveRequest(
      bytes,
      request({ noteEdits: [edit({ objNum: 9999, rect: [1, 1, 9, 9], oldContents: 'gone' })] }),
    )
    const all = await textAnnots(out)
    expect(all.map((a) => a.contents).sort()).toEqual(['Original comment', 'Saved reply'])
  })

  it('parents a new reply onto the old contents before the edit lands', async () => {
    const { bytes, rootObjNum } = await fixtureWithThread()
    // Same request: a reply targeting the pre-edit contents plus the edit itself —
    // the reply must still find its /IRT parent (drawings apply before noteEdits)
    const { bytes: out } = await applySaveRequest(
      bytes,
      request({
        drawings: [
          {
            kind: 'note',
            pageIndex: 0,
            color: [1, 0.78, 0.13],
            at: [100, 700],
            contents: 'fresh reply',
            replyToSaved: { objNum: rootObjNum, rect: ROOT_RECT, contents: 'Original comment' },
          },
        ],
        noteEdits: [edit({ objNum: rootObjNum })],
      }),
    )
    const all = await textAnnots(out)
    const reply = all.find((a) => a.contents === 'fresh reply')!
    expect(reply.irt?.objectNumber).toBe(rootObjNum)
    expect(all.some((a) => a.contents === 'Rewritten comment')).toBe(true)
  })
})
