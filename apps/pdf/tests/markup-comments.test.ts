import { expect, it } from 'vitest'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFRef,
  PDFRawStream,
} from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { loadSavedAnnots } from '../src/renderer/annotation-catalog'
import {
  findThreadRoot,
  flattenThread,
  savedNoteKey,
  threadSubtree,
  visibleNoteThreads,
} from '../src/renderer/note-threads'
import { applySaveRequest } from '../src/main/save-pdf'
import type { MarkupType, SavePdfRequest } from '../src/shared/ipc'

const rect: [number, number, number, number] = [50, 50, 200, 70]
const names: Record<MarkupType, string> = {
  highlight: 'Highlight',
  underline: 'Underline',
  strikeout: 'StrikeOut',
  squiggly: 'Squiggly',
}
async function fixture(
  type: MarkupType,
  contents = 'This comment should be visible',
  other = false,
) {
  const doc = await PDFDocument.create()
  const page = doc.addPage([595, 842])
  page.drawText('Synthetic highlighted text', { x: 50, y: 60, size: 12 })
  const refs: PDFRef[] = []
  for (const text of other ? [contents, 'Keep this other comment'] : [contents]) {
    const annotation = doc.context.obj({
      Type: 'Annot',
      Subtype: names[type],
      Rect: rect,
      QuadPoints: [50, 70, 200, 70, 50, 50, 200, 50],
      C: [1, 1, 0],
      F: 4,
      P: page.ref,
    })
    annotation.set(PDFName.of('Contents'), PDFHexString.fromText(text))
    annotation.set(PDFName.of('T'), PDFHexString.fromText('Reviewer'))
    refs.push(doc.context.register(annotation))
  }
  page.node.set(PDFName.of('Annots'), doc.context.obj(refs))
  return doc.save({ useObjectStreams: false })
}
async function catalog(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), useSystemFonts: true })
  const pdf = await task.promise
  try {
    return await loadSavedAnnots(pdf, 0)
  } finally {
    await task.destroy()
  }
}
const request = (extra: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: 'synthetic.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...extra,
})

it.each(['highlight', 'underline', 'strikeout', 'squiggly'] as const)(
  'reads %s Contents as one thread while retaining the markup and stable click identity',
  async (type) => {
    const bytes = await fixture(type)
    const first = await catalog(bytes),
      second = await catalog(bytes)
    expect(first.markups).toHaveLength(1)
    expect(first.notes).toHaveLength(1)
    expect(first.notes[0]).toMatchObject({
      type,
      author: 'Reviewer',
      contents: 'This comment should be visible',
    })
    const threads = visibleNoteThreads(first.notes, [], new Set(), new Map())
    expect(threads).toHaveLength(1)
    expect(threads[0].key).toBe(savedNoteKey(first.markups[0].objNum))
    expect(findThreadRoot(threads, savedNoteKey(first.markups[0].objNum))?.key).toBe(threads[0].key)
    expect(second.notes[0].objNum).toBe(first.notes[0].objNum)
  },
)

it('does not manufacture a thread for a blank markup Contents', async () => {
  const parsed = await catalog(await fixture('highlight', '  '))
  expect(parsed.markups).toHaveLength(1)
  expect(parsed.notes).toEqual([])
})

it.each(['highlight', 'underline', 'strikeout', 'squiggly'] as const)(
  'edits %s Contents and saves a reply to the existing markup object, without duplicate threads',
  async (type) => {
    const bytes = await fixture(type)
    const root = (await catalog(bytes)).notes[0]!
    expect(root).toBeDefined()
    const target = {
      objNum: root.objNum,
      subtype: root.type,
      quadPoints: root.quadPoints,
      rect: root.rect,
      contents: root.contents,
    }
    const { bytes: output } = await applySaveRequest(
      bytes,
      request({
        drawings: [
          {
            kind: 'note',
            pageIndex: 0,
            at: [50, 70],
            color: [1, 1, 0],
            contents: 'Reply',
            author: 'Second reviewer',
            replyToSaved: target,
          },
        ],
        noteEdits: [
          {
            pageIndex: 0,
            objNum: root.objNum,
            subtype: root.type,
            quadPoints: root.quadPoints,
            rect: root.rect,
            oldContents: root.contents,
            contents: 'Edited highlight comment',
          },
        ],
      }),
    )
    const reopened = await catalog(output)
    expect(reopened.markups).toHaveLength(1)
    const threads = visibleNoteThreads(reopened.notes, [], new Set(), new Map())
    expect(threads).toHaveLength(1)
    expect(threads[0].contents).toBe('Edited highlight comment')
    expect(flattenThread(threads[0]).map(({ item }) => item.contents)).toEqual([
      'Edited highlight comment',
      'Reply',
    ])
    expect(new Set(flattenThread(threads[0]).map(({ item }) => item.key)).size).toBe(2)
    const pdf = await PDFDocument.load(output)
    const dict = pdf.context.lookup(PDFRef.of(root.objNum)) as PDFDict
    expect(dict.lookup(PDFName.of('Subtype'), PDFName)).toBe(PDFName.of(names[type]))
    expect(
      dict
        .lookup(PDFName.of('QuadPoints'), PDFArray)
        .asArray()
        .map((value) => Number(value.toString())),
    ).toEqual([50, 70, 200, 70, 50, 50, 200, 50])
    expect(threads[0].author).toBe('Reviewer')
    expect(threads[0].saved?.objNum).toBe(root.objNum)
    const contentBytes = (doc: PDFDocument) => {
      const contents = doc.getPage(0).node.lookup(PDFName.of('Contents'), PDFArray)
      return contents
        .asArray()
        .map((ref) => Array.from((doc.context.lookup(ref) as PDFRawStream).contents))
    }
    expect(contentBytes(pdf)).toEqual(contentBytes(await PDFDocument.load(bytes)))
    const deletes = threadSubtree(threads[0]).saved.map((annotation) => ({
      pageIndex: 0,
      objNum: annotation.objNum,
      subtype: annotation.type,
      quadPoints: annotation.quadPoints,
      rect: annotation.rect,
      contents: annotation.contents,
    }))
    const { bytes: removed } = await applySaveRequest(output, request({ annotDeletes: deletes }))
    const afterDelete = await catalog(removed)
    expect(afterDelete.notes).toEqual([])
    expect(afterDelete.markups).toEqual([])
  },
)

it('deletes a commented markup by its real subtype and contents, preserving a same-rect sibling', async () => {
  const bytes = await fixture('highlight', 'Delete only this comment', true)
  const root = visibleNoteThreads((await catalog(bytes)).notes, [], new Set(), new Map())[0]!
  expect(root).toBeDefined()
  const deleted = threadSubtree(root).saved.map((annotation) => ({
    pageIndex: 0,
    objNum: annotation.objNum,
    subtype: annotation.type,
    quadPoints: annotation.quadPoints,
    rect: annotation.rect,
    contents: annotation.contents,
  }))
  const { bytes: output } = await applySaveRequest(bytes, request({ annotDeletes: deleted }))
  const reopened = await catalog(output)
  expect(reopened.markups).toHaveLength(1)
  expect(reopened.notes.map((note) => note.contents)).toEqual(['Keep this other comment'])
})

it.each([100, Number.NaN])(
  'refuses an incorrect or non-finite markup identity (%s)',
  async (coordinate) => {
    const bytes = await fixture('highlight')
    const root = (await catalog(bytes)).notes[0]!
    expect(root).toBeDefined()
    const quadPoints = [...root.quadPoints!]
    quadPoints[0] = coordinate
    const target = {
      pageIndex: 0,
      objNum: root.objNum,
      subtype: root.type,
      rect: root.rect,
      quadPoints,
      contents: root.contents,
    }
    const { bytes: output } = await applySaveRequest(
      bytes,
      request({
        annotDeletes: [target],
        noteEdits: [
          { ...target, oldContents: root.contents, contents: 'Must not replace the original' },
        ],
      }),
    )
    const after = await catalog(output)
    expect(after.notes.map((note) => note.contents)).toEqual([root.contents])
    expect(after.markups).toHaveLength(1)
  },
)

it('keeps replies and deletion attached to a multi-line highlight', async () => {
  const doc = await PDFDocument.load(await fixture('highlight'))
  const ref = doc.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray).get(0)
  const annotation = doc.context.lookup(ref) as PDFDict
  annotation.set(
    PDFName.of('QuadPoints'),
    doc.context.obj([50, 70, 200, 70, 50, 50, 200, 50, 50, 40, 150, 40, 50, 20, 150, 20]),
  )
  annotation.set(PDFName.of('Rect'), doc.context.obj([50, 20, 200, 70]))
  const bytes = await doc.save({ useObjectStreams: false })
  const root = (await catalog(bytes)).notes[0]!
  expect(root.quadPoints).toHaveLength(16)
  const { bytes: output } = await applySaveRequest(
    bytes,
    request({
      drawings: [
        {
          kind: 'note',
          pageIndex: 0,
          at: [50, 70],
          color: [1, 1, 0],
          contents: 'Multi-line reply',
          replyToSaved: {
            objNum: root.objNum,
            subtype: root.type,
            quadPoints: root.quadPoints,
            rect: root.rect,
            contents: root.contents,
          },
        },
      ],
    }),
  )
  const threads = visibleNoteThreads((await catalog(output)).notes, [], new Set(), new Map())
  expect(threads).toHaveLength(1)
  expect(flattenThread(threads[0])).toHaveLength(2)
  const annotDeletes = threadSubtree(threads[0]).saved.map((note) => ({
    pageIndex: 0,
    objNum: note.objNum,
    subtype: note.type,
    rect: note.rect,
    contents: note.contents,
    quadPoints: note.quadPoints,
  }))
  const { bytes: removed } = await applySaveRequest(output, request({ annotDeletes }))
  expect((await catalog(removed)).notes).toEqual([])
})
