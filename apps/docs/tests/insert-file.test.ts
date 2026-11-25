/** Insert > File: the body of another document goes in at the selection as a paste of it would. */
import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import {
  docxBookmarkRange,
  docxRootNamespaces,
  parseDocx,
  type NoteInfo,
} from '@genoffice/docx-engine'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
  TABLE_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { pmDocToSavePlan } from '../src/renderer/editor/convert'

// jsdom has no clipboard types; the paste path builds an event around the payload
class FakeDataTransfer {
  private data = new Map<string, string>()
  setData(type: string, value: string) {
    this.data.set(type, value)
  }
  getData(type: string) {
    return this.data.get(type) ?? ''
  }
  get types() {
    return [...this.data.keys()]
  }
}
class FakeClipboardEvent extends Event {
  clipboardData: FakeDataTransfer | null
  constructor(type: string, init?: { clipboardData?: FakeDataTransfer }) {
    super(type)
    this.clipboardData = init?.clipboardData ?? null
  }
}
const g = globalThis as Record<string, unknown>
g.DataTransfer = FakeDataTransfer
g.ClipboardEvent = FakeClipboardEvent

const { fileBodyPayload, insertFileBody } = await import('../src/renderer/editor/insert-file')

const para = (inner: string) => `<w:p>${inner}</w:p>`
const run = (text: string, rPr = '') => `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`

/** a rectangle anchored in a paragraph: a drawing that needs nothing but its own XML */
const SHAPE =
  '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ' +
  'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="1" name="Rectangle 1"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>' +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

function open(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'Before after' }] }],
    },
  })
}

/** the other document as Insert > File reads it */
const WPS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'

async function sourceOf(bytes: Uint8Array) {
  return { parsed: await parseDocx(bytes), namespaces: await docxRootNamespaces(bytes) }
}

const FOOTNOTES_XML =
  '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:footnote w:id="7"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t>Carried note</w:t></w:r></w:p></w:footnote></w:footnotes>'

const NOTE_MARK =
  '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="7"/></w:r>'

function notesHost(existing: NoteInfo[] = []) {
  const lists = { footnote: [...existing], endnote: [] as NoteInfo[] }
  return {
    lists,
    host: {
      list: (kind: 'footnote' | 'endnote') => lists[kind],
      add: (kind: 'footnote' | 'endnote', notes: NoteInfo[]) => void lists[kind].push(...notes),
    },
  }
}

describe('inserting a file', () => {
  it('carries text, a table, a picture and a shape, and counts what it leaves behind', async () => {
    const other = await sourceOf(
      await buildDocx({
        bodyXml:
          para(run('Bold start', '<w:rPr><w:b/></w:rPr>') + run(' then plain')) +
          TABLE_XML +
          IMAGE_PARAGRAPH_XML +
          SHAPE +
          SHAPE.replace('<wps:bodyPr/>', '<wps:bodyPr/><a:blip r:embed="rId9"/>') +
          para(run('The end')),
        withImage: true,
        docRootExtraAttrs: WPS,
      }),
    )
    const editor = open()
    const payload = fileBodyPayload(editor, other)
    // the shape with a picture of its own needs a part of the other file
    expect(payload.leftOut).toBe(1)
    expect(payload.carried.blocks.size).toBe(1)
    expect(payload.text).toContain('Bold start then plain')
    expect(payload.text).toContain('The end')
    expect(payload.html).toContain('<table')
    expect(payload.html).toContain('data:image/')
    editor.destroy()
  })

  it('puts the body in at the caret', async () => {
    const other = await sourceOf(
      await buildDocx({
        bodyXml: para(run('Inserted', '<w:rPr><w:b/></w:rPr>')) + para(run('Second line')),
      }),
    )
    const editor = open()
    editor.commands.setTextSelection(8)
    expect(insertFileBody(editor, other)).toBe(0)
    const text = editor.state.doc.textBetween(0, editor.state.doc.content.size, '|')
    expect(text).toContain('Inserted')
    expect(text).toContain('Second line')
    expect(text.indexOf('Before')).toBeLessThan(text.indexOf('Inserted'))
    expect(text.indexOf('Second line')).toBeLessThan(text.indexOf('after'))
    // the formatting came along
    let bold = false
    editor.state.doc.descendants((node) => {
      if (node.isText && node.text === 'Inserted')
        bold = node.marks.some((m) => m.type.name === 'bold')
    })
    expect(bold).toBe(true)
    editor.destroy()
  })

  it('brings a shape in as its own XML, which a save writes as it came', async () => {
    const other = await sourceOf(await buildDocx({ bodyXml: SHAPE, docRootExtraAttrs: WPS }))
    const editor = open()
    expect(insertFileBody(editor, other)).toBe(0)
    const shapes: Array<Record<string, unknown>> = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docProtected') shapes.push(node.attrs)
    })
    expect(shapes).toHaveLength(1)
    expect(shapes[0].docxIndex).toBeNull()
    const xml = String(shapes[0].genXml)
    expect(xml).toContain('<wps:wsp>')
    expect(xml).toContain(
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    )
    expect(xml).toContain('xmlns:wp=')
    const plan = pmDocToSavePlan(editor.getJSON() as never, [])
    expect(plan.saveBlocks.some((b) => b.kind === 'xml' && b.xml === xml)).toBe(true)
    // once the shape is changed it is saved the way a shape made here is
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docProtected') at = pos
    })
    const box = (shapes[0].textboxes as Array<Record<string, unknown>>)[0]
    editor.view.dispatch(
      editor.state.tr.setNodeAttribute(at, 'textboxes', [{ ...box, widthPx: 300 }]),
    )
    const edited = pmDocToSavePlan(editor.getJSON() as never, [])
    expect(edited.saveBlocks.some((b) => b.kind === 'xml' && b.xml === xml)).toBe(false)
    expect(edited.saveBlocks.some((b) => b.kind === 'xml' && b.xml.includes('<wps:wsp>'))).toBe(
      true,
    )
    editor.destroy()
  })

  it('brings a section break along with its page setup when asked to, and leaves it behind otherwise', async () => {
    const BREAK =
      '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId7"/>' +
      '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
      '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr></w:p>'
    const bytes = await buildDocx({ bodyXml: para(run('Wide part')) + BREAK + para(run('Tail')) })
    const other = await sourceOf(bytes)

    const without = open()
    expect(fileBodyPayload(without, other).leftOut).toBe(1)
    without.destroy()

    const editor = open()
    const sections = { count: 0 }
    expect(insertFileBody(editor, { ...other, sections })).toBe(0)
    expect(sections.count).toBe(1)
    const breaks: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docProtected') breaks.push(String(node.attrs.genXml))
    })
    expect(breaks).toHaveLength(1)
    expect(breaks[0]).toContain('<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>')
    expect(breaks[0]).toContain('<w:cols w:num="2" w:space="720"/>')
    expect(breaks[0]).not.toContain('headerReference')
    const plan = pmDocToSavePlan(editor.getJSON() as never, [])
    expect(plan.saveBlocks.some((b) => b.kind === 'xml' && b.xml === breaks[0])).toBe(true)
    // text only: the page setup stays behind
    const plain = { count: 0 }
    expect(insertFileBody(open(), { ...other, plain: true, sections: plain })).toBe(1)
    expect(plain.count).toBe(0)
    editor.destroy()
  })

  it('gives the footnotes of the other document ids of this one', async () => {
    const other = await sourceOf(
      await buildDocx({
        bodyXml:
          para(run('Claim') + NOTE_MARK) +
          para(run('Said again') + NOTE_MARK.replace('<w:rPr>', '<w:rPr><w:b/>')),
        extraParts: [
          {
            path: 'word/footnotes.xml',
            xml: FOOTNOTES_XML,
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
          },
        ],
        extraRels:
          '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      }),
    )
    expect(other.parsed.footnotes.map((n) => n.id)).toEqual(['7'])
    const editor = open()
    const { lists, host } = notesHost([{ id: '7', text: 'Already here' }])
    expect(insertFileBody(editor, other, host)).toBe(0)
    expect(lists.footnote.map((n) => [n.id, n.text.trim(), n.created])).toEqual([
      ['7', 'Already here', undefined],
      ['8', 'Carried note', true],
    ])
    expect(lists.footnote[1].richParas?.flat().some((r) => r.bold)).toBe(true)
    const refs: string[] = []
    const bold: boolean[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name !== 'docNoteRef') return
      refs.push(`${node.attrs.kind}:${node.attrs.id}`)
      bold.push(String(node.attrs.rawRPr).includes('<w:b/>'))
    })
    // the two marks of the one note still share a note
    expect(refs).toEqual(['footnote:8', 'footnote:8'])
    // and each keeps its own run formatting
    expect(bold).toEqual([false, true])
    editor.destroy()

    // with nowhere to put the note, the reference does not come without it
    const bare = open()
    expect(insertFileBody(bare, other)).toBe(0)
    let orphan = 0
    bare.state.doc.descendants((node) => {
      if (node.type.name === 'docNoteRef') orphan++
    })
    expect(orphan).toBe(0)
    bare.destroy()
  })

  it('takes the inserted body and the new numbers of its notes back in one undo', async () => {
    const other = await sourceOf(
      await buildDocx({
        bodyXml: para(run('Claim') + NOTE_MARK),
        extraParts: [
          {
            path: 'word/footnotes.xml',
            xml: FOOTNOTES_XML,
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
          },
        ],
        extraRels:
          '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'Before' },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '1', num: 1 } },
            ],
          },
          { type: 'docParagraph', content: [{ type: 'text', text: 'After' }] },
        ],
      },
    })
    editor.commands.setTextSelection(editor.state.doc.content.size - 1)
    const before = editor.state.doc
    const marks = () => {
      const found: string[] = []
      editor.state.doc.descendants((node) => {
        if (node.type.name === 'docNoteRef') found.push(`${node.attrs.id}=${node.attrs.num}`)
      })
      return found
    }
    const { host } = notesHost([{ id: '1', text: 'Already here' }])
    expect(insertFileBody(editor, other, host)).toBe(0)
    // the mark came in as number 1 of its own document
    expect(marks()).toEqual(['1=1', '2=2'])
    editor.commands.undo()
    expect(editor.state.doc.eq(before)).toBe(true)
    editor.commands.redo()
    expect(marks()).toEqual(['1=1', '2=2'])
    editor.destroy()
  })

  it('puts the range of a bookmark inside a paragraph in as text', async () => {
    const bytes = await buildDocx({
      bodyXml:
        para(
          run('skip ') +
            '<w:bookmarkStart w:id="1" w:name="Clause"/>' +
            run('the clause') +
            '<w:bookmarkEnd w:id="1"/>' +
            run(' skip'),
        ) + para(run('Other paragraph')),
    })
    const range = await docxBookmarkRange(bytes, 'Clause')
    const editor = open()
    editor.commands.setTextSelection(8)
    expect(
      insertFileBody(editor, { ...(await sourceOf(range!.bytes)), inline: range!.inline }),
    ).toBe(0)
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.textContent).toBe('Before the clauseafter')
    editor.destroy()
  })

  it('has nothing to insert from a document of charts only', async () => {
    const other = await sourceOf(
      await buildDocx({
        bodyXml: SHAPE.replace('<wps:bodyPr/>', '<wps:bodyPr/><a:blip r:embed="rId9"/>'),
        docRootExtraAttrs: WPS,
      }),
    )
    const editor = open()
    expect(insertFileBody(editor, other)).toBeNull()
    expect(editor.state.doc.textContent).toBe('Before after')
    editor.destroy()
  })
})
