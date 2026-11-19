import { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { generateCaptionXml, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { allowDirtyFields } from '../../../packages/docx-engine/tests/helpers/dirty-fields'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  addBookmark,
  collectBookmarks,
  crossRefField,
  insertCrossReference,
  listRefTargets,
  removeBookmark,
  sortBookmarks,
  validBookmarkName,
} from '../src/renderer/editor/cross-reference'
import { collectCaptions } from '../src/renderer/commands/reference-commands'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { listGoToItems } from '../src/renderer/editor/goto'

const WORDS = { above: 'above', below: 'below', onPage: (n: number) => `on page ${n}` }

const BODY =
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Scope</w:t></w:r></w:p>' +
  '<w:p><w:bookmarkStart w:id="1" w:name="Zeta"/><w:bookmarkEnd w:id="1"/>' +
  '<w:bookmarkStart w:id="2" w:name="_Toc9"/><w:bookmarkEnd w:id="2"/>' +
  '<w:r><w:t>Zeta paragraph.</w:t></w:r></w:p>' +
  '<w:p><w:bookmarkStart w:id="3" w:name="Alpha"/><w:bookmarkEnd w:id="3"/>' +
  '<w:r><w:t>Alpha paragraph.</w:t></w:r></w:p>' +
  generateCaptionXml('Figure', 1, 'Plan') +
  '<w:p><w:r><w:t>Last paragraph.</w:t></w:r></w:p>'

async function open(bodyXml = BODY) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

async function saved(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const bytes = await saveDocx(parsed, plan.saveBlocks)
  const zip = await JSZip.loadAsync(bytes)
  return { bytes, xml: await zip.file('word/document.xml')!.async('string') }
}

function caretAtEnd(editor: Editor) {
  const end = editor.state.doc.content.size - 1
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(end))),
  )
}

describe('Bookmark dialog model', () => {
  it('lists by name or by location, with Word internals only on request', async () => {
    const { editor } = await open()
    const { doc } = editor.state
    expect(collectBookmarks(doc).map((b) => b.name)).toEqual(['Zeta', 'Alpha'])
    expect(sortBookmarks(collectBookmarks(doc), 'name').map((b) => b.name)).toEqual([
      'Alpha',
      'Zeta',
    ])
    expect(collectBookmarks(doc, true).map((b) => [b.name, b.hidden])).toEqual([
      ['_Toc9', true],
      ['Zeta', false],
      ['Alpha', false],
    ])
    editor.destroy()
  })

  it('takes the names Word takes', () => {
    expect(validBookmarkName('Chapter_2')).toBe(true)
    expect(validBookmarkName('2nd')).toBe(false)
    expect(validBookmarkName('two words')).toBe(false)
    expect(validBookmarkName('a'.repeat(41))).toBe(false)
  })

  it('adding a name already in use moves the bookmark, and the saved bookmark spans the text', async () => {
    const { editor, parsed } = await open()
    caretAtEnd(editor)
    expect(addBookmark(editor, 'Zeta')).toBe(true)
    const all = collectBookmarks(editor.state.doc)
    expect(all.map((b) => [b.name, b.preview])).toEqual([
      ['Alpha', 'Alpha paragraph.'],
      ['Zeta', 'Last paragraph.'],
    ])
    removeBookmark(editor, all[0])
    const { xml, bytes } = await saved(editor, parsed)
    expect(xml).not.toContain('w:name="Alpha"')
    expect(xml).toMatch(
      /<w:bookmarkStart w:id="(\d+)" w:name="Zeta"\/><w:r>(?:(?!<\/w:p>).)*Last paragraph\.<\/w:t><\/w:r><w:bookmarkEnd w:id="\1"\/><\/w:p>/,
    )
    const reparsed = await parseDocx(bytes)
    expect(reparsed.blocks.flatMap((b) => b.bookmarks ?? [])).toEqual(['Zeta'])
    editor.destroy()
  })
})

describe('Cross-reference dialog model', () => {
  it('a heading reference gives the heading a _Ref bookmark and saves a REF field to it', async () => {
    const { editor, parsed } = await open()
    caretAtEnd(editor)
    const [heading] = listRefTargets(editor, parsed.blocks, { kind: 'heading' })
    expect(heading).toMatchObject({ text: 'Scope', bookmark: null, level: 1 })
    expect(
      insertCrossReference(editor, parsed.blocks, {
        target: heading,
        content: 'text',
        hyperlink: true,
        includeAboveBelow: false,
        pages: null,
        words: WORDS,
      }),
    ).toBe(true)
    // the second reference to the same heading reuses its bookmark
    const again = listRefTargets(editor, parsed.blocks, { kind: 'heading' })[0]
    expect(again.bookmark).toMatch(/^_Ref\d{9}$/)
    const { xml, bytes } = await saved(editor, parsed)
    const name = again.bookmark!
    expect(xml).toMatch(
      new RegExp(
        `<w:bookmarkStart w:id="(\\d+)" w:name="${name}"/><w:r>.*?Scope</w:t></w:r><w:bookmarkEnd w:id="\\1"/>`,
      ),
    )
    expect(xml).toContain(` REF ${name} \\h `)
    const reparsed = await parseDocx(bytes)
    const ref = reparsed.blocks.flatMap((b) => b.runs ?? []).find((r) => r.refField === name)
    expect(ref?.text).toBe('Scope')
    expect(reparsed.blocks[0].hiddenBookmarks).toEqual([name])
    editor.destroy()
  })

  it('builds the field for each kind of reference', () => {
    const base = {
      hyperlink: true,
      includeAboveBelow: false,
      pages: { target: 2, here: 5 },
      words: WORDS,
      text: 'Scope',
      before: true,
    }
    expect(crossRefField('B', { ...base, content: 'text', hyperlink: false })).toEqual({
      text: 'Scope',
      mark: { type: 'refField', attrs: { name: 'B', instr: ' REF B ', dirty: false } },
    })
    expect(crossRefField('B', { ...base, content: 'aboveBelow' })).toEqual({
      text: 'above',
      mark: { type: 'refField', attrs: { name: 'B', instr: ' REF B \\p \\h ', dirty: false } },
    })
    expect(crossRefField('B', { ...base, content: 'page' })).toEqual({
      text: '2',
      mark: { type: 'instrField', attrs: { instr: 'PAGEREF B \\h', dirty: true } },
    })
    expect(crossRefField('B', { ...base, content: 'page', includeAboveBelow: true })).toMatchObject(
      { text: 'on page 2', mark: { attrs: { instr: 'PAGEREF B \\p \\h' } } },
    )
    expect(
      crossRefField('B', {
        ...base,
        content: 'page',
        includeAboveBelow: true,
        before: false,
        pages: { target: 5, here: 5 },
      }).text,
    ).toBe('below')
  })

  it('a page reference to a bookmark saves a PAGEREF field Word recomputes', async () => {
    const { editor, parsed } = await open()
    caretAtEnd(editor)
    const target = listRefTargets(editor, parsed.blocks, { kind: 'bookmark' }).find(
      (b) => b.bookmark === 'Alpha',
    )!
    insertCrossReference(editor, parsed.blocks, {
      target,
      content: 'page',
      hyperlink: false,
      includeAboveBelow: false,
      pages: { target: 1, here: 1 },
      words: WORDS,
    })
    // the page is Word's to compute: this field keeps its mark
    allowDirtyFields()
    const { xml } = await saved(editor, parsed)
    expect(xml).toMatch(
      /<w:fldChar w:fldCharType="begin" w:dirty="true"\/>.*?PAGEREF Alpha\s*<\/w:instrText>/,
    )
    editor.destroy()
  })

  it('a caption read from the file gets its bookmark inside its own paragraph, once', async () => {
    const { editor, parsed } = await open()
    caretAtEnd(editor)
    const [caption] = listRefTargets(editor, parsed.blocks, { kind: 'caption', label: 'Figure' })
    expect(caption).toMatchObject({ text: 'Figure 1 Plan', bookmark: null })
    expect(
      insertCrossReference(editor, parsed.blocks, {
        target: caption,
        content: 'text',
        hyperlink: true,
        includeAboveBelow: false,
        pages: null,
        words: WORDS,
      }),
    ).toBe(true)
    const name = listRefTargets(editor, parsed.blocks, { kind: 'caption', label: 'Figure' })[0]
      .bookmark!
    expect(name).toMatch(/^_Ref\d{9}$/)
    const { xml, bytes } = await saved(editor, parsed)
    expect(xml.match(/ SEQ Figure /g)).toHaveLength(1)
    expect(xml).toMatch(
      new RegExp(
        `</w:pPr><w:bookmarkStart w:id="(\\d+)" w:name="${name}"/>.*? SEQ Figure .*?<w:bookmarkEnd w:id="\\1"/></w:p>`,
      ),
    )
    expect(xml).toContain(` REF ${name} \\h `)
    const reparsed = await parseDocx(bytes)
    expect(reparsed.blocks.map((b) => b.docxIndex)).toHaveLength(parsed.blocks.length)
    editor.destroy()
  })
})

describe('captions and references around tables', () => {
  const TABLE =
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid>' +
    `<w:tr><w:tc>${generateCaptionXml('Figure', 9, 'In a cell')}</w:tc></w:tr></w:tbl>`

  it('a table holding a caption in a cell is not a caption itself', async () => {
    const { editor, parsed } = await open(TABLE + BODY)
    expect(collectCaptions(editor, parsed.blocks).map((c) => c.text)).toEqual(['Figure 1 Plan'])
    editor.destroy()
  })

  it('with a block selected the reference goes to the text position after it', async () => {
    const { editor, parsed } = await open()
    let captionPos = -1
    editor.state.doc.forEach((node, pos) => {
      if (node.type.name === 'docProtected') captionPos = pos
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, captionPos)),
    )
    const target = listRefTargets(editor, parsed.blocks, { kind: 'bookmark' }).find(
      (b) => b.bookmark === 'Alpha',
    )!
    expect(
      insertCrossReference(editor, parsed.blocks, {
        target,
        content: 'text',
        hyperlink: false,
        includeAboveBelow: false,
        pages: null,
        words: WORDS,
      }),
    ).toBe(true)
    expect(editor.state.doc.lastChild!.textContent).toBe('Alpha paragraph.Last paragraph.')
    editor.destroy()
  })

  it('the bookmark a first reference gives a caption is listed with the hidden ones', async () => {
    const { editor, parsed } = await open()
    caretAtEnd(editor)
    const [caption] = listRefTargets(editor, parsed.blocks, { kind: 'caption', label: 'Figure' })
    insertCrossReference(editor, parsed.blocks, {
      target: caption,
      content: 'text',
      hyperlink: false,
      includeAboveBelow: false,
      pages: null,
      words: WORDS,
    })
    const hidden = collectBookmarks(editor.state.doc, true).filter((b) => b.hidden)
    expect(hidden.map((b) => b.name)).toEqual(['_Toc9', expect.stringMatching(/^_Ref\d{9}$/)])
    expect(collectBookmarks(editor.state.doc).map((b) => b.name)).toEqual(['Zeta', 'Alpha'])

    // Delete in the Bookmark dialog takes it out of the caption
    removeBookmark(editor, hidden[1])
    expect(collectBookmarks(editor.state.doc, true).map((b) => b.name)).not.toContain(
      hidden[1].name,
    )
    const { xml } = await saved(editor, parsed)
    expect(xml).not.toContain(`w:name="${hidden[1].name}"`)
    expect(xml.match(/<w:bookmarkStart /g)).toHaveLength(xml.match(/<w:bookmarkEnd /g)!.length)
    expect(xml.match(/ SEQ Figure /g)).toHaveLength(1)
    editor.destroy()
  })

  it('a selected table with no text after it is left as it is', async () => {
    const { editor, parsed } = await open(BODY + TABLE)
    let tablePos = -1
    editor.state.doc.forEach((node, pos) => {
      if (node.type.name === 'docTable') tablePos = pos
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, tablePos)),
    )
    const before = editor.state.doc.toJSON()
    const target = listRefTargets(editor, parsed.blocks, { kind: 'bookmark' }).find(
      (b) => b.bookmark === 'Alpha',
    )!
    expect(
      insertCrossReference(editor, parsed.blocks, {
        target,
        content: 'text',
        hyperlink: false,
        includeAboveBelow: false,
        pages: null,
        words: WORDS,
      }),
    ).toBe(false)
    expect(editor.state.doc.toJSON()).toEqual(before)
    editor.destroy()
  })
})

describe('Go To kinds', () => {
  it('finds fields, graphics, equations and objects', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'Page ' },
              {
                type: 'text',
                text: '3',
                marks: [{ type: 'instrField', attrs: { instr: 'PAGE' } }],
              },
              { type: 'text', text: ' see ' },
              { type: 'text', text: 'Scope', marks: [{ type: 'refField', attrs: { name: 'B' } }] },
              { type: 'docInlineMath', attrs: { omml: '<m:oMath/>', text: 'x+y' } },
            ],
          },
          { type: 'docProtected', attrs: { blockType: 'image', label: 'Picture' } },
          { type: 'docProtected', attrs: { blockType: 'image', oleProgId: 'Excel.Sheet.12' } },
          {
            type: 'docProtected',
            attrs: { fieldDisplay: { kind: 'text', left: 'Figure 1' }, previewText: 'Figure 1' },
          },
        ],
      } as never,
    })
    const labels = (kind: Parameters<typeof listGoToItems>[1]) =>
      listGoToItems(editor.state.doc, kind).map((i) => i.label)
    expect(labels('field')).toEqual(['3', 'Scope', 'Figure 1'])
    expect(labels('graphic')).toEqual(['Picture'])
    expect(labels('equation')).toEqual(['x+y'])
    expect(labels('object')).toEqual(['Excel.Sheet.12'])
    editor.destroy()
  })
})
