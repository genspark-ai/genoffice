import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { Node as ProseNode } from '@tiptap/pm/model'
import JSZip from 'jszip'
import type { Run } from '@genoffice/docx-engine'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  inlineToRuns,
  pmDocToSavePlan,
  runsToInline,
  type PmNode,
} from '../src/renderer/editor/convert'
import {
  buildNotesContext,
  editNoteText,
  noteInsertPos,
  protectedNoteMarkBlock,
} from '../src/renderer/ai/note-ops'

function createEditor(content: PmNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

describe('footnote / endnote reference conversion', () => {
  it('a styled reference keeps its rPr through an edit to its own paragraph', async () => {
    // genoffice#2005: the apply path rebuilds the edited paragraph and the
    // marker used to come out as a bare superscript — style, fonts, colour and
    // size gone. The rPr rides the parse, the atom and the save plan verbatim.
    const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    const body =
      '<w:p><w:r><w:t xml:space="preserve">The court held so. </w:t></w:r>' +
      '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/>' +
      '<w:rFonts w:ascii="Palatino Linotype" w:hAnsi="Palatino Linotype"/>' +
      '<w:color w:val="000000"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr>' +
      '<w:footnoteReference w:id="2"/></w:r></w:p>'
    const footnotes =
      `${XML_DECL}<w:footnotes ${W}>` +
      '<w:footnote w:id="2"><w:p><w:r><w:t> A footnote.</w:t></w:r></w:p></w:footnote>' +
      '</w:footnotes>'
    const source = await buildDocx({
      bodyXml: body,
      extraRels:
        '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      extraParts: [
        {
          path: 'word/footnotes.xml',
          xml: footnotes,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        },
      ],
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    // the op: reword the text run; the marker atom is not touched
    editor.commands.insertContentAt(2, 'X')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const out = await saveDocx(parsed, plan.saveBlocks)
    const xml = await JSZip.loadAsync(out).then((z) =>
      z.file('word/document.xml')!.async('string'),
    )
    const at = xml.indexOf('<w:footnoteReference')
    const refRun = xml.slice(xml.lastIndexOf('<w:r>', at), xml.indexOf('</w:r>', at) + 6)
    expect(refRun).toContain('<w:rStyle w:val="FootnoteReference"/>')
    expect(refRun).toContain('w:ascii="Palatino Linotype"')
    expect(refRun).toContain('<w:color w:val="000000"/>')
    expect(refRun).toContain('<w:sz w:val="26"/>')
  })

  it('maps Run.noteRef to a docNoteRef atom and back', () => {
    const runs: Run[] = [
      { text: 'body' },
      { text: '2', noteRef: { kind: 'footnote', id: '3' } },
      { text: 'continued' },
    ]
    const inline = runsToInline(runs)
    expect(inline).toEqual([
      { type: 'text', text: 'body' },
      { type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 2, rawRPr: null } },
      { type: 'text', text: 'continued' },
    ])
    expect(inlineToRuns(inline)).toEqual(runs)
  })

  it('maps Run.xeTerm to a trailing docXeMark atom and back', () => {
    const runs: Run[] = [{ text: 'artificial intelligence', xeTerm: 'artificial intelligence' }]
    const inline = runsToInline(runs)
    expect(inline).toEqual([
      { type: 'text', text: 'artificial intelligence' },
      { type: 'docXeMark', attrs: { term: 'artificial intelligence' } },
    ])
    // splits back into a plain text run + an empty XE-carrier run
    expect(inlineToRuns(inline)).toEqual([
      { text: 'artificial intelligence' },
      { text: '', xeTerm: 'artificial intelligence' },
    ])
  })

  it('renders both atoms in the editor schema without dropping them', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: 0 },
        content: [
          { type: 'text', text: 'a' },
          { type: 'docNoteRef', attrs: { kind: 'endnote', id: '5', num: 1 } },
          { type: 'docXeMark', attrs: { term: 'glossary term' } },
        ],
      },
    ])
    const para = (editor.getJSON() as PmNode).content![0]
    const types = (para.content ?? []).map((n) => n.type)
    expect(types).toEqual(['text', 'docNoteRef', 'docXeMark'])
    const html = editor.getHTML()
    expect(html).toContain('data-note-ref="5"')
    expect(html).toContain('data-xe-term="glossary term"')
    editor.destroy()
  })
})

describe('noteInsertPos', () => {
  it('skips afterText matches that end inside a field result', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: 0 },
        content: [
          { type: 'text', text: 'See ' },
          { type: 'text', text: 'Table 12', marks: [{ type: 'refField', attrs: { name: 'tbl' } }] },
          { type: 'text', text: ' below' },
        ],
      },
    ])
    const doc = editor.state.doc
    expect(noteInsertPos(doc, 0, 'Table 1')).toEqual({
      error: expect.stringContaining('ends inside the field { REF tbl }'),
    })
    expect(noteInsertPos(doc, 0, 'Table 12')).toEqual({ pos: 1 + 'See Table 12'.length })
    expect(noteInsertPos(doc, 0, 'See')).toEqual({ pos: 1 + 'See'.length })
    editor.destroy()
  })
})

describe('buildNotesContext', () => {
  it('locates marks that live only in protected block XML', () => {
    const editor = createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: 0 },
        content: [
          { type: 'text', text: 'a' },
          { type: 'docNoteRef', attrs: { kind: 'footnote', id: '1', num: 1 } },
        ],
      },
      {
        type: 'docProtected',
        attrs: {
          docxIndex: 1,
          blockType: 'passthrough',
          label: 'p',
          previewText: '',
          genXml: '<w:p><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
        },
      },
    ])
    const doc = editor.state.doc
    const footnotes = [
      { id: '1', text: 'first' },
      { id: '2', text: 'second' },
      { id: '3', text: 'orphan' },
    ]
    const xmlOf = (block: ProseNode) => String(block.attrs.genXml ?? '')
    const out = buildNotesContext(doc, footnotes, [], (kind, id) =>
      protectedNoteMarkBlock(doc, xmlOf, kind, id),
    )
    expect(out).toContain('id 1 (mark 1 in block 0)')
    expect(out).toContain('id 2 (mark in protected block 1, cannot be deleted here)')
    expect(out).toContain('id 3 (no reference mark in the text)')
  })
})

describe('editNoteText', () => {
  it('replaces every occurrence in the plain text and reports the count', () => {
    const { note, count } = editNoteText(
      { id: '3', text: 'See Smith 2019; cf. Smith 2019.' },
      'Smith 2019',
      'Smith 2020',
    )
    expect(count).toBe(2)
    expect(note).toEqual({ id: '3', text: 'See Smith 2020; cf. Smith 2020.' })
  })

  it('returns the same note untouched when nothing matches', () => {
    const original = { id: '1', text: 'Docket 12-345.' }
    const { note, count } = editNoteText(original, '99-000', '12-346')
    expect(count).toBe(0)
    expect(note).toBe(original)
  })

  it('matches case-insensitively only when asked', () => {
    const original = { id: '1', text: 'ibid. Ibid.' }
    expect(editNoteText(original, 'ibid.', 'id.').note.text).toBe('id. Ibid.')
    expect(editNoteText(original, 'ibid.', 'id.', false).note.text).toBe('id. id.')
  })

  it('patches display runs in place, keeping the formatting of the run the match starts in', () => {
    const { note } = editNoteText(
      {
        id: '2',
        text: 'Brown v. Board, 347 U.S. 483 (1954).',
        richParas: [
          [
            { text: 'Brown v. Board', italic: true },
            { text: ', 347 U.S. ' },
            { text: '483', bold: true },
            { text: ' (1954).' },
          ],
        ],
      },
      'U.S. 483',
      'U.S. 484',
    )
    expect(note.text).toBe('Brown v. Board, 347 U.S. 484 (1954).')
    expect(note.richParas).toEqual([
      [
        { text: 'Brown v. Board', italic: true },
        { text: ', 347 U.S. 484' },
        { text: '', bold: true },
        { text: ' (1954).' },
      ],
    ])
  })

  it('drops display runs when the replacement introduces a paragraph break', () => {
    const { note } = editNoteText(
      { id: '4', text: 'one two', richParas: [[{ text: 'one two', italic: true }]] },
      ' ',
      '\n',
    )
    expect(note.text).toBe('one\ntwo')
    expect(note.richParas).toBeUndefined()
  })

  it('drops display runs that disagree with the text instead of showing stale formatting', () => {
    const { note } = editNoteText(
      { id: '2', text: 'alpha beta', richParas: [[{ text: 'gamma', bold: true }]] },
      'beta',
      'delta',
    )
    expect(note.text).toBe('alpha delta')
    expect(note.richParas).toBeUndefined()
  })
})
