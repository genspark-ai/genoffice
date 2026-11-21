import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { insertCrossReference, listRefTargets } from '../src/renderer/editor/cross-reference'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { anchorPos } from '../src/renderer/editor/link-targets'
import { fieldLinkBookmark, refreshNoteRefFields } from '../src/renderer/editor/note-ref-fields'
import { insertNoteRef, removeNoteRefs } from '../src/renderer/ai/note-ops'

const WORDS = { above: 'above', below: 'below', onPage: (n: number) => `on page ${n}` }

const NOTE = (id: number, text: string) =>
  `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:footnote>`
const FOOTNOTES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  NOTE(1, 'First note text') +
  NOTE(2, 'Second note text') +
  '</w:footnotes>'
const MARK = (id: number) =>
  `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`
const field = (instr: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

const PLAIN_BODY =
  `<w:p><w:r><w:t>Intro</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>Alpha</w:t></w:r>${MARK(1)}</w:p>` +
  `<w:p><w:r><w:t>Beta</w:t></w:r>${MARK(2)}</w:p>` +
  '<w:p><w:r><w:t xml:space="preserve">See note </w:t></w:r></w:p>'

// what Word writes: the bookmark around the mark alone, the field naming it
const WORD_BODY =
  `<w:p><w:r><w:t>Alpha</w:t></w:r>${MARK(1)}</w:p>` +
  '<w:p><w:bookmarkStart w:id="4" w:name="Whole"/><w:r><w:t>Beta</w:t></w:r>' +
  `<w:bookmarkStart w:id="7" w:name="_Ref100200300"/>${MARK(2)}<w:bookmarkEnd w:id="7"/>` +
  '<w:r><w:t xml:space="preserve"> tail</w:t></w:r><w:bookmarkEnd w:id="4"/></w:p>' +
  `<w:p><w:r><w:t xml:space="preserve">See note </w:t></w:r>${field('NOTEREF _Ref100200300 \\h', '2')}</w:p>`

async function open(bodyXml: string) {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml,
      extraRels:
        '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      extraParts: [
        {
          path: 'word/footnotes.xml',
          xml: FOOTNOTES_XML,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        },
      ],
    }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

async function savedXml(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
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

const fields = (editor: Editor): Array<[string, string]> => {
  const out: Array<[string, string]> = []
  editor.state.doc.descendants((node) => {
    const mark = node.marks.find((m) => m.type.name === 'instrField')
    if (node.isText && mark) out.push([String(mark.attrs.instr), node.text ?? ''])
  })
  return out
}

describe('cross-reference to a footnote', () => {
  it('lists the notes by number and the start of their text', async () => {
    const { editor, parsed } = await open(PLAIN_BODY)
    const notes = { footnote: parsed.footnotes, endnote: parsed.endnotes }
    const targets = listRefTargets(editor, parsed.blocks, { kind: 'footnote' }, notes)
    expect(targets.map((t) => [t.text, t.noteMark, t.bookmark])).toEqual([
      ['1 First note text', '1', null],
      ['2 Second note text', '2', null],
    ])
    expect(listRefTargets(editor, parsed.blocks, { kind: 'endnote' }, notes)).toEqual([])
    editor.destroy()
  })

  it('puts a _Ref bookmark on the mark alone and a NOTEREF field at the caret', async () => {
    const { editor, parsed } = await open(PLAIN_BODY)
    const notes = { footnote: parsed.footnotes, endnote: parsed.endnotes }
    const target = listRefTargets(editor, parsed.blocks, { kind: 'footnote' }, notes)[1]
    caretAtEnd(editor)
    const request = { target, hyperlink: true, includeAboveBelow: false, pages: null, words: WORDS }
    expect(insertCrossReference(editor, parsed.blocks, { ...request, content: 'noteNumber' })).toBe(
      true,
    )
    const again = listRefTargets(editor, parsed.blocks, { kind: 'footnote' }, notes)[1]
    expect(again.bookmark).toMatch(/^_Ref\d+$/)
    const name = again.bookmark!
    // a second reference reuses the bookmark
    caretAtEnd(editor)
    insertCrossReference(editor, parsed.blocks, {
      ...request,
      target: again,
      content: 'noteNumberFormatted',
      hyperlink: false,
    })
    caretAtEnd(editor)
    insertCrossReference(editor, parsed.blocks, {
      ...request,
      target: again,
      content: 'aboveBelow',
    })
    expect(fields(editor)).toEqual([
      [`NOTEREF ${name} \\h`, '2'],
      [`NOTEREF ${name} \\f`, '2'],
      [`NOTEREF ${name} \\p \\h`, 'above'],
    ])
    expect(anchorPos(editor.state.doc, name)).toBe(again.pos)
    // only the formatted number is raised, though the next reference went in right behind it
    const raised: boolean[] = []
    editor.state.doc.descendants((node) => {
      if (node.isText && node.marks.some((m) => m.type.name === 'instrField'))
        raised.push(node.marks.some((m) => m.attrs.vertAlign === 'superscript'))
    })
    expect(raised).toEqual([false, true, false])

    const { bytes, xml } = await savedXml(editor, parsed)
    const id = new RegExp(`<w:bookmarkStart w:id="(\\d+)" w:name="${name}"/>`).exec(xml)?.[1]
    expect(id).toBeTruthy()
    // the bookmark wraps the reference run and nothing else
    expect(xml).toContain(
      `<w:bookmarkStart w:id="${id}" w:name="${name}"/>${MARK(2)}<w:bookmarkEnd w:id="${id}"/>`,
    )
    expect(xml.match(new RegExp(`w:name="${name}"`, 'g'))).toHaveLength(1)
    expect(xml).toContain(`<w:instrText xml:space="preserve"> NOTEREF ${name} \\h </w:instrText>`)
    expect(xml).toContain(`<w:instrText xml:space="preserve"> NOTEREF ${name} \\f </w:instrText>`)
    expect(xml).toMatch(
      /NOTEREF _Ref\d+ \\f <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:rPr>(?:(?!<\/w:rPr>).)*<w:vertAlign w:val="superscript"\/>/,
    )

    // read back: the bookmark is on the mark, not on its paragraph, and the fields are fields
    const reparsed = await parseDocx(bytes)
    const runs = reparsed.blocks.flatMap((b) => ('runs' in b ? (b.runs ?? []) : []))
    const host = reparsed.blocks.find(
      (b) => 'runs' in b && (b.runs ?? []).some((r) => r.noteRef?.id === '2'),
    )
    expect(host && 'hiddenBookmarks' in host ? host.hiddenBookmarks : undefined).toBeUndefined()
    expect(runs.find((r) => r.noteRef?.id === '2')?.noteRef?.bookmarks).toEqual([name])
    expect(runs.filter((r) => r.instrField).map((r) => [r.instrField, r.text])).toEqual([
      [`NOTEREF ${name} \\h`, '2'],
      [`NOTEREF ${name} \\f`, '2'],
      [`NOTEREF ${name} \\p \\h`, 'above'],
    ])
    editor.destroy()
  })

  it('follows the number when a note is added or removed before the one referred to', async () => {
    const { editor, parsed } = await open(PLAIN_BODY)
    const notes = { footnote: parsed.footnotes, endnote: parsed.endnotes }
    const target = listRefTargets(editor, parsed.blocks, { kind: 'footnote' }, notes)[1]
    caretAtEnd(editor)
    insertCrossReference(editor, parsed.blocks, {
      target,
      content: 'noteNumber',
      hyperlink: true,
      includeAboveBelow: false,
      pages: null,
      words: WORDS,
    })
    expect(fields(editor).map((f) => f[1])).toEqual(['2'])
    // a new note in the first paragraph
    insertNoteRef(editor, 'footnote', '9', 3)
    expect(fields(editor).map((f) => f[1])).toEqual(['3'])
    removeNoteRefs(editor, 'footnote', '9')
    removeNoteRefs(editor, 'footnote', '1')
    expect(fields(editor).map((f) => f[1])).toEqual(['1'])
    editor.destroy()
  })
})

describe('a reference inserted as a hyperlink', () => {
  it('names the bookmark it goes to, for a note number and for a page number', () => {
    expect(fieldLinkBookmark('NOTEREF _Ref1 \\h')).toBe('_Ref1')
    expect(fieldLinkBookmark('NOTEREF _Ref1 \\f \\h')).toBe('_Ref1')
    expect(fieldLinkBookmark('PAGEREF _Ref1 \\p \\h')).toBe('_Ref1')
    expect(fieldLinkBookmark('PAGEREF "_Ref 1" \\h')).toBe('_Ref 1')
    expect(fieldLinkBookmark('NOTEREF _Ref1 \\f')).toBeNull()
    expect(fieldLinkBookmark('PAGEREF _Ref1')).toBeNull()
    expect(fieldLinkBookmark('PAGE \\h')).toBeNull()
  })
})

describe('a file with a NOTEREF field', () => {
  it('reads the bookmark on the mark and the field, and saves both back untouched', async () => {
    const { editor, parsed } = await open(WORD_BODY)
    const host = parsed.blocks[1]
    expect('bookmarks' in host ? host.bookmarks : undefined).toEqual(['Whole'])
    expect('hiddenBookmarks' in host ? host.hiddenBookmarks : undefined).toBeUndefined()
    expect(fields(editor)).toEqual([['NOTEREF _Ref100200300 \\h', '2']])
    const notes = { footnote: parsed.footnotes, endnote: parsed.endnotes }
    expect(
      listRefTargets(editor, parsed.blocks, { kind: 'footnote' }, notes).map((t) => t.bookmark),
    ).toEqual([null, '_Ref100200300'])
    const untouched = await savedXml(editor, parsed)
    for (const piece of WORD_BODY.split('</w:p>').filter(Boolean))
      expect(untouched.xml).toContain(piece)

    // an edit to the marked paragraph keeps the bookmark on the mark
    editor.commands.insertContentAt(editor.state.doc.child(0).nodeSize + 1, 'New ')
    const edited = await savedXml(editor, parsed)
    expect(edited.xml).toMatch(
      /<w:bookmarkStart w:id="\d+" w:name="_Ref100200300"\/><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="2"\/><\/w:r><w:bookmarkEnd w:id="\d+"\/>/,
    )
    expect(edited.xml.match(/w:name="_Ref100200300"/g)).toHaveLength(1)
    editor.destroy()
  })

  it('updates a stale result, and the above / below one only when asked with the words', async () => {
    const stale = WORD_BODY.replace(
      field('NOTEREF _Ref100200300 \\h', '2'),
      field('NOTEREF _Ref100200300 \\h', '7') + field('NOTEREF _Ref100200300 \\p \\h', 'below'),
    )
    const { editor } = await open(stale)
    const quiet = editor.state.tr
    expect(refreshNoteRefFields(quiet)).toBe(1)
    editor.view.dispatch(quiet)
    expect(fields(editor).map((f) => f[1])).toEqual(['2', 'below'])
    const asked = editor.state.tr
    expect(refreshNoteRefFields(asked, WORDS)).toBe(1)
    editor.view.dispatch(asked)
    expect(fields(editor).map((f) => f[1])).toEqual(['2', 'above'])
    editor.destroy()
  })
})
