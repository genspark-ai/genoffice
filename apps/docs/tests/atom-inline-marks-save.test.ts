import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { parseDocx, type Run } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  blocksToPmDoc,
  inlineToRuns,
  pmDocToSavePlan,
  runsToInline,
  type PmNode,
} from '../src/renderer/editor/convert'

/**
 * genoffice#524: inline atoms (note refs, XE marks, formulas, ruby, pictures)
 * used to drop their run marks on the way into the editor, so a footnote
 * reference inside a w:ins never matched its original signature; the paragraph
 * regenerated and lost the tracked change and the character style.
 */

const W16DU = 'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"'
const INS = 'w:author="Author A" w:date="2026-01-01T10:00:00Z" w16du:dateUtc="2026-01-01T10:00:00Z"'
const REV = { author: 'Author A', date: '2026-01-01T10:00:00Z', dateUtc: '2026-01-01T10:00:00Z' }
const FOOTNOTES_PART = {
  path: 'word/footnotes.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:footnote w:id="1"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:footnote></w:footnotes>',
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
}
const FOOTNOTES_REL =
  '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'

describe('marks on atomic inline nodes', () => {
  it('round-trips revision, comment and style marks on every atom kind', () => {
    const runs: Run[] = [
      {
        text: '1',
        noteRef: { kind: 'footnote', id: '1' },
        ins: { ...REV, id: '7' },
        styleId: 'FootnoteReference',
      },
      { text: '', xeTerm: 'term', del: { author: 'B', id: '8' } },
      { text: 'x', math: { omml: '<m:oMath/>' }, commentIds: ['3'] },
      { text: 'kanji', ruby: { rt: 'rt', xml: '<w:ruby/>' }, ins: { author: 'C' } },
      {
        text: '',
        image: { dataUrl: 'data:image/png;base64,AA==', xml: '<w:drawing/>' },
        ins: { author: 'D' },
      },
    ]
    const inline = runsToInline(runs)
    expect(inline.map((n) => n.type)).toEqual([
      'docNoteRef',
      'docXeMark',
      'docInlineMath',
      'docRuby',
      'docInlineImage',
    ])
    expect(inline[0].marks?.map((m) => m.type)).toEqual(['ins', 'docTextStyle'])
    expect(inline[0].marks?.[0].attrs).toMatchObject({ ...REV, id: '7' })
    expect(inlineToRuns(inline)).toEqual(runs)
  })

  it('keeps a footnote reference inside w:ins verbatim on save', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        docRootExtraAttrs: W16DU,
        extraParts: [FOOTNOTES_PART],
        extraRels: FOOTNOTES_REL,
        bodyXml:
          `<w:p><w:r><w:t xml:space="preserve">Body </w:t></w:r><w:ins w:id="7" ${INS}>` +
          '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
          '</w:ins></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>',
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const doc = editor.getJSON() as PmNode
    const atom = doc.content![0].content!.find((n) => n.type === 'docNoteRef')!
    expect(atom.marks?.find((m) => m.type === 'ins')?.attrs).toMatchObject({ ...REV, id: '7' })
    expect(pmDocToSavePlan(doc, parsed.blocks).saveBlocks[0]).toEqual({
      kind: 'original',
      docxIndex: 0,
    })
    editor.destroy()
  })

  it('carries w16du:dateUtc through the ins/del marks into regenerated runs', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        docRootExtraAttrs: W16DU,
        bodyXml:
          `<w:p><w:ins w:id="21" ${INS}><w:r><w:t>Inserted</w:t></w:r></w:ins>` +
          `<w:del w:id="22" w:author="B" w:date="2026-01-01T11:00:00Z" w16du:dateUtc="2026-01-01T11:00:00Z">` +
          '<w:r><w:delText>Gone</w:delText></w:r></w:del></w:p>',
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    editor.commands.insertContentAt(1, 'X')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const block = plan.saveBlocks[0]
    expect(block.kind).toBe('generated')
    if (block.kind !== 'generated') return
    const runs = block.block.runs!
    expect(runs.find((r) => r.text === 'Inserted')?.ins).toEqual({ ...REV, id: '21' })
    expect(runs.find((r) => r.text === 'Gone')?.del).toMatchObject({
      dateUtc: '2026-01-01T11:00:00Z',
      id: '22',
    })
    editor.destroy()
  })
})
