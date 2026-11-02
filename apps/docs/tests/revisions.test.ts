import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import type { Fragment, Node as ProseNode } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'
import { addRowAfter, deleteRow } from '@tiptap/pm/tables'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  tableModelToPmNode,
  type PmNode,
} from '../src/renderer/editor/convert'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
  REVISION_TABLE_XML,
  TABLE_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'

import { deleteTableColumn, insertRowsOrColumns } from '../src/renderer/editor/table-ops'
import { replaceInlineRange } from '../src/renderer/ai/protocol'
import { setAuthorName } from '../src/renderer/author-name'
import {
  acceptAllRevisions,
  acceptCurrentRevision,
  applyRevisions,
  applyRevisionsBy,
  collectRevisions,
  revisionCount,
  gotoRevision,
  rejectAllRevisions,
  rejectCurrentRevision,
  type TrackChangesStorage,
} from '../src/renderer/editor/revisions'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
}

const text = (t: string, marks?: JsonNode['marks']): JsonNode => ({
  type: 'text',
  text: t,
  ...(marks ? { marks } : {}),
})
const ins = (author = 'Alice'): NonNullable<JsonNode['marks']>[0] => ({
  type: 'ins',
  attrs: { author, date: '2026-07-01T10:00:00Z' },
})
const del = (author = 'Bob'): NonNullable<JsonNode['marks']>[0] => ({
  type: 'del',
  attrs: { author, date: '2026-07-02T10:00:00Z' },
})
const para = (...content: JsonNode[]): JsonNode => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content,
})

function createEditor(content: JsonNode[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

const revisedDoc = () => [para(text('KE '), text('IN', [ins()]), text('DE', [del()]), text(' PT'))]

describe('collectRevisions', () => {
  it('finds contiguous revision ranges with kind and author', () => {
    const editor = createEditor(revisedDoc())
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(2)
    expect(ranges[0]).toMatchObject({ kind: 'ins', author: 'Alice' })
    expect(ranges[1]).toMatchObject({ kind: 'del', author: 'Bob' })
    editor.destroy()
  })

  it('reports ins+del on the same text as two revisions, deletion first', () => {
    const editor = createEditor([para(text('ins-then-del', [ins(), del()]))])
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(2)
    expect(ranges[0]).toMatchObject({ kind: 'del', author: 'Bob', overlapped: true })
    expect(ranges[1]).toMatchObject({ kind: 'ins', author: 'Alice', overlapped: true })
    expect(ranges[0].from).toBe(ranges[1].from)
    expect(ranges[0].to).toBe(ranges[1].to)
    editor.destroy()
  })
})

describe('accept / reject all', () => {
  it('accept keeps insertions as plain text and removes deletions', () => {
    const editor = createEditor(revisedDoc())
    acceptAllRevisions(editor)
    expect(editor.state.doc.textContent).toBe('KE IN PT')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })

  it('reject removes insertions and restores deletions as plain text', () => {
    const editor = createEditor(revisedDoc())
    rejectAllRevisions(editor)
    expect(editor.state.doc.textContent).toBe('KE DE PT')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })

  it('removes insert-then-delete text on both accept and reject', () => {
    for (const apply of [acceptAllRevisions, rejectAllRevisions]) {
      const editor = createEditor([para(text('a'), text('tmp', [ins(), del()]), text('b'))])
      apply(editor)
      expect(editor.state.doc.textContent).toBe('ab')
      editor.destroy()
    }
  })

  it('does not merge overlapped text into a neighbouring plain insertion', () => {
    const editor = createEditor([para(text('kept', [ins()]), text('gone', [ins(), del()]))])
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.map((r) => [r.kind, !!r.overlapped])).toEqual([
      ['ins', false],
      ['del', true],
      ['ins', true],
    ])
    editor.destroy()
  })

  it("rejecting only the deletion of another author's insertion restores the text", () => {
    const editor = createEditor([para(text('a'), text('tmp', [ins(), del()]), text('b'))])
    applyRevisionsBy(editor, 'Bob', 'reject')
    expect(editor.state.doc.textContent).toBe('atmpb')
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(1)
    expect(ranges[0]).toMatchObject({ kind: 'ins', author: 'Alice' })
    editor.destroy()
  })

  it('accepting only the insertion keeps the text as a pending deletion', () => {
    const editor = createEditor([para(text('a'), text('tmp', [ins(), del()]), text('b'))])
    applyRevisionsBy(editor, 'Alice', 'accept')
    expect(editor.state.doc.textContent).toBe('atmpb')
    expect(collectRevisions(editor.state.doc)).toMatchObject([{ kind: 'del', author: 'Bob' }])
    editor.destroy()
  })

  it('accept all does not strip marks from text after a removed deletion', () => {
    const editor = createEditor([
      para(text('a'), text('tmp', [ins(), del()]), text('keep', [del()]), text('b')),
    ])
    rejectAllRevisions(editor)
    expect(editor.state.doc.textContent).toBe('akeepb')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })

  it('resolving a whole-paragraph block revision also clears its run marks', () => {
    // a redline on disk marks the runs AND wraps the block; the editor holds
    // both, and resolving the block must not leave the run marks pending
    const blockIns = { kind: 'ins', author: 'Alice', date: '2026-07-01T10:00:00Z' }
    const blockDel = { kind: 'del', author: 'Bob', date: '2026-07-02T10:00:00Z' }
    const redline = () => [
      para(text('kept')),
      { ...para(text('added', [ins()])), attrs: { docxIndex: null, blockRevision: blockIns } },
      { ...para(text('gone', [del()])), attrs: { docxIndex: null, blockRevision: blockDel } },
    ]
    const accepted = createEditor(redline())
    acceptAllRevisions(accepted)
    expect(accepted.state.doc.textContent).toBe('keptadded')
    expect(collectRevisions(accepted.state.doc)).toHaveLength(0)
    accepted.destroy()

    const rejected = createEditor(redline())
    rejectAllRevisions(rejected)
    expect(rejected.state.doc.textContent).toBe('keptgone')
    expect(collectRevisions(rejected.state.doc)).toHaveLength(0)
    rejected.destroy()
  })
})

describe('accept / reject current + navigation', () => {
  it('accepts only the revision at the cursor', () => {
    const editor = createEditor(revisedDoc())
    // cursor inside the ins range (the leading kept text is 3 chars; +1 doc offset)
    const pos = 5
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos)))
    expect(acceptCurrentRevision(editor)).toBe(true)
    expect(editor.state.doc.textContent).toBe('KE INDE PT')
    expect(collectRevisions(editor.state.doc)).toHaveLength(1)
    expect(collectRevisions(editor.state.doc)[0].kind).toBe('del')
    editor.destroy()
  })

  it('falls back to the next revision when the cursor is outside any', () => {
    const editor = createEditor(revisedDoc())
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)))
    expect(rejectCurrentRevision(editor)).toBe(true)
    // the ins range was rejected (removed)
    expect(editor.state.doc.textContent).toBe('KE DE PT')
    editor.destroy()
  })

  it('gotoRevision selects the next range and wraps', () => {
    const editor = createEditor(revisedDoc())
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1)))
    expect(gotoRevision(editor, 1)?.kind).toBe('ins')
    const first = editor.state.selection
    expect(editor.state.doc.textBetween(first.from, first.to)).toBe('IN')
    expect(gotoRevision(editor, 1)?.kind).toBe('del')
    expect(
      editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to),
    ).toBe('DE')
    editor.destroy()
  })

  it('returns nothing with no revisions', () => {
    const editor = createEditor([para(text('plain'))])
    expect(acceptCurrentRevision(editor)).toBe(false)
    expect(gotoRevision(editor, 1)).toBeNull()
    editor.destroy()
  })
})

describe('track changes recording', () => {
  function trackingEditor(content: JsonNode[]): Editor {
    const editor = createEditor(content)
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Tester'
    return editor
  }

  it('marks typed text as an insertion', () => {
    const editor = trackingEditor([para(text('hi'))])
    editor.commands.insertContentAt(3, 'yo')
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(1)
    expect(ranges[0]).toMatchObject({ kind: 'ins', author: 'Tester' })
    expect(editor.state.doc.textBetween(ranges[0].from, ranges[0].to)).toBe('yo')
    editor.destroy()
  })

  it('re-materializes deleted text struck through', () => {
    const editor = trackingEditor([para(text('ABCD'))])
    editor.commands.deleteRange({ from: 2, to: 4 }) // delete the 2nd and 3rd characters
    expect(editor.state.doc.textContent).toBe('ABCD')
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(1)
    expect(ranges[0].kind).toBe('del')
    expect(editor.state.doc.textBetween(ranges[0].from, ranges[0].to)).toBe('BC')
    editor.destroy()
  })

  it('deleting your own tracked insertion removes it outright', () => {
    const editor = trackingEditor([para(text('hi'))])
    editor.commands.insertContentAt(3, 'XY')
    const [insRange] = collectRevisions(editor.state.doc)
    editor.commands.deleteRange({ from: insRange.from, to: insRange.to })
    expect(editor.state.doc.textContent).toBe('hi')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })

  it('replacing a selection tracks both the insertion and the deletion', () => {
    const editor = trackingEditor([para(text('old text'))])
    editor.commands.insertContentAt({ from: 1, to: 3 }, 'new ')
    const ranges = collectRevisions(editor.state.doc)
    const kinds = ranges.map((r) => r.kind).sort()
    expect(kinds).toEqual(['del', 'ins'])
    // document still shows both old (struck) and new text
    expect(editor.state.doc.textContent).toBe('new old text')
    editor.destroy()
  })

  it('tracks hyperlink text insertion/deletion and find-replace transactions', () => {
    const linkEditor = trackingEditor([para(text('base '))])
    linkEditor.commands.insertContentAt(6, {
      type: 'text',
      text: 'site',
      marks: [{ type: 'link', attrs: { href: 'https://example.com', rId: null } }],
    })
    expect(collectRevisions(linkEditor.state.doc).map((revision) => revision.kind)).toContain('ins')
    const linkNode = linkEditor.state.doc.nodeAt(6)
    expect(linkNode?.marks.some((mark) => mark.type.name === 'link')).toBe(true)

    const replaceEditor = trackingEditor([para(text('find this value'))])
    replaceEditor.commands.command(({ tr }) => {
      tr.insertText('that', 6, 10)
      return true
    })
    expect(
      collectRevisions(replaceEditor.state.doc)
        .map((revision) => revision.kind)
        .sort(),
    ).toEqual(['del', 'ins'])
    acceptAllRevisions(replaceEditor)
    expect(replaceEditor.state.doc.textContent).toBe('find that value')
    linkEditor.destroy()
    replaceEditor.destroy()
  })

  it('records nothing while disabled', () => {
    const editor = createEditor([para(text('ab'))])
    editor.commands.insertContentAt(3, 'cd')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })

  it('records character mark changes and accept/reject preserve the correct format', () => {
    const acceptEditor = trackingEditor([para(text('fm'))])
    acceptEditor.chain().setTextSelection({ from: 1, to: 3 }).toggleMark('bold').run()
    expect(collectRevisions(acceptEditor.state.doc)).toMatchObject([
      { kind: 'rPrChange', author: 'Tester' },
    ])
    acceptAllRevisions(acceptEditor)
    expect(collectRevisions(acceptEditor.state.doc)).toHaveLength(0)
    expect(acceptEditor.state.doc.nodeAt(1)?.marks.some((m) => m.type.name === 'bold')).toBe(true)

    const rejectEditor = trackingEditor([para(text('fm'))])
    rejectEditor.chain().setTextSelection({ from: 1, to: 3 }).toggleMark('bold').run()
    rejectAllRevisions(rejectEditor)
    expect(collectRevisions(rejectEditor.state.doc)).toHaveLength(0)
    expect(rejectEditor.state.doc.nodeAt(1)?.marks.some((m) => m.type.name === 'bold')).toBe(false)
    acceptEditor.destroy()
    rejectEditor.destroy()
  })

  it('restores previous color and highlight when a format revision is rejected', () => {
    const editor = trackingEditor([
      para(text('cl', [{ type: 'docTextStyle', attrs: { color: '0000FF', highlight: 'yellow' } }])),
    ])
    editor
      .chain()
      .setTextSelection({ from: 1, to: 3 })
      .setMark('docTextStyle', { color: 'FF0000', highlight: 'green' })
      .run()
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toContain('rPrChange')
    rejectAllRevisions(editor)
    const textStyle = editor.state.doc.nodeAt(1)?.marks.find((m) => m.type.name === 'docTextStyle')
    expect(textStyle?.attrs.color).toBe('0000FF')
    expect(textStyle?.attrs.highlight).toBe('yellow')
    editor.destroy()
  })

  it('tracks paragraph alignment, indentation, and spacing with accept/reject', () => {
    const acceptEditor = trackingEditor([
      {
        ...para(text('para')),
        attrs: {
          docxIndex: null,
          align: 'left',
          indentLeft: 120,
          lineSpacing: 1,
          spaceAfter: 80,
        },
      },
    ])
    acceptEditor.commands.updateAttributes('docParagraph', {
      align: 'center',
      indentLeft: 360,
      lineSpacing: 1.5,
      spaceAfter: 240,
    })
    expect(collectRevisions(acceptEditor.state.doc)).toMatchObject([
      { kind: 'pPrChange', author: 'Tester' },
    ])
    acceptAllRevisions(acceptEditor)
    expect(acceptEditor.state.doc.firstChild?.attrs).toMatchObject({
      align: 'center',
      indentLeft: 360,
      lineSpacing: 1.5,
      spaceAfter: 240,
      pPrChange: null,
    })

    const rejectEditor = trackingEditor([
      {
        ...para(text('para')),
        attrs: {
          docxIndex: null,
          align: 'left',
          indentLeft: 120,
          lineSpacing: 1,
          spaceAfter: 80,
        },
      },
    ])
    rejectEditor.commands.updateAttributes('docParagraph', {
      align: 'right',
      indentLeft: 480,
      lineSpacing: 2,
      spaceAfter: 300,
    })
    rejectAllRevisions(rejectEditor)
    expect(rejectEditor.state.doc.firstChild?.attrs).toMatchObject({
      align: 'left',
      indentLeft: 120,
      lineSpacing: 1,
      spaceAfter: 80,
      pPrChange: null,
    })
    acceptEditor.destroy()
    rejectEditor.destroy()
  })

  it('tracks paragraph style/type and list level changes', () => {
    const headingEditor = trackingEditor([para(text('title'))])
    const paragraph = headingEditor.state.doc.firstChild!
    headingEditor.view.dispatch(
      headingEditor.state.tr.setNodeMarkup(0, headingEditor.schema.nodes.docHeading, {
        ...paragraph.attrs,
        styleId: 'Heading1',
        level: 1,
      }),
    )
    expect(collectRevisions(headingEditor.state.doc).map((revision) => revision.kind)).toEqual([
      'pPrChange',
    ])
    rejectAllRevisions(headingEditor)
    expect(headingEditor.state.doc.firstChild?.type.name).toBe('docParagraph')
    expect(headingEditor.state.doc.firstChild?.attrs.styleId).toBeNull()

    const listEditor = trackingEditor([
      {
        type: 'docListItem',
        attrs: { docxIndex: null, kind: 'bullet', numId: '1', ilvl: 0 },
        content: [text('item')],
      },
    ])
    listEditor.commands.updateAttributes('docListItem', { ilvl: 2 })
    expect(collectRevisions(listEditor.state.doc).map((revision) => revision.kind)).toEqual([
      'pPrChange',
    ])
    rejectAllRevisions(listEditor)
    expect(listEditor.state.doc.firstChild?.attrs.ilvl).toBe(0)
    headingEditor.destroy()
    listEditor.destroy()
  })

  // anchored paragraphs: docxIndex is cloned onto both halves of a split, so
  // the structural differ must group anchors instead of assuming uniqueness
  const anchored = (idx: number, ...content: JsonNode[]): JsonNode => ({
    type: 'docParagraph',
    attrs: { docxIndex: idx },
    content,
  })

  it('Enter in an anchored paragraph splits it without duplicating the block', () => {
    const editor = trackingEditor([anchored(0, text('alpha beta gamma'))])
    editor.chain().setTextSelection(7).splitBlock().run()
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.textContent).toBe('alpha beta gamma')
    expect(editor.state.doc.child(0).textContent).toBe('alpha ')
    expect(editor.state.doc.child(1).textContent).toBe('beta gamma')
    const blocks = collectRevisions(editor.state.doc).filter(
      (r) => r.kind === 'blockIns' || r.kind === 'blockDel',
    )
    expect(blocks).toHaveLength(0)
    // rejecting whatever was recorded must not lose or duplicate text
    rejectAllRevisions(editor)
    expect(editor.state.doc.textContent).toBe('alpha beta gamma')
    editor.destroy()
  })

  it('joining two halves that share an anchor records a deleted paragraph mark, no phantom block', () => {
    const editor = trackingEditor([anchored(0, text('first ')), anchored(0, text('second'))])
    editor.chain().setTextSelection(9).joinBackward().run()
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.textContent).toBe('first second')
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['paraMarkDel'])
    rejectAllRevisions(editor)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.textContent).toBe('first second')
    editor.destroy()
  })

  it('joining paragraphs with distinct anchors records only the deleted mark; accept joins them', () => {
    const editor = trackingEditor([anchored(0, text('first ')), anchored(1, text('second'))])
    editor.chain().setTextSelection(9).joinBackward().run()
    expect(editor.state.doc.childCount).toBe(2)
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['paraMarkDel'])
    acceptAllRevisions(editor)
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.textContent).toBe('first second')
    editor.destroy()
  })

  it('deleting a paragraph whose text recurs in its neighbour is still a block deletion', () => {
    const editor = trackingEditor([anchored(0, text('a')), anchored(1, text('banana'))])
    editor.commands.deleteRange({ from: 0, to: 3 })
    const blocks = collectRevisions(editor.state.doc).filter((r) => r.kind === 'blockDel')
    expect(blocks).toHaveLength(1)
    rejectAllRevisions(editor)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.child(0).textContent).toBe('a')
    expect(editor.state.doc.child(1).textContent).toBe('banana')
    editor.destroy()
  })

  it('joining an image-only paragraph records no phantom deleted block', () => {
    const editor = trackingEditor([
      anchored(0, text('first')),
      anchored(1, {
        type: 'docInlineImage',
        attrs: { dataUrl: 'data:image/png;base64,AA==', widthPx: 10, heightPx: 10 },
      }),
    ])
    editor.chain().setTextSelection(8).joinBackward().run()
    expect(editor.state.doc.childCount).toBe(2)
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['paraMarkDel'])
    editor.destroy()
  })

  it('still re-materializes a genuinely deleted anchored paragraph', () => {
    const editor = trackingEditor([anchored(0, text('keep')), anchored(1, text('drop'))])
    editor.commands.deleteRange({ from: 6, to: 12 })
    // the deleted paragraph comes back as a struck-through block
    expect(editor.state.doc.textContent).toBe('keepdrop')
    const blocks = collectRevisions(editor.state.doc).filter((r) => r.kind === 'blockDel')
    expect(blocks).toHaveLength(1)
    acceptAllRevisions(editor)
    expect(editor.state.doc.textContent).toBe('keep')
    editor.destroy()
  })

  it('tracks a whole inserted paragraph and rejects it as one block', () => {
    const editor = trackingEditor([para(text('original para'))])
    editor.commands.insertContentAt(editor.state.doc.content.size, para(text('new para')))
    const insertion = collectRevisions(editor.state.doc).find(
      (revision) => revision.kind === 'blockIns',
    )
    expect(insertion).toBeDefined()
    rejectAllRevisions(editor)
    expect(editor.state.doc.childCount).toBe(1)
    expect(editor.state.doc.textContent).toBe('original para')
    editor.destroy()
  })

  it('accept-all after tracked editing yields the edited plain document', () => {
    const editor = trackingEditor([para(text('ABCD'))])
    editor.commands.deleteRange({ from: 2, to: 4 })
    editor.commands.insertContentAt(1, 'Z')
    acceptAllRevisions(editor)
    expect(editor.state.doc.textContent).toBe('ZAD')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })

  it('reject-all after tracked editing restores the original document', () => {
    const editor = trackingEditor([para(text('ABCD'))])
    editor.commands.deleteRange({ from: 2, to: 4 })
    editor.commands.insertContentAt(1, 'Z')
    rejectAllRevisions(editor)
    expect(editor.state.doc.textContent).toBe('ABCD')
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })
})

describe('rPrChange run-level format revisions', () => {
  const RPRCHANGE_PARA =
    '<w:p><w:r><w:rPr><w:b/><w:color w:val="FF0000"/>' +
    '<w:rPrChange w:id="9" w:author="Reviewer" w:date="2026-07-01T00:00:00Z">' +
    '<w:rPr><w:i/><w:color w:val="0000FF"/></w:rPr></w:rPrChange>' +
    '</w:rPr><w:t>restyled</w:t></w:r><w:r><w:t>plain</w:t></w:r></w:p>'

  it('parses as an editable paragraph; run carries rPrChange with the old snapshot; unedited round-trip is byte-identical', async () => {
    const bytes = await buildDocx({ bodyXml: RPRCHANGE_PARA })
    const parsed = await parseDocx(bytes)
    expect(parsed.blocks[0].type).toBe('paragraph')
    const run = parsed.blocks[0].runs![0]
    expect(run.bold).toBe(true)
    expect(run.rPrChange).toMatchObject({
      author: 'Reviewer',
      id: '9',
      old: { italic: true, color: '0000FF' },
    })
    // Unedited save is byte-identical (the record is preserved verbatim in rawRPr)
    const blocks = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    expect(await saveDocx(parsed, blocks)).toBe(bytes)
  })

  it('editor: collectRevisions receives rPrChange; accept clears the record and keeps the new format; reject restores the old format', async () => {
    const bytes = await buildDocx({ bodyXml: RPRCHANGE_PARA })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const revs = collectRevisions(editor.state.doc)
    const rpr = revs.find((r) => r.kind === 'rPrChange')
    expect(rpr).toBeDefined()
    expect(rpr!.author).toBe('Reviewer')

    // Accept: mark cleared, w:rPrChange stripped from rawRPr, bold kept
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, rpr!.from + 1)),
    )
    acceptCurrentRevision(editor)
    expect(collectRevisions(editor.state.doc).some((r) => r.kind === 'rPrChange')).toBe(false)
    let boldKept = false
    let rawHasRecord = false
    editor.state.doc.descendants((node) => {
      if (!node.isText || node.text !== 'restyled') return
      boldKept = node.marks.some((m) => m.type.name === 'bold')
      const ts = node.marks.find((m) => m.type.name === 'docTextStyle')
      rawHasRecord = /<w:rPrChange/.test(String(ts?.attrs.rawRPr ?? ''))
    })
    expect(boldKept).toBe(true)
    expect(rawHasRecord).toBe(false)

    // Reject path (reload): restore italic/blue, drop bold
    const editor2 = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc((await parseDocx(bytes)).blocks),
    })
    const rpr2 = collectRevisions(editor2.state.doc).find((r) => r.kind === 'rPrChange')!
    editor2.view.dispatch(
      editor2.state.tr.setSelection(TextSelection.create(editor2.state.doc, rpr2.from + 1)),
    )
    rejectCurrentRevision(editor2)
    let italicRestored = false
    let boldRemoved = true
    let colorRestored: string | null = null
    editor2.state.doc.descendants((node) => {
      if (!node.isText || node.text !== 'restyled') return
      italicRestored = node.marks.some((m) => m.type.name === 'italic')
      boldRemoved = !node.marks.some((m) => m.type.name === 'bold')
      const ts = node.marks.find((m) => m.type.name === 'docTextStyle')
      colorRestored = (ts?.attrs.color as string | null) ?? null
    })
    expect(italicRestored).toBe(true)
    expect(boldRemoved).toBe(true)
    expect(colorRestored).toBe('0000FF')
    editor.destroy()
    editor2.destroy()
  })

  it('live format edits are saved as w:rPrChange preserving the old-format snapshot', async () => {
    const bytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>fm</w:t></w:r></w:p>' })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Tester'
    editor.chain().setTextSelection({ from: 1, to: 3 }).toggleMark('bold').run()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const run = reparsed.blocks[0].runs![0]
    expect(run.bold).toBe(true)
    expect(run.rPrChange).toMatchObject({ author: 'Tester' })
    expect(run.rPrChange?.old?.bold).not.toBe(true)
    editor.destroy()
  })
})

describe('live pPrChange paragraph formatting', () => {
  it('writes alignment/indent/spacing revision and reopens with the old snapshot', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/>' +
        '<w:ind w:left="120"/><w:jc w:val="left"/></w:pPr><w:r><w:t>para</w:t></w:r></w:p>',
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Paragraph Reviewer'
    editor.commands.updateAttributes('docParagraph', {
      align: 'center',
      indentLeft: 360,
      lineSpacing: 1.5,
      spaceAfter: 240,
    })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].format).toMatchObject({
      align: 'center',
      indentLeft: 360,
      lineSpacing: 1.5,
      spaceAfter: 240,
    })
    expect(reparsed.blocks[0].pPrChangeInfo).toMatchObject({
      author: 'Paragraph Reviewer',
      old: { align: 'left', indentLeft: 120, lineSpacing: 1, spaceAfter: 80 },
    })

    const reopened = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(reparsed.blocks),
    })
    rejectAllRevisions(reopened)
    expect(reopened.state.doc.firstChild?.attrs).toMatchObject({
      align: 'left',
      indentLeft: 120,
      lineSpacing: 1,
      spaceAfter: 80,
    })
    const rejectedPlan = pmDocToSavePlan(reopened.getJSON() as PmNode, reparsed.blocks)
    const rejected = await parseDocx(await saveDocx(reparsed, rejectedPlan.saveBlocks))
    expect(rejected.blocks[0].pPrChangeInfo).toBeUndefined()
    expect(rejected.blocks[0].format).toMatchObject({
      align: 'left',
      indentLeft: 120,
      lineSpacing: 1,
      spaceAfter: 80,
    })
    editor.destroy()
    reopened.destroy()
  })

  it('reopens and rejects paragraph style and list-level revisions', async () => {
    const headingBytes = await buildDocx({
      bodyXml: '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>',
    })
    const headingParsed = await parseDocx(headingBytes)
    const headingEditor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(headingParsed.blocks),
    })
    const headingStorage = headingEditor.storage.trackChanges as TrackChangesStorage
    headingStorage.enabled = true
    headingStorage.author = 'Paragraph Reviewer'
    headingEditor.view.dispatch(
      headingEditor.state.tr.setNodeMarkup(0, headingEditor.schema.nodes.docHeading, {
        ...headingEditor.state.doc.firstChild!.attrs,
        styleId: 'Heading1',
        level: 1,
      }),
    )
    const headingPlan = pmDocToSavePlan(headingEditor.getJSON() as PmNode, headingParsed.blocks)
    const headingReparsed = await parseDocx(await saveDocx(headingParsed, headingPlan.saveBlocks))
    expect(headingReparsed.blocks[0].type).toBe('heading')
    expect(headingReparsed.blocks[0].pPrChangeInfo?.old).toMatchObject({
      type: 'docParagraph',
      styleId: 'Normal',
    })
    const reopenedHeading = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(headingReparsed.blocks),
    })
    rejectAllRevisions(reopenedHeading)
    expect(reopenedHeading.state.doc.firstChild?.type.name).toBe('docParagraph')
    expect(reopenedHeading.state.doc.firstChild?.attrs.styleId).toBe('Normal')

    const listBytes = await buildDocx({
      withNumbering: true,
      bodyXml:
        '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/>' +
        '<w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>',
    })
    const listParsed = await parseDocx(listBytes)
    const listEditor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(listParsed.blocks),
    })
    const listStorage = listEditor.storage.trackChanges as TrackChangesStorage
    listStorage.enabled = true
    listStorage.author = 'Paragraph Reviewer'
    listEditor.commands.updateAttributes('docListItem', { ilvl: 2 })
    const listPlan = pmDocToSavePlan(listEditor.getJSON() as PmNode, listParsed.blocks)
    const listReparsed = await parseDocx(await saveDocx(listParsed, listPlan.saveBlocks))
    expect(listReparsed.blocks[0].list?.ilvl).toBe(2)
    expect(listReparsed.blocks[0].pPrChangeInfo?.old).toMatchObject({
      type: 'docListItem',
      numId: '1',
      ilvl: 0,
    })
    const reopenedList = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(listReparsed.blocks),
    })
    rejectAllRevisions(reopenedList)
    expect(reopenedList.state.doc.firstChild?.attrs.ilvl).toBe(0)
    headingEditor.destroy()
    reopenedHeading.destroy()
    listEditor.destroy()
    reopenedList.destroy()
  })
})

describe('live structural revisions', () => {
  async function trackedDoc(bodyXml: string, withImage = false) {
    const parsed = await parseDocx(await buildDocx({ bodyXml, withImage }))
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Structural Reviewer'
    return { editor, parsed }
  }

  it('re-materializes a deleted whole paragraph and supports accept/reject', async () => {
    const body =
      '<w:p><w:r><w:t>first</w:t></w:r></w:p>' + '<w:p><w:r><w:t>second</w:t></w:r></w:p>'
    const accepted = await trackedDoc(body)
    accepted.editor.view.dispatch(
      accepted.editor.state.tr.delete(0, accepted.editor.state.doc.firstChild!.nodeSize),
    )
    expect(collectRevisions(accepted.editor.state.doc).map((revision) => revision.kind)).toContain(
      'blockDel',
    )
    acceptAllRevisions(accepted.editor)
    expect(accepted.editor.state.doc.textContent).toBe('second')

    const rejected = await trackedDoc(body)
    rejected.editor.view.dispatch(
      rejected.editor.state.tr.delete(0, rejected.editor.state.doc.firstChild!.nodeSize),
    )
    rejectAllRevisions(rejected.editor)
    expect(rejected.editor.state.doc.textContent).toBe('firstsecond')
    accepted.editor.destroy()
    rejected.editor.destroy()
  })

  it('represents a whole-paragraph move as linked block deletion/insertion semantics', async () => {
    const body =
      '<w:p><w:r><w:t>A</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>B</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>C</w:t></w:r></w:p>'
    const moved = await trackedDoc(body)
    const [a, b, c] = Array.from({ length: moved.editor.state.doc.childCount }, (_, index) =>
      moved.editor.state.doc.child(index),
    )
    moved.editor.view.dispatch(
      moved.editor.state.tr.replaceWith(0, moved.editor.state.doc.content.size, [a, c, b]),
    )
    const kinds = collectRevisions(moved.editor.state.doc).map((revision) => revision.kind)
    expect(kinds).toContain('blockDel')
    expect(kinds).toContain('blockIns')
    acceptAllRevisions(moved.editor)
    expect(moved.editor.state.doc.textContent).toBe('ACB')
    moved.editor.destroy()
  })

  it('saves/reopens image deletion and table/page-break insertion revisions', async () => {
    const imageDoc = await trackedDoc(
      IMAGE_PARAGRAPH_XML + '<w:p><w:r><w:t>tail</w:t></w:r></w:p>',
      true,
    )
    imageDoc.editor.view.dispatch(
      imageDoc.editor.state.tr.delete(0, imageDoc.editor.state.doc.firstChild!.nodeSize),
    )
    expect(collectRevisions(imageDoc.editor.state.doc).map((revision) => revision.kind)).toContain(
      'blockDel',
    )
    const imagePlan = pmDocToSavePlan(imageDoc.editor.getJSON() as PmNode, imageDoc.parsed.blocks)
    expect(imagePlan.saveBlocks[0].revision?.kind).toBe('del')
    const reopenedImage = await parseDocx(await saveDocx(imageDoc.parsed, imagePlan.saveBlocks))
    expect(reopenedImage.internal.documentXml).toContain('<w:del ')
    const revisedImage = reopenedImage.blocks.find((block) => block.type === 'image')
    expect(revisedImage?.blockRevision?.kind).toBe('del')
    const reopenedImageEditor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(reopenedImage.blocks),
    })
    acceptAllRevisions(reopenedImageEditor)
    const acceptedImagePlan = pmDocToSavePlan(
      reopenedImageEditor.getJSON() as PmNode,
      reopenedImage.blocks,
    )
    const acceptedImage = await parseDocx(
      await saveDocx(reopenedImage, acceptedImagePlan.saveBlocks),
    )
    expect(acceptedImage.blocks.some((block) => block.type === 'image')).toBe(false)

    const inserted = await trackedDoc('<w:p><w:r><w:t>base</w:t></w:r></w:p>')
    const end = inserted.editor.state.doc.content.size
    inserted.editor.commands.insertContentAt(
      end,
      tableModelToPmNode({ rows: [[{ paras: ['cell'] }]] }),
    )
    inserted.editor.commands.insertContentAt(inserted.editor.state.doc.content.size, {
      type: 'docParagraph',
      attrs: { pageBreakBefore: true },
    })
    expect(
      collectRevisions(inserted.editor.state.doc).filter(
        (revision) => revision.kind === 'blockIns',
      ),
    ).toHaveLength(2)
    const plan = pmDocToSavePlan(inserted.editor.getJSON() as PmNode, inserted.parsed.blocks)
    const reopened = await parseDocx(await saveDocx(inserted.parsed, plan.saveBlocks))
    expect(reopened.blocks.filter((block) => block.blockRevision?.kind === 'ins')).toHaveLength(2)
    const reopenedEditor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(reopened.blocks),
    })
    rejectAllRevisions(reopenedEditor)
    // the untracked paragraph appended below the inserted table survives the reject
    expect(reopenedEditor.state.doc.childCount).toBe(2)
    expect(reopenedEditor.state.doc.lastChild?.textContent).toBe('')
    expect(reopenedEditor.state.doc.textContent).toBe('base')
    imageDoc.editor.destroy()
    reopenedImageEditor.destroy()
    inserted.editor.destroy()
    reopenedEditor.destroy()
  })

  it('records live row insertion and preserves it through DOCX reopen', async () => {
    const { editor, parsed } = await trackedDoc(TABLE_XML)
    let firstCell = -1
    editor.state.doc.descendants((node, pos) => {
      if (firstCell < 0 && node.type.name === 'docTableCell') firstCell = pos
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, firstCell + 2)),
    )
    expect(addRowAfter(editor.state, editor.view.dispatch)).toBe(true)
    expect(collectRevisions(editor.state.doc).map((revision) => revision.kind)).toContain('rowIns')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const savedBytes = await saveDocx(parsed, plan.saveBlocks)
    const reopened = await parseDocx(savedBytes)
    expect(
      reopened.blocks[0].table?.rowRevisions?.some((revision) => revision?.kind === 'ins'),
    ).toBe(true)
    // the save mints the row record's id from the shared revision counter
    const savedXml = await (
      await JSZip.loadAsync(savedBytes)
    )
      .file('word/document.xml')!
      .async('string')
    expect(savedXml).toMatch(/<w:trPr><w:ins w:id="9001" w:author="Structural Reviewer"/)
    expect(savedXml).not.toContain('w:id=""')
    editor.destroy()

    const deleted = await trackedDoc(
      TABLE_XML.replace('<w:tr>', '<w:tr><w:trPr><w:cantSplit/></w:trPr>'),
    )
    let deletedCell = -1
    deleted.editor.state.doc.descendants((node, pos) => {
      if (deletedCell < 0 && node.type.name === 'docTableCell') deletedCell = pos
    })
    deleted.editor.view.dispatch(
      deleted.editor.state.tr.setSelection(
        TextSelection.create(deleted.editor.state.doc, deletedCell + 2),
      ),
    )
    expect(deleteRow(deleted.editor.state, deleted.editor.view.dispatch)).toBe(true)
    expect(collectRevisions(deleted.editor.state.doc).map((revision) => revision.kind)).toContain(
      'rowDel',
    )
    // the record joins the row's existing trPr children instead of replacing them
    const deletedPlan = pmDocToSavePlan(deleted.editor.getJSON() as PmNode, deleted.parsed.blocks)
    const deletedXml = await (
      await JSZip.loadAsync(await saveDocx(deleted.parsed, deletedPlan.saveBlocks))
    )
      .file('word/document.xml')!
      .async('string')
    expect(deletedXml).toMatch(
      /<w:trPr><w:cantSplit\/><w:del w:id="9001" w:author="Structural Reviewer"/,
    )
    rejectAllRevisions(deleted.editor)
    expect(deleted.editor.state.doc.firstChild?.childCount).toBe(2)
    deleted.editor.destroy()
  })
})

describe('table row/cell revisions (rowIns/rowDel/cellIns/cellDel)', () => {
  async function openRevisionTable() {
    const bytes = await buildDocx({ bodyXml: REVISION_TABLE_XML })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    return { editor, parsed, bytes }
  }

  it('collectRevisions reports row delete and cell insert, inline delText not double-reported', async () => {
    const { editor } = await openRevisionTable()
    const revs = collectRevisions(editor.state.doc)
    expect(revs.filter((r) => r.kind === 'rowDel')).toHaveLength(1)
    expect(revs.filter((r) => r.kind === 'cellIns')).toHaveLength(1)
    expect(revs.find((r) => r.kind === 'rowDel')!.author).toBe('Alice')
    // delText inside a deleted row belongs to the row-level revision, no longer reported as a separate del
    expect(revs.filter((r) => r.kind === 'del')).toHaveLength(0)
    editor.destroy()
  })

  it('accept all: deleted row removed, cellIns marker stripped, save round-trip is clean', async () => {
    const { editor, parsed } = await openRevisionTable()
    acceptAllRevisions(editor)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const model = reparsed.blocks[0].table!
    expect(model.rows.length).toBe(2)
    expect(model.rows.map((r) => r.map((c) => c.paras[0]))).toEqual([
      ['A1', 'B1'],
      ['A3', 'B3'],
    ])
    expect(model.rowRevisions).toBeUndefined()
    expect(model.rows[1][0].cellRevision).toBeUndefined()
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })

  it('reject all: deleted row restored to body text, inserted cell removed, trPr revision record stripped', async () => {
    const { editor, parsed } = await openRevisionTable()
    rejectAllRevisions(editor)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const model = reparsed.blocks[0].table!
    expect(model.rows.length).toBe(3)
    // row 2 text restored to plain text (delText -> w:t, no del marker)
    expect(model.rows[1].map((c) => c.paras[0])).toEqual(['A2del', 'B2del'])
    expect(model.rows[1][0].richParas?.[0]?.runs?.[0]?.del).toBeUndefined()
    expect(model.rowRevisions).toBeUndefined()
    // rejecting cellIns = clear the cell's content (grid preserved, cell not removed outright)
    expect(model.rows[2].map((c) => c.paras[0])).toEqual(['', 'B3'])
    expect(model.rows[2][0].cellRevision).toBeUndefined()
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    editor.destroy()
  })
})

describe('recorder scope gating', () => {
  function trackingEditor(content: JsonNode[]): Editor {
    const editor = createEditor(content)
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Tester'
    return editor
  }
  const blockPos = (editor: Editor, index: number): number => {
    let pos = 0
    for (let i = 0; i < index; i++) pos += editor.state.doc.child(i).nodeSize
    return pos
  }

  it('mid-document edits record the same revisions and leave distant blocks unmarked', () => {
    const editor = trackingEditor(Array.from({ length: 20 }, (_, i) => para(text(`block ${i}`))))
    editor.commands.insertContentAt(blockPos(editor, 10) + 3, 'NEW')
    editor.commands.deleteRange({ from: blockPos(editor, 12) + 1, to: blockPos(editor, 12) + 3 })
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.map((r) => r.kind).sort()).toEqual(['del', 'ins'])
    const inBlock = (r: { from: number; to: number }, index: number) => {
      const start = blockPos(editor, index)
      return r.from >= start && r.to <= start + editor.state.doc.child(index).nodeSize
    }
    expect(ranges.every((r) => inBlock(r, 10) || inBlock(r, 12))).toBe(true)
    expect(editor.state.doc.childCount).toBe(20)
    editor.destroy()
  })

  it('records pPrChange for the edited paragraph only in a long document', () => {
    const editor = trackingEditor(Array.from({ length: 15 }, (_, i) => para(text(`p${i}`))))
    editor.commands.setTextSelection(blockPos(editor, 7) + 2)
    editor.commands.updateAttributes('docParagraph', { align: 'center' })
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toMatchObject([{ kind: 'pPrChange', author: 'Tester' }])
    expect(ranges[0].from).toBe(blockPos(editor, 7))
    editor.destroy()
  })
})

describe('IME composition deferral', () => {
  function composingEditor(content: JsonNode[]): { editor: Editor; end: () => void } {
    const editor = createEditor(content)
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Tester'
    let composing = true
    Object.defineProperty(editor.view, 'composing', {
      get: () => composing,
      configurable: true,
    })
    return { editor, end: () => (composing = false) }
  }

  it('defers recording while composing and marks the result after compositionend', () => {
    const { editor, end } = composingEditor([para(text('ab'))])
    editor.view.dispatch(editor.state.tr.insertText('n', 3, 3))
    editor.view.dispatch(editor.state.tr.insertText('ね', 3, 4))
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    end()
    editor.view.dispatch(editor.state.tr) // the compositionend poke
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(1)
    expect(ranges[0]).toMatchObject({ kind: 'ins', author: 'Tester' })
    expect(editor.state.doc.textBetween(ranges[0].from, ranges[0].to)).toBe('ね')
    editor.destroy()
  })

  it('records the selection a composition replaced as a struck deletion', () => {
    const { editor, end } = composingEditor([para(text('abcd'))])
    editor.view.dispatch(editor.state.tr.insertText('X', 2, 4)) // IME replaces "bc"
    end()
    editor.view.dispatch(editor.state.tr)
    expect(editor.state.doc.textContent).toBe('aXbcd')
    const kinds = collectRevisions(editor.state.doc)
      .map((r) => r.kind)
      .sort()
    expect(kinds).toEqual(['del', 'ins'])
    acceptAllRevisions(editor)
    expect(editor.state.doc.textContent).toBe('aXd')
    editor.destroy()
  })

  // Chromium dispatches the pinyin→hanzi commit after view.composing flips
  // false; it must fold into the composition, not strike the provisional text
  it('does not strike the provisional pinyin removed by the commit transaction', () => {
    const { editor, end } = composingEditor([para(text('ab'))])
    editor.view.dispatch(editor.state.tr.insertText('ni', 3, 3))
    end()
    editor.view.dispatch(editor.state.tr.insertText('你', 3, 5).setMeta('composition', 1))
    editor.view.dispatch(editor.state.tr) // the compositionend poke
    expect(editor.state.doc.textContent).toBe('ab你')
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(1)
    expect(ranges[0].kind).toBe('ins')
    expect(editor.state.doc.textBetween(ranges[0].from, ranges[0].to)).toBe('你')
    editor.destroy()
  })

  it('marks only the typed character when the IME re-read rewrites the whole run', () => {
    const { editor, end } = composingEditor([para(text('Hello world'))])
    // PM's DOM re-parse can replace the entire run around the composed char
    editor.view.dispatch(editor.state.tr.insertText('Hello 中world', 1, 12))
    end()
    editor.view.dispatch(editor.state.tr)
    expect(editor.state.doc.textContent).toBe('Hello 中world')
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(1)
    expect(ranges[0].kind).toBe('ins')
    expect(editor.state.doc.textBetween(ranges[0].from, ranges[0].to)).toBe('中')
    editor.destroy()
  })

  it('marks only the committed character when the commit rewrites the whole paragraph', () => {
    const { editor, end } = composingEditor([para(text('ab'))])
    editor.view.dispatch(editor.state.tr.insertText('ni', 3, 3))
    end()
    editor.view.dispatch(editor.state.tr.insertText('ab你', 1, 5).setMeta('composition', 1))
    editor.view.dispatch(editor.state.tr)
    expect(editor.state.doc.textContent).toBe('ab你')
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(1)
    expect(ranges[0]).toMatchObject({ kind: 'ins', author: 'Tester' })
    expect(editor.state.doc.textBetween(ranges[0].from, ranges[0].to)).toBe('你')
    acceptAllRevisions(editor)
    expect(editor.state.doc.textContent).toBe('ab你')
    editor.destroy()
  })
})

describe('repeated backspace under track changes', () => {
  function trackingEditor(content: JsonNode[]): Editor {
    const editor = createEditor(content)
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Tester'
    return editor
  }

  it('leaves the caret in front of the struck text so the next delete hits live text', () => {
    const editor = trackingEditor([para(text('ABCD'))])
    editor.commands.setTextSelection(4)
    editor.commands.deleteRange({ from: 3, to: 4 }) // backspace over C
    expect(editor.state.selection.from).toBe(3)

    // second backspace at the caret the editor actually has: must strike B too
    const caret = editor.state.selection.from
    editor.commands.deleteRange({ from: caret - 1, to: caret })
    const struck = collectRevisions(editor.state.doc)
      .filter((r) => r.kind === 'del')
      .map((r) => editor.state.doc.textBetween(r.from, r.to))
      .join('')
    expect(struck).toBe('BC')
    expect(editor.state.doc.textContent).toBe('ABCD')
    editor.destroy()
  })

  it('Backspace steps over an existing struck run instead of stalling on it', () => {
    const editor = trackingEditor([para(text('AB'), text('CD', [del('Tester')]))])
    // caret sits right after the struck "CD"
    editor.commands.setTextSelection(5)
    const handled = editor.commands.keyboardShortcut('Backspace')
    expect(handled).toBe(true)
    const struck = collectRevisions(editor.state.doc)
      .filter((r) => r.kind === 'del')
      .map((r) => editor.state.doc.textBetween(r.from, r.to))
      .join('')
    // B (the live character before the struck run) is now struck as well
    expect(struck).toContain('B')
    expect(editor.state.doc.textContent).toBe('ABCD')
    editor.destroy()
  })
})

describe('moves and paragraph marks (genoffice#2086)', () => {
  const AT = 'w:author="X" w:date="2026-01-01T00:00:00Z"'
  const R = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`
  const DR = (t: string) => `<w:r><w:delText xml:space="preserve">${t}</w:delText></w:r>`
  const MOVE_XML =
    `<w:p>${R('P0 intro')}</w:p>` +
    `<w:p>${R('P1 alpha ')}<w:moveFromRangeStart w:id="1" ${AT} w:name="m"/><w:moveFrom w:id="2" ${AT}>${DR('MOVED ')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>${R('beta')}</w:p>` +
    `<w:p>${R('P2 gamma ')}<w:moveToRangeStart w:id="3" ${AT} w:name="m"/><w:moveTo w:id="4" ${AT}>${R('MOVED ')}</w:moveTo><w:moveToRangeEnd w:id="3"/>${R('delta')}</w:p>` +
    `<w:p>${R('P3 tail')}</w:p>`
  const paragraphs = (editor: Editor) => {
    const out: string[] = []
    editor.state.doc.forEach((n) => out.push(n.textContent))
    return out
  }
  async function openEditor(bodyXml: string) {
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    return { parsed, editor }
  }
  async function savedXml(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const zip = await JSZip.loadAsync(await saveDocx(parsed, plan.saveBlocks))
    return zip.file('word/document.xml')!.async('string')
  }

  it('a move is two run-level revisions; accepting removes only the moved text at its source', async () => {
    const { parsed, editor } = await openEditor(MOVE_XML)
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.map((r) => [r.kind, editor.state.doc.textBetween(r.from, r.to)])).toEqual([
      ['moveFrom', 'MOVED '],
      ['moveTo', 'MOVED '],
    ])
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual([
      'P0 intro',
      'P1 alpha beta',
      'P2 gamma MOVED delta',
      'P3 tail',
    ])
    expect(collectRevisions(editor.state.doc)).toEqual([])
    expect(await savedXml(editor, parsed)).not.toMatch(/<w:(?:ins|del|moveFrom|moveTo)\b/)
    editor.destroy()

    const again = await openEditor(MOVE_XML)
    rejectAllRevisions(again.editor)
    expect(paragraphs(again.editor)).toEqual([
      'P0 intro',
      'P1 alpha MOVED beta',
      'P2 gamma delta',
      'P3 tail',
    ])
    again.editor.destroy()
  })

  it('an empty moved paragraph is a block-level move: gone on accept at the source, kept on reject', async () => {
    const xml =
      `<w:p>${R('before')}</w:p>` +
      `<w:p><w:moveFrom w:id="7" ${AT}><w:r><w:t></w:t></w:r></w:moveFrom></w:p>` +
      `<w:p>${R('after')}</w:p>`
    const { editor } = await openEditor(xml)
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['moveFrom'])
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['before', 'after'])
    expect(collectRevisions(editor.state.doc)).toEqual([])
    editor.destroy()

    const again = await openEditor(xml)
    rejectAllRevisions(again.editor)
    expect(paragraphs(again.editor)).toEqual(['before', '', 'after'])
    expect(collectRevisions(again.editor.state.doc)).toEqual([])
    again.editor.destroy()
  })

  it('a paragraph mark inside a move is a paragraph-mark revision', async () => {
    const xml =
      `<w:p><w:pPr><w:rPr><w:moveFrom w:id="8" ${AT}/></w:rPr></w:pPr><w:moveFrom w:id="9" ${AT}>${DR('gone')}</w:moveFrom></w:p>` +
      `<w:p>${R('after')}</w:p>`
    const { parsed, editor } = await openEditor(xml)
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual([
      'paraMarkDel',
      'moveFrom',
    ])
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['after'])
    expect(await savedXml(editor, parsed)).not.toMatch(/moveFrom|<w:del\b/)
    editor.destroy()
  })

  it('a moved paragraph mark with live text joins the next paragraph and leaves no move behind', async () => {
    const xml =
      `<w:p><w:pPr><w:rPr><w:moveFrom w:id="8" ${AT}/></w:rPr></w:pPr>${R('keep')}</w:p>` +
      `<w:p>${R('after')}</w:p>`
    const { parsed, editor } = await openEditor(xml)
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['paraMarkDel'])
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['keepafter'])
    expect(collectRevisions(editor.state.doc)).toEqual([])
    expect(await savedXml(editor, parsed)).not.toMatch(/moveFrom|moveRevision/)
    editor.destroy()
  })

  it('a deleted paragraph mark joins the next paragraph on accept and stays a mark on reject', async () => {
    const xml =
      `<w:p><w:pPr><w:rPr><w:del w:id="5" ${AT}/><w:b/></w:rPr></w:pPr>${R('one')}</w:p>` +
      `<w:p>${R('two')}</w:p><w:p>${R('three')}</w:p>`
    const { parsed, editor } = await openEditor(xml)
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['paraMarkDel'])
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['onetwo', 'three'])
    expect(collectRevisions(editor.state.doc)).toEqual([])
    const accepted = await savedXml(editor, parsed)
    expect(accepted).not.toMatch(/<w:del\b/)
    expect(accepted).toContain('<w:rPr><w:b/></w:rPr>')
    editor.destroy()

    const again = await openEditor(xml)
    rejectAllRevisions(again.editor)
    expect(paragraphs(again.editor)).toEqual(['one', 'two', 'three'])
    const rejected = await savedXml(again.editor, again.parsed)
    expect(rejected).not.toMatch(/<w:del\b/)
    expect(rejected).toContain('<w:rPr><w:b/></w:rPr>')
    again.editor.destroy()
  })

  it('an inserted paragraph mark becomes permanent on accept and joins on reject', async () => {
    const xml = `<w:p><w:pPr><w:rPr><w:ins w:id="6" ${AT}/></w:rPr></w:pPr>${R('one')}</w:p><w:p>${R('two')}</w:p>`
    const { parsed, editor } = await openEditor(xml)
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['paraMarkIns'])
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['one', 'two'])
    expect(await savedXml(editor, parsed)).not.toMatch(/<w:ins\b|<w:rPr>/)
    editor.destroy()

    const again = await openEditor(xml)
    rejectAllRevisions(again.editor)
    expect(paragraphs(again.editor)).toEqual(['onetwo'])
    expect(collectRevisions(again.editor.state.doc)).toEqual([])
    again.editor.destroy()
  })

  it('a format change on inserted text is its own revision and is accepted with the insertion', async () => {
    const xml =
      `<w:p>${R('P1 plain ')}<w:ins w:id="1" ${AT}><w:r><w:rPr><w:i/><w:rPrChange w:id="2" ${AT}><w:rPr/></w:rPrChange></w:rPr><w:t>inserted italic</w:t></w:r></w:ins></w:p>` +
      `<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="4" ${AT}><w:rPr/></w:rPrChange></w:rPr><w:t>P3 now bold</w:t></w:r></w:p>`
    const { parsed, editor } = await openEditor(xml)
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual([
      'ins',
      'rPrChange',
      'rPrChange',
    ])
    acceptAllRevisions(editor)
    expect(collectRevisions(editor.state.doc)).toEqual([])
    expect(paragraphs(editor)).toEqual(['P1 plain inserted italic', 'P3 now bold'])
    const accepted = await savedXml(editor, parsed)
    expect(accepted).not.toMatch(/<w:ins\b|<w:rPrChange/)
    expect(accepted).toContain('<w:i/>')
    editor.destroy()

    const again = await openEditor(xml)
    rejectAllRevisions(again.editor)
    expect(paragraphs(again.editor)).toEqual(['P1 plain ', 'P3 now bold'])
    expect(collectRevisions(again.editor.state.doc)).toEqual([])
    const rejected = await savedXml(again.editor, again.parsed)
    expect(rejected).not.toMatch(/<w:ins\b|<w:rPrChange|<w:b\/>/)
    again.editor.destroy()
  })

  it('a paragraph-mark format change is a pending revision; accept drops the record, reject restores the old mark', async () => {
    const xml =
      `<w:p><w:pPr><w:rPr><w:caps/><w:rPrChange w:id="3" ${AT}><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr></w:pPr>${R('P2 mark')}</w:p>` +
      `<w:p>${R('after')}</w:p>`
    const { parsed, editor } = await openEditor(xml)
    expect(parsed.blocks[0].paraMarkRPrChange).toMatchObject({ author: 'X', id: '3' })
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['paraMarkRPrChange'])
    acceptAllRevisions(editor)
    expect(collectRevisions(editor.state.doc)).toEqual([])
    expect(paragraphs(editor)).toEqual(['P2 mark', 'after'])
    const accepted = await savedXml(editor, parsed)
    expect(accepted).not.toMatch(/<w:rPrChange/)
    expect(accepted).toContain('<w:pPr><w:rPr><w:caps/></w:rPr></w:pPr>')
    editor.destroy()

    const again = await openEditor(xml)
    rejectAllRevisions(again.editor)
    expect(collectRevisions(again.editor.state.doc)).toEqual([])
    const rejected = await savedXml(again.editor, again.parsed)
    expect(rejected).not.toMatch(/<w:rPrChange|<w:caps/)
    expect(rejected).toContain('<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>')
    again.editor.destroy()
  })
})

describe('recorder: Delete, cross-paragraph cuts, paragraph marks, columns, AI author', () => {
  function trackingEditor(content: JsonNode[]): Editor {
    const editor = createEditor(content)
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Tester'
    return editor
  }
  const anchored = (idx: number, ...content: JsonNode[]): JsonNode => ({
    type: 'docParagraph',
    attrs: { docxIndex: idx },
    content,
  })
  const kinds = (editor: Editor) => collectRevisions(editor.state.doc).map((r) => r.kind)
  const struck = (editor: Editor) =>
    collectRevisions(editor.state.doc)
      .filter((r) => r.kind === 'del')
      .map((r) => editor.state.doc.textBetween(r.from, r.to))
  const paragraphs = (editor: Editor) => {
    const out: string[] = []
    editor.state.doc.forEach((n) => out.push(n.textContent))
    return out
  }
  const markRecord = (editor: Editor, index: number, attr: 'paraMarkDel' | 'paraMarkIns') =>
    JSON.parse((editor.state.doc.child(index).attrs[attr] as string | null) ?? 'null')
  // a real keydown: commands.keyboardShortcut captures the handler's transactions and drops their selections
  const press = (editor: Editor, key: string) =>
    !!editor.view.someProp('handleKeyDown', (f) =>
      f(editor.view, new KeyboardEvent('keydown', { key, bubbles: true })),
    )

  it('Delete steps over an existing struck run and strikes the next live character', () => {
    const editor = trackingEditor([para(text('AB'), text('CD', [del('Tester')]), text('EF'))])
    editor.commands.setTextSelection(3)
    expect(press(editor, 'Delete')).toBe(true)
    expect(struck(editor).join('')).toBe('CDE')
    expect(editor.state.doc.textContent).toBe('ABCDEF')
    // caret rests after the struck run, in front of the newly struck E
    expect(editor.state.selection.from).toBe(5)
    editor.destroy()
  })

  it('Delete before a struck tail with nothing live beyond it parks the caret after the run', () => {
    const editor = trackingEditor([para(text('AB'), text('CD', [del('Tester')]))])
    editor.commands.setTextSelection(3)
    expect(press(editor, 'Delete')).toBe(true)
    expect(editor.state.selection.from).toBe(5)
    expect(struck(editor).join('')).toBe('CD')
    editor.destroy()
  })

  it('deleting from inside one paragraph into the next strikes both parts and deletes the first mark', () => {
    const editor = trackingEditor([
      anchored(0, text('alpha beta')),
      anchored(1, text('gamma delta')),
    ])
    editor.commands.deleteRange({ from: 7, to: 19 })
    expect(paragraphs(editor)).toEqual(['alpha beta', 'gamma delta'])
    expect(kinds(editor)).toEqual(['paraMarkDel', 'del', 'del'])
    expect(struck(editor)).toEqual(['beta', 'gamma '])
    expect(markRecord(editor, 0, 'paraMarkDel')).toMatchObject({ author: 'Tester' })
    expect(editor.state.selection.from).toBe(7)
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['alpha delta'])

    const rejected = trackingEditor([
      anchored(0, text('alpha beta')),
      anchored(1, text('gamma delta')),
    ])
    rejected.commands.deleteRange({ from: 7, to: 19 })
    rejectAllRevisions(rejected)
    expect(paragraphs(rejected)).toEqual(['alpha beta', 'gamma delta'])
    expect(collectRevisions(rejected.state.doc)).toEqual([])
    editor.destroy()
    rejected.destroy()
  })

  it('typing over a cross-paragraph selection keeps the insertion live before the struck tail', () => {
    const editor = trackingEditor([
      anchored(0, text('alpha beta')),
      anchored(1, text('gamma delta')),
    ])
    editor.commands.insertContentAt({ from: 7, to: 19 }, 'X')
    expect(paragraphs(editor)).toEqual(['alpha Xbeta', 'gamma delta'])
    expect(kinds(editor)).toEqual(['paraMarkDel', 'ins', 'del', 'del'])
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['alpha Xdelta'])
    editor.destroy()
  })

  it('a cut across a middle paragraph strikes it whole and deletes its mark too', () => {
    const editor = trackingEditor([
      anchored(0, text('one')),
      anchored(1, text('mid')),
      anchored(2, text('three')),
    ])
    editor.commands.deleteRange({ from: 2, to: 12 })
    expect(paragraphs(editor)).toEqual(['one', 'mid', 'three'])
    expect(struck(editor)).toEqual(['ne', 'mid', 't'])
    expect(kinds(editor).filter((k) => k === 'paraMarkDel')).toHaveLength(2)
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['ohree'])
    editor.destroy()
  })

  it("a cut through this author's own inserted text and mark removes them for real", () => {
    const editor = trackingEditor([anchored(0, text('first')), anchored(1, text('second'))])
    editor.chain().setTextSelection(4).splitBlock().run()
    expect(paragraphs(editor)).toEqual(['fir', 'st', 'second'])
    expect(kinds(editor)).toEqual(['paraMarkIns'])
    editor.commands.deleteRange({ from: 3, to: 7 })
    expect(paragraphs(editor)).toEqual(['first', 'second'])
    expect(kinds(editor)).toEqual(['del'])
    expect(struck(editor)).toEqual(['rs'])
    editor.destroy()
  })

  it('Enter records the new paragraph mark as an insertion on the first half', () => {
    const editor = trackingEditor([anchored(0, text('alpha beta'))])
    editor.chain().setTextSelection(7).splitBlock().run()
    expect(paragraphs(editor)).toEqual(['alpha ', 'beta'])
    expect(kinds(editor)).toEqual(['paraMarkIns'])
    expect(markRecord(editor, 0, 'paraMarkIns')).toMatchObject({ author: 'Tester' })
    expect(editor.state.doc.child(1).attrs.paraMarkIns).toBeNull()
    rejectAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['alpha beta'])

    const plain = trackingEditor([para(text('ab'))])
    plain.chain().setTextSelection(2).splitBlock().run()
    expect(kinds(plain)).toEqual(['paraMarkIns'])
    editor.destroy()
    plain.destroy()
  })

  it('Backspace at a paragraph start deletes the previous mark instead of joining', () => {
    const editor = trackingEditor([anchored(0, text('first')), anchored(1, text('second'))])
    editor.commands.setTextSelection(8)
    expect(press(editor, 'Backspace')).toBe(true)
    expect(paragraphs(editor)).toEqual(['first', 'second'])
    expect(kinds(editor)).toEqual(['paraMarkDel'])
    expect(editor.state.selection.from).toBe(6)
    acceptAllRevisions(editor)
    expect(paragraphs(editor)).toEqual(['firstsecond'])
    editor.destroy()
  })

  it('Delete at a paragraph end deletes its own mark the same way', () => {
    const editor = trackingEditor([anchored(0, text('first')), anchored(1, text('second'))])
    editor.commands.setTextSelection(6)
    expect(press(editor, 'Delete')).toBe(true)
    expect(paragraphs(editor)).toEqual(['first', 'second'])
    expect(kinds(editor)).toEqual(['paraMarkDel'])
    editor.destroy()
  })

  it("Backspace onto this author's own inserted mark joins for real", () => {
    const editor = trackingEditor([anchored(0, text('alpha beta'))])
    editor.chain().setTextSelection(7).splitBlock().run()
    editor.commands.setTextSelection(9)
    expect(press(editor, 'Backspace')).toBe(true)
    expect(paragraphs(editor)).toEqual(['alpha beta'])
    expect(collectRevisions(editor.state.doc)).toEqual([])
    editor.destroy()
  })

  it('Backspace behind an already deleted mark steps over it and strikes the previous character', () => {
    const editor = trackingEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: 0, paraMarkDel: JSON.stringify({ author: 'Tester' }) },
        content: [text('ab')],
      },
      anchored(1, text('cd')),
    ])
    editor.commands.setTextSelection(5)
    expect(press(editor, 'Backspace')).toBe(true)
    expect(paragraphs(editor)).toEqual(['ab', 'cd'])
    expect(struck(editor)).toEqual(['b'])
    editor.destroy()
  })

  it('AI edits with the ribbon tracking on are credited to the Settings user name', () => {
    setAuthorName('leaf')
    const editor = trackingEditor([para(text('old text'))])
    replaceInlineRange(editor, 1, 4, [editor.schema.text('new')])
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.map((r) => r.kind).sort()).toEqual(['del', 'ins'])
    expect(ranges.every((r) => r.author === 'leaf')).toBe(true)
    setAuthorName('')
    editor.destroy()
  })

  async function trackedDoc(bodyXml: string) {
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Structural Reviewer'
    return { editor, parsed }
  }
  async function savedXml(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const bytes = await saveDocx(parsed, plan.saveBlocks)
    const zip = await JSZip.loadAsync(bytes)
    return { xml: await zip.file('word/document.xml')!.async('string'), bytes }
  }
  const selectFirstCell = (editor: Editor) => {
    let firstCell = -1
    editor.state.doc.descendants((node, pos) => {
      if (firstCell < 0 && node.type.name === 'docTableCell') firstCell = pos
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, firstCell + 2)),
    )
  }
  const P = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`

  it('recorder-created paragraph mark revisions save as pPr/rPr markers and reopen', async () => {
    const { editor, parsed } = await trackedDoc(P('first') + P('second') + P('third'))
    editor.commands.setTextSelection(8)
    expect(press(editor, 'Backspace')).toBe(true)
    editor.chain().setTextSelection(19).splitBlock().run()
    expect(paragraphs(editor)).toEqual(['first', 'second', 'thi', 'rd'])
    expect(kinds(editor)).toEqual(['paraMarkDel', 'paraMarkIns'])
    const { xml, bytes } = await savedXml(editor, parsed)
    const AUTHOR = 'w:author="Structural Reviewer" w:date="\\d{4}-[^"]+Z"'
    expect(xml).toMatch(
      new RegExp(
        `<w:p><w:pPr><w:rPr><w:del w:id="\\d+" ${AUTHOR}/></w:rPr></w:pPr><w:r><w:t[^>]*>first`,
      ),
    )
    expect(xml).toMatch(
      new RegExp(
        `<w:p><w:pPr><w:rPr><w:ins w:id="\\d+" ${AUTHOR}/></w:rPr></w:pPr><w:r><w:t[^>]*>thi<`,
      ),
    )
    const markIds = [
      ...xml.matchAll(/<w:(?:del|ins) w:id="(\d+)" w:author="Structural Reviewer"/g),
    ].map((m) => Number(m[1]))
    expect(markIds).toHaveLength(2)
    expect(new Set(markIds).size).toBe(2)
    expect(Math.min(...markIds)).toBeGreaterThanOrEqual(9001)
    const reopened = await parseDocx(bytes)
    expect(reopened.blocks[0].paraMarkDel).toMatchObject({ author: 'Structural Reviewer' })
    expect(reopened.blocks[2].paraMarkIns).toMatchObject({ author: 'Structural Reviewer' })
    expect(reopened.blocks[3].paraMarkIns).toBeUndefined()
    const again = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(reopened.blocks),
    })
    acceptAllRevisions(again)
    expect(paragraphs(again)).toEqual(['firstsecond', 'thi', 'rd'])
    expect(collectRevisions(again.state.doc)).toEqual([])
    editor.destroy()
    again.destroy()
  })

  it('inserting a column while tracking marks every new cell cellIns, saved as w:cellIns', async () => {
    const { editor, parsed } = await trackedDoc(TABLE_XML)
    selectFirstCell(editor)
    expect(insertRowsOrColumns(1, 'right', 600)(editor.state, editor.view.dispatch)).toBe(true)
    expect(kinds(editor)).toEqual(['cellIns', 'cellIns'])
    const { xml, bytes } = await savedXml(editor, parsed)
    expect(xml.match(/<w:cellIns w:id="\d+" w:author="Structural Reviewer"/g)).toHaveLength(2)
    const reopened = await parseDocx(bytes)
    const rows = reopened.blocks[0].table!.rows
    expect(rows.map((r) => r.length)).toEqual([3, 3])
    expect(rows.map((r) => r[1].cellRevision?.kind)).toEqual(['ins', 'ins'])
    const again = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(reopened.blocks),
    })
    rejectAllRevisions(again)
    const model = pmDocToSavePlan(again.getJSON() as PmNode, reopened.blocks)
    const saved = await parseDocx(await saveDocx(reopened, model.saveBlocks))
    expect(saved.blocks[0].table!.rows.map((r) => r.map((c) => c.paras[0]))).toEqual([
      ['A1', 'B1'],
      ['A2', 'B2'],
    ])
    editor.destroy()
    again.destroy()
  })

  it('deleting a column while tracking keeps it struck; accept removes the column', async () => {
    const { editor, parsed } = await trackedDoc(TABLE_XML)
    selectFirstCell(editor)
    expect(deleteTableColumn(editor.state, editor.view.dispatch)).toBe(true)
    expect(editor.state.doc.firstChild!.child(0).childCount).toBe(2)
    expect(kinds(editor)).toEqual(['cellDel', 'cellDel'])
    expect(editor.state.doc.rangeHasMark(0, 20, editor.schema.marks.del)).toBe(true)
    const { xml, bytes } = await savedXml(editor, parsed)
    expect(xml.match(/<w:cellDel w:id="\d+" w:author="Structural Reviewer"/g)).toHaveLength(2)
    expect(xml).toContain('<w:delText xml:space="preserve">A1</w:delText>')
    const reopened = await parseDocx(bytes)
    expect(reopened.blocks[0].table!.rows.map((r) => r[0].cellRevision?.kind)).toEqual([
      'del',
      'del',
    ])
    acceptAllRevisions(editor)
    expect(collectRevisions(editor.state.doc)).toEqual([])
    const table = editor.state.doc.firstChild!
    expect([table.child(0).childCount, table.child(1).childCount]).toEqual([1, 1])
    expect(table.textContent).toBe('B1B2')
    editor.destroy()
  })

  it('a deleted column beside an inserted row resolves on the right nodes (positions mapped)', async () => {
    const run = async (mode: 'accept' | 'reject') => {
      const { editor } = await trackedDoc(TABLE_XML)
      selectFirstCell(editor)
      expect(addRowAfter(editor.state, editor.view.dispatch)).toBe(true)
      selectFirstCell(editor)
      expect(deleteTableColumn(editor.state, editor.view.dispatch)).toBe(true)
      expect(kinds(editor).sort()).toEqual(['cellDel', 'cellDel', 'cellDel', 'rowIns'])
      if (mode === 'accept') acceptAllRevisions(editor)
      else rejectAllRevisions(editor)
      expect(collectRevisions(editor.state.doc)).toEqual([])
      const table = editor.state.doc.firstChild!
      const shape = Array.from({ length: table.childCount }, (_, i) => table.child(i).childCount)
      editor.destroy()
      return { shape, text: table.textContent }
    }
    expect(await run('accept')).toEqual({ shape: [1, 1, 1], text: 'B1B2' })
    expect(await run('reject')).toEqual({ shape: [2, 2], text: 'A1B1A2B2' })
  })

  it('a tracked column delete leaves a merged cell that reaches into a surviving column alone', async () => {
    const xml =
      '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
      '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>'
    const { editor } = await trackedDoc(xml)
    let a2 = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'A2') a2 = pos
    })
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, a2)))
    expect(deleteTableColumn(editor.state, editor.view.dispatch)).toBe(true)
    expect(kinds(editor)).toEqual(['cellDel'])
    expect(editor.state.doc.firstChild!.child(0).childCount).toBe(1)
    expect(struck(editor)).toEqual([])
    acceptAllRevisions(editor)
    // not a whole column: the cell clears in place, the merged cell keeps its span and text
    expect(editor.state.doc.firstChild!.child(0).child(0).attrs.colspan).toBe(2)
    expect(editor.state.doc.firstChild!.textContent).toBe('wideB2')
    editor.destroy()
  })
})

describe('moves (cut + paste, drag-drop)', () => {
  const trackedEditor = () => {
    const editor = createEditor([para(text('one two three')), para(text('four'))])
    const storage = editor.storage.trackChanges as TrackChangesStorage
    storage.enabled = true
    storage.author = 'Alice'
    return editor
  }
  const textOf = (editor: Editor, r: { from: number; to: number }) =>
    editor.state.doc.textBetween(r.from, r.to)
  // "two " sits at 5..9 in the first paragraph; the second paragraph's text ends at 20
  const cutThenPaste = (editor: Editor, pasted?: Fragment | ProseNode[]) => {
    const slice = editor.state.doc.slice(5, 9)
    editor.view.dispatch(editor.state.tr.delete(5, 9).setMeta('uiEvent', 'cut'))
    const end = editor.state.doc.content.size - 1
    editor.view.dispatch(
      editor.state.tr.insert(end, pasted ?? slice.content).setMeta('uiEvent', 'paste'),
    )
  }

  it('pasting the cut content records moveFrom + moveTo under one move id', () => {
    const editor = trackedEditor()
    cutThenPaste(editor)
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.map((r) => r.kind)).toEqual(['moveFrom', 'moveTo'])
    expect(ranges[0].moveId).toMatch(/^move\d+$/)
    expect(ranges[1].moveId).toBe(ranges[0].moveId)
    expect(ranges.map((r) => textOf(editor, r))).toEqual(['two ', 'two '])
    expect(ranges.every((r) => r.author === 'Alice')).toBe(true)
    expect(revisionCount(ranges)).toBe(1)
    editor.destroy()
  })

  it('pasting something else leaves a deletion and a plain insertion', () => {
    const editor = trackedEditor()
    cutThenPaste(editor, [editor.schema.text('other')])
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.map((r) => r.kind)).toEqual(['del', 'ins'])
    expect(ranges.some((r) => r.moveId)).toBe(false)
    editor.destroy()
  })

  it('a cut followed by typing is no longer a pending move', () => {
    const editor = trackedEditor()
    const slice = editor.state.doc.slice(5, 9)
    editor.view.dispatch(editor.state.tr.delete(5, 9).setMeta('uiEvent', 'cut'))
    editor.view.dispatch(editor.state.tr.insertText('x', 1))
    const end = editor.state.doc.content.size - 1
    editor.view.dispatch(editor.state.tr.insert(end, slice.content).setMeta('uiEvent', 'paste'))
    expect(collectRevisions(editor.state.doc).map((r) => r.kind)).toEqual(['ins', 'del', 'ins'])
    editor.destroy()
  })

  it('an unrelated deletion by the same author in the same second stays a deletion', () => {
    const editor = trackedEditor()
    // "four" -> strike "fo" (16..18) first, then cut "two " and paste it
    editor.view.dispatch(editor.state.tr.delete(16, 18))
    const slice = editor.state.doc.slice(5, 9)
    editor.view.dispatch(editor.state.tr.delete(5, 9).setMeta('uiEvent', 'cut'))
    editor.view.dispatch(
      editor.state.tr
        .insert(editor.state.doc.content.size - 1, slice.content)
        .setMeta('uiEvent', 'paste'),
    )
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.map((r) => [r.kind, textOf(editor, r)])).toEqual([
      ['moveFrom', 'two '],
      ['del', 'fo'],
      ['moveTo', 'two '],
    ])
    editor.destroy()
  })

  it('a drag-drop of a selection is one move', () => {
    const editor = trackedEditor()
    const slice = editor.state.doc.slice(5, 9)
    const tr = editor.state.tr.delete(5, 9)
    tr.insert(tr.doc.content.size - 1, slice.content).setMeta('uiEvent', 'drop')
    editor.view.dispatch(tr)
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges.map((r) => r.kind)).toEqual(['moveFrom', 'moveTo'])
    expect(ranges[0].moveId).toBe(ranges[1].moveId)
    editor.destroy()
  })

  it('accepting either side resolves the pair: source gone, destination kept', () => {
    for (const side of [0, 1]) {
      const editor = trackedEditor()
      cutThenPaste(editor)
      const ranges = collectRevisions(editor.state.doc)
      applyRevisions(editor, [ranges[side]], 'accept')
      expect(collectRevisions(editor.state.doc)).toHaveLength(0)
      expect(editor.state.doc.textContent).toBe('one threefourtwo ')
      editor.destroy()
    }
  })

  it('rejecting either side resolves the pair: source restored, destination removed', () => {
    for (const side of [0, 1]) {
      const editor = trackedEditor()
      cutThenPaste(editor)
      const ranges = collectRevisions(editor.state.doc)
      applyRevisions(editor, [ranges[side]], 'reject')
      expect(collectRevisions(editor.state.doc)).toHaveLength(0)
      expect(editor.state.doc.textContent).toBe('one two threefour')
      editor.destroy()
    }
  })

  it('a move saves as w:moveFrom / w:moveTo ranges sharing one name', async () => {
    const editor = trackedEditor()
    cutThenPaste(editor)
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p>',
      }),
    )
    const out = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    const xml = await (await JSZip.loadAsync(out)).file('word/document.xml')!.async('string')
    const names = [...xml.matchAll(/<w:move(?:From|To)RangeStart [^>]*w:name="([^"]+)"/g)].map(
      (m) => m[1],
    )
    expect(names).toHaveLength(2)
    expect(names[0]).toBe(names[1])
    expect(xml).toMatch(/<w:moveFrom [^>]*><w:r><w:delText[^>]*>two <\/w:delText>/)
    expect(xml).toMatch(/<w:moveTo [^>]*><w:r><w:t[^>]*>two <\/w:t>/)
    editor.destroy()
  })
})

describe('rPrChange reject restores the full previous rPr', () => {
  const OLD_RPR =
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:i/><w:color w:val="0000FF"/>' +
    '<w:spacing w:val="20"/><w:kern w:val="28"/><w:sz w:val="28"/>' +
    '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:lang w:val="fr-FR"/>'
  const P =
    `<w:p><w:r><w:rPr><w:b/><w:color w:val="FF0000"/>` +
    `<w:rPrChange w:id="9" w:author="Reviewer" w:date="2026-07-01T00:00:00Z"><w:rPr>${OLD_RPR}</w:rPr></w:rPrChange>` +
    `</w:rPr><w:t>x</w:t></w:r></w:p>`

  async function resolveAndSave(mode: 'accept' | 'reject') {
    const bytes = await buildDocx({ bodyXml: P })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const rpr = collectRevisions(editor.state.doc).find((r) => r.kind === 'rPrChange')!
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, rpr.from + 1)),
    )
    if (mode === 'reject') rejectCurrentRevision(editor)
    else acceptCurrentRevision(editor)
    expect(collectRevisions(editor.state.doc)).toHaveLength(0)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const zip = await JSZip.loadAsync(await saveDocx(parsed, plan.saveBlocks))
    const xml = await zip.file('word/document.xml')!.async('string')
    const marks = (() => {
      let out: Record<string, unknown> = {}
      editor.state.doc.descendants((node) => {
        if (!node.isText) return
        const ts = node.marks.find((m) => m.type.name === 'docTextStyle')
        out = { ...(ts?.attrs ?? {}), bold: node.marks.some((m) => m.type.name === 'bold') }
      })
      return out
    })()
    editor.destroy()
    return { xml, marks }
  }

  it('reject writes the recorded rPr verbatim (spacing, kern, shd, lang included) and previews it', async () => {
    const { xml, marks } = await resolveAndSave('reject')
    expect(xml).toContain(`<w:r><w:rPr>${OLD_RPR}</w:rPr><w:t`)
    expect(marks).toMatchObject({
      bold: false,
      color: '0000FF',
      fontAscii: 'Arial',
      sizeHalfPoints: 28,
      kernHalfPoints: 28,
      shading: 'FFFF00',
    })
  })

  it('reject keeps values the modeled subset used to drop (automatic color, zero spacing)', async () => {
    const OLD = '<w:color w:val="auto"/><w:spacing w:val="0"/>'
    const bytes = await buildDocx({
      bodyXml:
        `<w:p><w:r><w:rPr><w:color w:val="FF0000"/>` +
        `<w:rPrChange w:id="9" w:author="Reviewer"><w:rPr>${OLD}</w:rPr></w:rPrChange>` +
        `</w:rPr><w:t>x</w:t></w:r></w:p>`,
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const rpr = collectRevisions(editor.state.doc).find((r) => r.kind === 'rPrChange')!
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, rpr.from + 1)),
    )
    rejectCurrentRevision(editor)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const zip = await JSZip.loadAsync(await saveDocx(parsed, plan.saveBlocks))
    expect(await zip.file('word/document.xml')!.async('string')).toContain(
      `<w:r><w:rPr>${OLD}</w:rPr><w:t`,
    )
    editor.destroy()
  })

  it('accept strips only the record', async () => {
    const { xml, marks } = await resolveAndSave('accept')
    expect(xml).toContain('<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t')
    expect(marks).toMatchObject({ bold: true, color: 'FF0000' })
  })
})
