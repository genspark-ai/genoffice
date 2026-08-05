import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
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

import {
  acceptAllRevisions,
  acceptCurrentRevision,
  collectRevisions,
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

  it('treats ins+del on the same text as one "both" revision', () => {
    const editor = createEditor([para(text('ins-then-del', [ins(), del()]))])
    const ranges = collectRevisions(editor.state.doc)
    expect(ranges).toHaveLength(1)
    expect(ranges[0].kind).toBe('both')
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
    expect(gotoRevision(editor, 1)).toBe(true)
    const first = editor.state.selection
    expect(editor.state.doc.textBetween(first.from, first.to)).toBe('IN')
    expect(gotoRevision(editor, 1)).toBe(true)
    expect(
      editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to),
    ).toBe('DE')
    editor.destroy()
  })

  it('returns false with no revisions', () => {
    const editor = createEditor([para(text('plain'))])
    expect(acceptCurrentRevision(editor)).toBe(false)
    expect(gotoRevision(editor, 1)).toBe(false)
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
    expect(reopenedEditor.state.doc.childCount).toBe(1)
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
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(
      reopened.blocks[0].table?.rowRevisions?.some((revision) => revision?.kind === 'ins'),
    ).toBe(true)
    editor.destroy()

    const deleted = await trackedDoc(TABLE_XML)
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
