import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import {
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  parseDocx,
  saveDocx,
} from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { buildDocContext, getSelectionScope } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'

/**
 * Guard tests for issue #175: an AI rewrite of selected text (replace_blocks)
 * used to reset the rewritten paragraphs to the style defaults — the model
 * writes restricted HTML that carries no font/size/indent, and the old blocks
 * were dropped wholesale. New blocks now inherit the replaced blocks'
 * paragraph attrs and dominant run style; the block's own HTML-derived
 * formatting (<pre>/<blockquote> presets, <strong>…) still wins.
 */

const NUM_IDS = { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID }
const TRACK = { author: 'AI Assistant' }

/** SimSun 12pt body, Times New Roman Latin, 2-char first-line indent, justified, 1.5 lines */
const BODY_STYLE = {
  font: 'SimSun',
  fontAscii: 'Times New Roman',
  sizeHalfPoints: 24,
  color: '333333',
}
const BODY_ATTRS = { indentFirstLine: 480, align: 'justify', lineSpacing: 1.5 }

type BlockJson = {
  type: string
  attrs: Record<string, unknown>
  content?: Array<{
    type: string
    text?: string
    marks?: Array<{ type: string; attrs?: Record<string, unknown> }>
  }>
}

async function createEditor(blocks: PmNode[]) {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = await parseDocx(await buildBlankDocx())
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  const pm = blocks.map((b) => editor.schema.nodeFromJSON(b))
  editor.view.dispatch(editor.state.tr.replaceWith(0, editor.state.doc.content.size, pm))
  return { editor, parsed }
}

function para(
  text: string,
  attrs: Record<string, unknown> = BODY_ATTRS,
  style: Record<string, unknown> | null = BODY_STYLE,
  docxIndex: number | null = null,
): PmNode {
  return {
    type: 'docParagraph',
    attrs: { ...attrs, docxIndex },
    content: [
      { type: 'text', text, ...(style ? { marks: [{ type: 'docTextStyle', attrs: style }] } : {}) },
    ],
  }
}

function blocks(editor: Editor): BlockJson[] {
  return ((editor.getJSON() as PmNode).content ?? []) as BlockJson[]
}

function textStyle(block: BlockJson, i = 0): Record<string, unknown> | undefined {
  return block.content?.[i].marks?.find((m) => m.type === 'docTextStyle')?.attrs
}

async function replaceBlocks(
  editor: Editor,
  start: number,
  end: number,
  html: string,
  track?: typeof TRACK,
) {
  const r = await executeTool(
    editor,
    {
      id: 't',
      name: 'replace_blocks',
      input: { startBlockIndex: start, endBlockIndex: end, html },
    },
    NUM_IDS,
    track,
  )
  expect(r.isError).toBeFalsy()
  return r
}

describe('replace_blocks keeps the replaced blocks formatting (issue #175)', () => {
  it('a 1:1 rewrite inherits paragraph attrs, the dominant run style and the docx anchor', async () => {
    const styled = {
      ...BODY_ATTRS,
      styleId: 'a3',
      spaceAfterAuto: false,
      tabStops: '[{"pos":4320,"val":"left"}]',
    }
    const { editor } = await createEditor([
      para('A long-winded original paragraph.', styled, BODY_STYLE, 7),
    ])
    await replaceBlocks(editor, 0, 0, '<p>Condensed <strong>paragraph</strong>.</p>')

    const [p] = blocks(editor)
    expect(p.type).toBe('docParagraph')
    expect(p.attrs).toMatchObject({ ...styled, docxIndex: 7, aiChanged: true })
    expect(p.content?.map((c) => c.text)).toEqual(['Condensed ', 'paragraph', '.'])
    for (let i = 0; i < 3; i++) expect(textStyle(p, i)).toMatchObject(BODY_STYLE)
    // the HTML's own inline marks are still applied on top
    expect(p.content?.[1].marks?.some((m) => m.type === 'bold')).toBe(true)
    expect(p.content?.[0].marks?.some((m) => m.type === 'bold')).toBe(false)
  })

  it('the selection context never shows the formatting the model would need to restore it', async () => {
    const { editor } = await createEditor([para('Original text.', BODY_ATTRS, BODY_STYLE, 0)])
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, 1, editor.state.doc.content.size - 1),
      ),
    )
    const ctx = buildDocContext(editor, getSelectionScope(editor))
    expect(ctx).toContain('<p>Original text.</p>')
    expect(ctx).not.toMatch(/SimSun|Times New Roman|480|justify/)
  })

  it('condensing several paragraphs into one keeps the first paragraph formatting', async () => {
    const { editor } = await createEditor([
      para('First.', BODY_ATTRS, BODY_STYLE, 3),
      para('Second.', { align: 'center' }, { sizeHalfPoints: 20 }, 4),
      para('Third.', { align: 'right' }, null, 5),
    ])
    await replaceBlocks(editor, 0, 2, '<p>All three in one.</p>')
    const all = blocks(editor)
    expect(all).toHaveLength(1)
    expect(all[0].attrs).toMatchObject({ ...BODY_ATTRS, docxIndex: 3 })
    expect(textStyle(all[0])).toMatchObject(BODY_STYLE)
  })

  it('expanding one paragraph into several gives every new paragraph the same formatting; the anchor is lent once', async () => {
    const { editor } = await createEditor([para('Seed.', BODY_ATTRS, BODY_STYLE, 2)])
    await replaceBlocks(editor, 0, 0, '<p>One.</p><p>Two.</p><p>Three.</p>')
    const all = blocks(editor)
    expect(all).toHaveLength(3)
    for (const p of all) {
      expect(p.attrs).toMatchObject(BODY_ATTRS)
      expect(textStyle(p)).toMatchObject(BODY_STYLE)
    }
    // docxIndex must stay unique: the save plan and the TrackChanges recorder key blocks by it
    expect(all.map((p) => p.attrs.docxIndex)).toEqual([2, null, null])
  })

  it('a page break before the replaced paragraph lands on the first new paragraph only', async () => {
    const { editor } = await createEditor([
      para('Intro.', {}, null, 0),
      para('Chapter opener.', { ...BODY_ATTRS, pageBreakBefore: true }, BODY_STYLE, 1),
    ])
    await replaceBlocks(editor, 1, 1, '<p>One.</p><p>Two.</p><p>Three.</p>')
    const all = blocks(editor)
    expect(all.map((p) => p.attrs.pageBreakBefore)).toEqual([false, true, false, false])
    for (const p of all.slice(1)) expect(p.attrs).toMatchObject(BODY_ATTRS)
  })

  it('a block the model turned into a heading is left as parsed; same-role neighbours still inherit', async () => {
    const { editor } = await createEditor([
      para('Intro paragraph.', BODY_ATTRS, BODY_STYLE, 0),
      para('Body paragraph.', BODY_ATTRS, BODY_STYLE, 1),
    ])
    await replaceBlocks(editor, 0, 1, '<h2>New heading</h2><p>Rewritten body.</p>')
    const [h, p] = blocks(editor)
    expect(h.type).toBe('docHeading')
    expect(h.attrs.indentFirstLine).toBeNull()
    expect(h.attrs.docxIndex).toBeNull()
    expect(textStyle(h)).toBeUndefined()
    expect(p.type).toBe('docParagraph')
    expect(p.attrs).toMatchObject({ ...BODY_ATTRS, docxIndex: 0 })
    expect(textStyle(p)).toMatchObject(BODY_STYLE)
  })

  it('headings inherit only from a heading of the same level', async () => {
    const { editor } = await createEditor([
      {
        type: 'docHeading',
        attrs: { level: 2, docxIndex: 9, spaceBefore: 240, styleId: '2' },
        content: [
          {
            type: 'text',
            text: 'Old H2',
            marks: [{ type: 'docTextStyle', attrs: { font: 'SimHei' } }],
          },
        ],
      },
    ])
    await replaceBlocks(editor, 0, 0, '<h2>New H2</h2>')
    expect(blocks(editor)[0].attrs).toMatchObject({
      level: 2,
      docxIndex: 9,
      spaceBefore: 240,
      styleId: '2',
    })
    expect(textStyle(blocks(editor)[0])).toMatchObject({ font: 'SimHei' })

    await replaceBlocks(editor, 0, 0, '<h3>Demoted</h3>')
    const h3 = blocks(editor)[0]
    expect(h3.attrs).toMatchObject({ level: 3, docxIndex: null, spaceBefore: null, styleId: null })
    expect(textStyle(h3)).toBeUndefined()
  })

  it('<pre> and <blockquote> presets win over inherited values, the rest is inherited', async () => {
    const { editor } = await createEditor([
      para('Code-ish.', BODY_ATTRS, BODY_STYLE, 0),
      para('Quote-ish.', BODY_ATTRS, BODY_STYLE, 1),
    ])
    await replaceBlocks(editor, 0, 1, '<pre>x = 1</pre><blockquote>Said someone.</blockquote>')
    const [pre, quote] = blocks(editor)
    expect(pre.attrs).toMatchObject({
      shadingFill: 'F2F2F2',
      borders: 'tblr',
      indentFirstLine: 480,
    })
    expect(textStyle(pre)).toMatchObject({ ...BODY_STYLE, fontAscii: 'Consolas' })
    expect(quote.attrs).toMatchObject({ indentLeft: 720, borders: 'l', align: 'justify' })
    expect(textStyle(quote)).toMatchObject({ ...BODY_STYLE, color: '666666' })
  })

  it('the run style comes from the run with the most text, not the first run', async () => {
    const { editor } = await createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: 0 },
        content: [
          {
            type: 'text',
            text: 'Note:',
            marks: [
              { type: 'bold' },
              { type: 'docTextStyle', attrs: { font: 'SimHei', sizeHalfPoints: 28 } },
            ],
          },
          {
            type: 'text',
            text: ' the body of the note runs on for a good while.',
            marks: [{ type: 'docTextStyle', attrs: BODY_STYLE }],
          },
        ],
      },
    ])
    await replaceBlocks(editor, 0, 0, '<p>Short note.</p>')
    expect(textStyle(blocks(editor)[0])).toEqual(expect.objectContaining(BODY_STYLE))
    expect(textStyle(blocks(editor)[0])?.font).toBe('SimSun')
  })

  it('a tracked rewrite gives the inserted (ins) blocks the same inheritance', async () => {
    const { editor } = await createEditor([para('Tracked original.', BODY_ATTRS, BODY_STYLE, 4)])
    await replaceBlocks(editor, 0, 0, '<p>Tracked rewrite.</p>', TRACK)
    const all = blocks(editor)
    expect(all).toHaveLength(2)
    const [old, fresh] = all
    expect(old.content?.[0].marks?.some((m) => m.type === 'del')).toBe(true)
    expect(fresh.content?.[0].marks?.some((m) => m.type === 'ins')).toBe(true)
    // the struck-through original stays in the document and keeps its anchor
    expect(old.attrs.docxIndex).toBe(4)
    expect(fresh.attrs).toMatchObject({ ...BODY_ATTRS, docxIndex: null, aiChanged: false })
    expect(textStyle(fresh)).toMatchObject(BODY_STYLE)
  })

  it('tracked-deleted blocks in the range never serve as formatting templates', async () => {
    const { editor } = await createEditor([
      {
        type: 'docParagraph',
        attrs: { docxIndex: 0, align: 'center' },
        content: [
          {
            type: 'text',
            text: 'Gone.',
            marks: [
              { type: 'del', attrs: { author: 'x', date: '2026-01-01T00:00:00Z' } },
              { type: 'docTextStyle', attrs: { sizeHalfPoints: 60 } },
            ],
          },
        ],
      },
      para('Live.', BODY_ATTRS, BODY_STYLE, 1),
    ])
    await replaceBlocks(editor, 0, 1, '<p>Rewritten.</p>')
    const [p] = blocks(editor)
    expect(p.attrs).toMatchObject({ ...BODY_ATTRS, docxIndex: 1 })
    expect(textStyle(p)).toMatchObject(BODY_STYLE)
  })

  it('list items of the same kind stay in the original numbering; a kind change takes the new list', async () => {
    const li = (text: string, kind: string, numId: string, docxIndex: number): PmNode => ({
      type: 'docListItem',
      attrs: { kind, numId, ilvl: 0, docxIndex, indentLeft: 1000 },
      content: [{ type: 'text', text, marks: [{ type: 'docTextStyle', attrs: BODY_STYLE }] }],
    })
    const { editor } = await createEditor([
      li('a', 'ordered', '42', 0),
      li('b', 'ordered', '42', 1),
    ])
    await replaceBlocks(editor, 0, 1, '<ol><li>A</li><li>B</li></ol>')
    for (const item of blocks(editor)) {
      expect(item.attrs).toMatchObject({ kind: 'ordered', numId: '42', indentLeft: 1000 })
      expect(textStyle(item)).toMatchObject(BODY_STYLE)
    }
    await replaceBlocks(editor, 0, 1, '<ul><li>A</li><li>B</li></ul>')
    for (const item of blocks(editor)) {
      expect(item.attrs).toMatchObject({
        kind: 'bullet',
        numId: BLANK_BULLET_NUM_ID,
        indentLeft: 1000,
      })
    }
  })

  it('end to end: font, size and indent survive a rewrite through save and reparse', async () => {
    // build a docx whose paragraph carries direct formatting, then reopen it
    const { editor: author, parsed: blank } = await createEditor([
      para('The original paragraph, rather verbose, that the user wants condensed.'),
    ])
    const authored = await saveDocx(
      blank,
      pmDocToSavePlan(author.getJSON() as PmNode, blank.blocks).saveBlocks,
    )
    const parsed = await parseDocx(authored)
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const before = parsed.blocks.find((b) => b.type === 'paragraph')!
    expect(before.format?.indentFirstLine).toBe(480)
    expect(before.runs?.[0]).toMatchObject({
      font: 'SimSun',
      fontAscii: 'Times New Roman',
      sizeHalfPoints: 24,
    })

    const idx = ((editor.getJSON() as PmNode).content ?? []).findIndex(
      (n) => n.type === 'docParagraph',
    )
    await replaceBlocks(editor, idx, idx, '<p>Condensed.</p>')
    const saved = await saveDocx(
      parsed,
      pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
    )
    const after = (await parseDocx(saved)).blocks.find((b) => b.type === 'paragraph')!
    expect(after.runs?.map((r) => r.text).join('')).toBe('Condensed.')
    expect(after.format).toMatchObject({ indentFirstLine: 480, align: 'justify', lineSpacing: 1.5 })
    expect(after.runs?.[0]).toMatchObject({
      font: 'SimSun',
      fontAscii: 'Times New Roman',
      sizeHalfPoints: 24,
      color: '333333',
    })
  })
})
