import { afterEach, describe, expect, it } from 'vitest'
import { Editor, type JSONContent } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  blocksToPmDoc,
  pmDocToSavePlan,
  tableModelToPmNode,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { acceptAllRevisions, rejectAllRevisions } from '../src/renderer/editor/revisions'
import {
  buildDocContext,
  getSelectionScope,
  partialSelection,
  serializeRangeToHtml,
} from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'

const NUM_IDS = { bullet: null, ordered: null }
const TRACK = { author: 'AI Assistant' }
const TARGET = '细微裂缝'
const ORIGINAL = `桥面存在${TARGET}，建议进一步检查。`
const REPLACEMENT = '局部细裂纹'
const editors: Editor[] = []
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

function createEditor(content: JSONContent) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    editorProps: { handleScrollToSelection: () => true },
    content,
  })
  editors.push(editor)
  return editor
}

async function fixture(pending = true) {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>二、病害描述</w:t></w:r></w:p>' +
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${ORIGINAL}</w:t></w:r></w:p></w:tc>` +
        '<w:tc><w:p><w:r><w:t>相邻单元格保持不变。</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        '<w:p><w:r><w:t>表格外文字保持不变。</w:t></w:r></w:p>',
    }),
  )
  const editor = createEditor(blocksToPmDoc(parsed.blocks))
  if (pending) {
    const replaced = await executeTool(
      editor,
      {
        id: 'whole-table',
        name: 'replace_blocks',
        input: {
          startBlockIndex: 1,
          endBlockIndex: 1,
          html: `<table><tr><td>${ORIGINAL}</td><td>相邻单元格保持不变。</td></tr></table>`,
        },
      },
      NUM_IDS,
      TRACK,
    )
    expect(replaced.isError).not.toBe(true)
  }
  return { editor, parsed, block: pending ? 2 : 1 }
}

function select(editor: Editor, block: number, text = TARGET) {
  let from = -1
  let index = 0
  editor.state.doc.forEach((node, offset) => {
    if (index++ !== block) return
    node.descendants((child, pos) => {
      if (child.isText && child.text!.includes(text))
        from = offset + 1 + pos + child.text!.indexOf(text)
    })
  })
  expect(from).toBeGreaterThan(0)
  const range = { from, to: from + text.length }
  editor.commands.setTextSelection(range)
  expect(editor.state.doc.textBetween(from, range.to)).toBe(text)
  return range
}

const paragraph = (content: JSONContent[]): JSONContent => ({ type: 'docParagraph', content })
const text = (value: string, marks?: JSONContent['marks']): JSONContent => ({
  type: 'text',
  text: value,
  marks,
})
const revision = (kind: 'ins' | 'del') => ({
  kind,
  author: 'Reviewer',
  date: '2026-11-01T00:00:00Z',
})
function cellFixture(content: JSONContent[]) {
  return createEditor({
    type: 'doc',
    content: [
      {
        type: 'docTable',
        content: [
          {
            type: 'docTableRow',
            content: [
              { type: 'docTableCell', content: [paragraph(content)] },
              { type: 'docTableCell', content: [paragraph([text('other cell')])] },
            ],
          },
        ],
      },
    ],
  })
}

describe('AI table cell selection through the SDK runtime', () => {
  it('reads the exact styled span and keeps the frozen action on the new table', async () => {
    const { editor, block } = await fixture()
    const range = select(editor, block)
    expect(partialSelection(editor, getSelectionScope(editor))).toEqual(range)
    const runtime = createDocsRuntime({
      current: () => ({ editor, identity: editor, numIds: NUM_IDS }),
    })
    const router = createDocsCommandRouter(runtime)
    router.bindScope('cell-action', runtime.capture())
    select(editor, 1)
    try {
      const result = (await router.handle({
        requestId: 'read',
        callerId: 'harness',
        bindingId: 'cell-action',
        command: 'read_document',
        payload: {},
      })) as { text: string }
      expect(result.text).toContain(`Current selection: block ${block} (part of the text only`)
      expect(result.text).toContain(`Selected text (${TARGET.length} characters): "${TARGET}"`)
      expect(result.text).toContain(`<sel><strong>${TARGET}</strong></sel>`)
      const replaced = (await router.handle({
        requestId: 'replace',
        callerId: 'harness',
        bindingId: 'cell-action',
        command: 'execute_tool',
        payload: { name: 'replace_selection', input: { html: REPLACEMENT } },
      })) as { mutated: boolean }
      expect(replaced.mutated).toBe(true)
      expect(editor.state.doc.child(block).textContent).toContain(
        ORIGINAL.replace(TARGET, REPLACEMENT),
      )
      expect(editor.state.doc.child(1).textContent).toContain(ORIGINAL)
    } finally {
      router.dispose()
      runtime.dispose()
    }
  })

  it.each([
    [false, false],
    [false, true],
    [true, false],
    [true, true],
  ])(
    'changes only the selected text, saves and resolves revisions (pending=%s, track=%s)',
    async (pending, tracking) => {
      const { editor, parsed, block } = await fixture(pending)
      const range = select(editor, block)
      const before = editor.state.doc
      const runtime = createDocsRuntime({
        current: () => ({
          editor,
          identity: editor,
          numIds: NUM_IDS,
          ...(tracking ? { track: TRACK } : {}),
        }),
      })
      const client = runtime.connect({ id: 'harness' })
      try {
        client.context()
        const result = await client.execute({
          id: 'replace',
          name: 'replace_selection',
          input: { html: REPLACEMENT },
        })
        expect(result.isError, result.output).not.toBe(true)
        expect(result.mutated).toBe(true)
        expect(serializeRangeToHtml(editor, block, block)).toContain(
          ORIGINAL.replace(TARGET, REPLACEMENT),
        )
        expect(
          editor.state.doc.child(block).child(0).child(1).eq(before.child(block).child(0).child(1)),
        ).toBe(true)
        expect(editor.state.doc.firstChild!.eq(before.firstChild!)).toBe(true)
        expect(editor.state.doc.lastChild!.eq(before.lastChild!)).toBe(true)
        if (pending) expect(editor.state.doc.child(1).eq(before.child(1))).toBe(true)
        const selected = editor.state.doc.resolve(editor.state.selection.from).nodeAfter!
        expect(
          editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to),
        ).toBe(REPLACEMENT)
        expect(selected.marks.some((mark) => mark.type.name === 'bold')).toBe(true)
        expect(editor.state.selection.to - editor.state.selection.from).toBe(REPLACEMENT.length)
        if (!tracking) expect(editor.state.selection.from).toBe(range.from)
        const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
        const reloaded = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
        const reopened = createEditor(blocksToPmDoc(reloaded.blocks))
        expect(serializeRangeToHtml(reopened, block, block)).toContain(
          ORIGINAL.replace(TARGET, REPLACEMENT),
        )
        if (pending || tracking) {
          acceptAllRevisions(reopened)
          expect(reopened.state.doc.textContent).toContain(ORIGINAL.replace(TARGET, REPLACEMENT))
          expect(reopened.state.doc.textContent).not.toContain(TARGET)
          rejectAllRevisions(editor)
          expect(editor.state.doc.textContent).toContain(ORIGINAL)
        }
      } finally {
        runtime.dispose()
      }
    },
  )

  it('keeps a full paragraph within a cell a partial table selection', () => {
    const editor = cellFixture([text(ORIGINAL)])
    const range = select(editor, 0, ORIGINAL)
    expect(partialSelection(editor, getSelectionScope(editor))).toEqual(range)
    expect(buildDocContext(editor)).toContain(`<sel>${ORIGINAL}</sel>`)
  })

  it.each(['blockRevision', 'rowRevision', 'cellRevision'])(
    'refuses text inside a pending deletion (%s)',
    async (attribute) => {
      const editor = cellFixture([text(ORIGINAL)])
      const range = select(editor, 0)
      const depth = attribute === 'blockRevision' ? 1 : attribute === 'rowRevision' ? 2 : 3
      const $from = editor.state.doc.resolve(range.from)
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup($from.before(depth), undefined, {
          ...$from.node(depth).attrs,
          [attribute]: revision('del'),
        }),
      )
      const before = editor.state.doc
      const result = await executeTool(
        editor,
        { id: 'replace', name: 'replace_selection', input: { html: REPLACEMENT } },
        NUM_IDS,
        TRACK,
      )
      expect(result.isError).toBe(true)
      expect(result.output).toContain('pending tracked deletion')
      expect(editor.state.doc.eq(before)).toBe(true)
    },
  )

  it.each([false, true])(
    'preserves existing inline deletion revisions when replacing live selected text (track=%s)',
    async (tracking) => {
      const editor = cellFixture([
        text('before '),
        text('old', [{ type: 'del', attrs: revision('del') }, { type: 'bold' }]),
        text('current', [{ type: 'italic' }]),
        text(' after'),
      ])
      const range = select(editor, 0, 'old')
      editor.commands.setTextSelection({ from: range.from, to: range.to + 'current'.length })
      expect(buildDocContext(editor)).toContain('Selected text (7 characters): "current"')
      const result = await executeTool(
        editor,
        { id: 'replace', name: 'replace_selection', input: { html: 'new' } },
        NUM_IDS,
        tracking ? TRACK : undefined,
      )
      expect(result.isError, result.output).not.toBe(true)
      const insertedMarks = editor.state.doc
        .resolve(editor.state.selection.from)
        .nodeAfter!.marks.map((mark) => mark.type.name)
      expect(insertedMarks).toContain('italic')
      expect(insertedMarks).not.toContain('bold')
      expect(serializeRangeToHtml(editor, 0, 0)).toContain('before new after')
      const runs: Array<{ text: string; author?: string }> = []
      editor.state.doc.descendants((node) => {
        if (node.isText)
          runs.push({
            text: node.text!,
            author: node.marks.find((mark) => mark.type.name === 'del')?.attrs.author,
          })
      })
      expect(runs).toContainEqual({ text: 'old', author: 'Reviewer' })
    },
  )

  it('refuses a selection consisting only of already deleted text', async () => {
    const editor = cellFixture([
      text('before '),
      text(TARGET, [{ type: 'del', attrs: revision('del') }]),
      text(' after'),
    ])
    select(editor, 0)
    const before = editor.state.doc
    const result = await executeTool(
      editor,
      { id: 'replace', name: 'replace_selection', input: { html: REPLACEMENT } },
      NUM_IDS,
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('pending tracked deletion')
    expect(editor.state.doc.eq(before)).toBe(true)
  })

  it('refuses a text selection spanning different cells', async () => {
    const editor = cellFixture([text(ORIGINAL)])
    const from = select(editor, 0).from
    const to = select(editor, 0, 'other cell').to
    editor.commands.setTextSelection({ from, to })
    const before = editor.state.doc
    const result = await executeTool(
      editor,
      { id: 'replace', name: 'replace_selection', input: { html: REPLACEMENT } },
      NUM_IDS,
    )
    expect(result.isError).toBe(true)
    expect(editor.state.doc.eq(before)).toBe(true)
  })

  it('does not leak withheld text or shift a selection after a withheld run', () => {
    const editor = cellFixture([
      text('secret', [{ type: 'redaction', attrs: { label: 'private' } }]),
      text(ORIGINAL),
    ])
    select(editor, 0)
    const context = buildDocContext(editor)
    expect(context).not.toContain('secret')
    expect(context).toContain('{{private}}')
    expect(context).toContain(`<sel>${TARGET}</sel>`)
    expect(context).toContain(`Selected text (${TARGET.length} characters): "${TARGET}"`)
  })

  it('keeps cell text and positions after both rowspan and orphan merge continuations', async () => {
    const table = tableModelToPmNode({
      rows: [
        [{ paras: ['anchor'], vMerge: 'restart' }, { paras: ['top'] }, { paras: ['top right'] }],
        [
          { paras: ['covered'], vMerge: 'continue' },
          { paras: ['orphan'], vMerge: 'continue' },
          { paras: [ORIGINAL] },
        ],
        [
          { paras: ['covered'], vMerge: 'continue' },
          { paras: ['bottom'] },
          { paras: ['bottom right'] },
        ],
      ],
    })
    const editor = createEditor({ type: 'doc', content: [table] })
    expect(serializeRangeToHtml(editor, 0, 0)).toContain(
      `<tr><td data-row="1" data-col="1"></td><td data-row="1" data-col="2">${ORIGINAL}</td></tr>`,
    )
    select(editor, 0)
    expect(buildDocContext(editor)).toContain(`<sel>${TARGET}</sel>`)
    const result = await executeTool(
      editor,
      { id: 'replace', name: 'replace_selection', input: { html: REPLACEMENT } },
      NUM_IDS,
    )
    expect(result.isError, result.output).not.toBe(true)
    expect(serializeRangeToHtml(editor, 0, 0)).toContain(ORIGINAL.replace(TARGET, REPLACEMENT))
    expect(serializeRangeToHtml(editor, 0, 0)).toContain('bottom right')
  })

  it('preserves ruby, formula text and note numbers in ordinary table reads', () => {
    const editor = cellFixture([
      { type: 'docRuby', attrs: { base: '漢字', ruby: 'かんじ' } },
      text(' '),
      { type: 'docInlineMath', attrs: { omml: '<m:oMath/>', text: 'E=mc2' } },
      text(' '),
      { type: 'docNoteRef', attrs: { id: '1', num: 2, kind: 'footnote' } },
      text(' '),
      text(TARGET),
    ])
    expect(serializeRangeToHtml(editor, 0, 0)).toContain(
      `<td data-row="0" data-col="0" data-doc-readonly="formula">漢字 E=mc2 2 ${TARGET}</td>`,
    )
  })
})
