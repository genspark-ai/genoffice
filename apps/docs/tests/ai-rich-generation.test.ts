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
import { executeTool } from '../src/renderer/ai/tools'

/**
 * Guard tests for AI rich generation (tables / code blocks / quotes / TOC):
 * 1. blank template + restricted HTML with the new tags saves to a docx that
 *    reparses with the expected block structure;
 * 2. reopen + save with no edits returns byte-identical output (the signature
 *    round-trip invariant that keeps untouched blocks byte-stable);
 * 3. reopened AI tables stay protected and cell-text patching works.
 */

const NUM_IDS = { bullet: BLANK_BULLET_NUM_ID, ordered: BLANK_ORDERED_NUM_ID }

async function loadEditorModule() {
  // extensions import react components indirectly; keep it lazy per test run
  return import('../src/renderer/editor/extensions')
}

async function createBlankEditor() {
  const { editorExtensions } = await loadEditorModule()
  const bytes = await buildBlankDocx()
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { editor, parsed }
}

const RICH_HTML = [
  '<h1>Annual Report</h1>',
  '<p>Overview paragraph.</p>',
  '<table><tr><th>Metric</th><th>Value</th></tr><tr><td>Revenue</td><td>100</td></tr><tr><td>Cost</td><td>60<br>(incl. tax)</td></tr></table>',
  '<pre>const x = 1\nconsole.log(x)</pre>',
  '<blockquote>A quoted sentence.</blockquote>',
  '<ul><li>Point one</li><li>Point two</li></ul>',
].join('\n')

async function generateRichDocx(): Promise<Uint8Array> {
  const { editor, parsed } = await createBlankEditor()
  const exec = await executeTool(
    editor,
    { id: 't', name: 'insert_content', input: { html: RICH_HTML } },
    NUM_IDS,
  )
  expect(exec.isError).toBeFalsy()
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const saved = await saveDocx(parsed, plan.saveBlocks)
  editor.destroy()
  return saved
}

describe('AI rich generation: blank template -> docx', () => {
  it('table/code block/blockquote/list fully generated then reparsed with correct structure', async () => {
    const saved = await generateRichDocx()
    const reparsed = await parseDocx(saved)
    const blocks = reparsed.blocks.filter((b) => !b.hidden)

    const heading = blocks.find((b) => b.type === 'heading')
    expect(heading?.level).toBe(1)
    expect(heading?.runs?.[0]?.text).toBe('Annual Report')

    const table = blocks.find((b) => b.type === 'table')
    expect(table?.table).toBeDefined()
    const model = table!.table!
    expect(model.rows.length).toBe(3)
    // header texts arrive bold with shading (patchTableCellTexts reuses the
    // generated header run's rPr)
    expect(model.rows[0][0].paras).toEqual(['Metric'])
    expect(model.rows[0][0].bold).toBe(true)
    expect(model.rows[0][0].fill).toBe('F2F2F2')
    expect(model.rows[1][1].paras).toEqual(['100'])
    expect(model.rows[2][1].paras).toEqual(['60', '(incl. tax)'])

    const code = blocks.find(
      (b) => b.format?.shadingFill === 'F2F2F2' && b.format?.borders === 'tblr',
    )
    expect(code?.type).toBe('paragraph')
    expect(code?.runs?.[0]?.font).toBe('Consolas')
    expect(code?.runs?.map((r) => r.text).join('')).toBe('const x = 1\nconsole.log(x)')

    const quote = blocks.find((b) => b.format?.borders === 'l')
    expect(quote?.type).toBe('paragraph')
    expect(quote?.format?.indentLeft).toBe(720)
    expect(quote?.runs?.[0]?.color).toBe('666666')
    expect(quote?.runs?.[0]?.text).toBe('A quoted sentence.')

    const listItems = blocks.filter((b) => b.type === 'listItem')
    expect(listItems.length).toBe(2)
    expect(listItems[0].list?.numId).toBe(BLANK_BULLET_NUM_ID)
  })

  it('AI-inserted tables carry an equal column grid (a colgroup pins fixed layout against pagination widgets)', async () => {
    const { editor } = await createBlankEditor()
    const exec = await executeTool(
      editor,
      { id: 't', name: 'insert_content', input: { html: RICH_HTML } },
      NUM_IDS,
    )
    expect(exec.isError).toBeFalsy()
    let tableNode: PmNode | undefined
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docTable') tableNode = node.toJSON() as PmNode
      return !tableNode
    })
    expect(tableNode?.attrs?.colWidthsPct).toEqual([50, 50])
    // the rendered DOM gets a colgroup: without it a page-gap row's spanning cell
    // would widen the fixed-layout grid and collapse every column (flicker loop)
    const cols = editor.view.dom.querySelectorAll('table.doc-table > colgroup > col')
    expect(cols.length).toBe(2)
    editor.destroy()
  })

  it('read_blocks serializes tables/code blocks/blockquotes back to restricted HTML', async () => {
    const saved = await generateRichDocx()
    const reparsed = await parseDocx(saved)
    const { editorExtensions } = await loadEditorModule()
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(reparsed.blocks) as never)
    const exec = await executeTool(
      editor,
      {
        id: 't',
        name: 'read_blocks',
        input: { startBlockIndex: 0, endBlockIndex: editor.state.doc.childCount - 1 },
      },
      NUM_IDS,
    )
    expect(exec.output).toContain('<table><tr><th>Metric</th><th>Value</th></tr>')
    expect(exec.output).toContain('<td>60<br>(incl. tax)</td>')
    expect(exec.output).toContain('<pre>const x = 1\nconsole.log(x)</pre>')
    expect(exec.output).toContain('<blockquote>A quoted sentence.</blockquote>')
    expect(exec.output).not.toContain('Protected content')
    editor.destroy()
  })
})

describe('AI rich generation: byte round-trip invariant', () => {
  it('reopen then save without edits = identical bytes (signature round-trip holds)', async () => {
    const saved = await generateRichDocx()
    const reparsed = await parseDocx(saved)
    const { editorExtensions } = await loadEditorModule()
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(reparsed.blocks) as never)

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, reparsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(plan.deletedCount).toBe(0)
    const saved2 = await saveDocx(reparsed, plan.saveBlocks)
    expect(Buffer.from(saved2).equals(Buffer.from(saved))).toBe(true)
    editor.destroy()
  })
})

describe('AI table: native editing after reopen + cell patches', () => {
  it('reopened table is a native node; editing cell text uses a targeted patch and other cells stay untouched', async () => {
    const saved = await generateRichDocx()
    const reparsed = await parseDocx(saved)
    const { editorExtensions } = await loadEditorModule()
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(reparsed.blocks) as never)

    let cellPos = -1
    let foundTable = false
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docTable') foundTable = true
      if (cellPos === -1 && node.type.name === 'docTableCell') cellPos = pos
    })
    expect(foundTable).toBe(true)
    expect(cellPos).toBeGreaterThanOrEqual(0)
    const cellTextLength = editor.state.doc.nodeAt(cellPos)!.textContent.length
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, cellPos + 2, cellPos + 2 + cellTextLength),
      ),
    )
    editor.view.dispatch(editor.state.tr.insertText('Net revenue'))

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, reparsed.blocks)
    expect(plan.changedCount).toBe(1)
    const saved2 = await saveDocx(reparsed, plan.saveBlocks)
    const final = await parseDocx(saved2)
    const table = final.blocks.find((b) => b.type === 'table')!.table!
    expect(table.rows[1][0].paras).toEqual(['Net revenue'])
    // untouched cells keep text and header formatting
    expect(table.rows[1][1].paras).toEqual(['100'])
    expect(table.rows[0][0].bold).toBe(true)
    expect(table.rows[0][0].fill).toBe('F2F2F2')
    editor.destroy()
  })
})

describe('insertToc command', () => {
  it('generates a real TOC field; blank template ships TOC1-9 styles; round-trip bytes stable', async () => {
    const { editor, parsed } = await createBlankEditor()
    await executeTool(
      editor,
      {
        id: 't1',
        name: 'insert_content',
        input: { html: '<h1>Chapter 1</h1><p>Content.</p><h2>Subsection</h2><p>Content.</p>' },
      },
      NUM_IDS,
    )
    const exec = await executeTool(
      editor,
      {
        id: 't2',
        name: 'apply_commands',
        input: { commands: [{ insertToc: { afterBlockIndex: -1 } }] },
      },
      NUM_IDS,
    )
    expect(exec.isError).toBeFalsy()
    expect(exec.output).toContain('目录')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    editor.destroy()

    // real field with dirty begin + TOC styles present in the package
    expect(reparsed.internal.documentXml).toContain('w:fldCharType="begin" w:dirty="true"')
    expect(reparsed.internal.documentXml).toContain('<w:instrText xml:space="preserve"> TOC')
    expect(reparsed.internal.documentXml).toContain('<w:pStyle w:val="TOC1"/>')
    expect(reparsed.internal.documentXml).toContain('<w:pStyle w:val="TOC2"/>')
    expect(reparsed.styles.get('TOC1')).toBeDefined()
    expect(reparsed.styles.get('TOC9')).toBeDefined()
    // TOC lines reopen as protected blocks rendered as clickable TOC entries
    const tocLines = reparsed.blocks.filter(
      (b) => b.type === 'passthrough' && b.fieldDisplay?.kind === 'tocLine',
    )
    expect(tocLines.length).toBe(2)
    expect(tocLines[0].fieldDisplay?.left).toBe('Chapter 1')
    expect(tocLines[1].fieldDisplay?.left).toBe('Subsection')

    // byte round trip on the TOC document too
    const { editorExtensions } = await loadEditorModule()
    const editor2 = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor2.commands.setContent(blocksToPmDoc(reparsed.blocks) as never)
    const plan2 = pmDocToSavePlan(editor2.getJSON() as PmNode, reparsed.blocks)
    expect(plan2.changedCount).toBe(0)
    expect(plan2.deletedCount).toBe(0)
    const saved2 = await saveDocx(reparsed, plan2.saveBlocks)
    expect(Buffer.from(saved2).equals(Buffer.from(saved))).toBe(true)
    editor2.destroy()
  })

  it('insertToc fails and the document is unchanged when there are no headings', async () => {
    const { editor } = await createBlankEditor()
    await executeTool(
      editor,
      { id: 't1', name: 'insert_content', input: { html: '<p>Body text only.</p>' } },
      NUM_IDS,
    )
    const before = JSON.stringify(editor.getJSON())
    const exec = await executeTool(
      editor,
      {
        id: 't2',
        name: 'apply_commands',
        input: { commands: [{ insertToc: { afterBlockIndex: -1 } }] },
      },
      NUM_IDS,
    )
    expect(exec.isError).toBe(true)
    expect(JSON.stringify(editor.getJSON())).toBe(before)
    editor.destroy()
  })
})
