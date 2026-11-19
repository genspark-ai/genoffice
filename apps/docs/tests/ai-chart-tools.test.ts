import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildChartDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { executeTool } from '../src/renderer/ai/tools'

const NUM_IDS = { bullet: null, ordered: null }

function blankEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null },
          content: [{ type: 'text', text: 'Body paragraph' }],
        },
      ],
    },
  })
}

async function nativeChartEditor() {
  const parsed = await parseDocx(await buildChartDocx())
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { parsed, editor }
}

const call = (name: string, input: Record<string, unknown>) => ({ id: 't', name, input })

describe('insert_chart tool', () => {
  it('inserts a gen chart block with normalized values', async () => {
    const editor = blankEditor()
    const exec = await executeTool(
      editor,
      call('insert_chart', {
        kind: 'bar',
        title: 'Quarterly Sales',
        categories: ['Q1', 'Q2', 'Q3'],
        series: [{ name: 'Sales', values: [10, '20', 'oops'] }],
        afterBlockIndex: 0,
      }),
      NUM_IDS,
    )
    expect(exec.isError).toBeUndefined()
    expect(exec.mutated).toBe(true)
    const node = editor.state.doc.child(1)
    expect(node.type.name).toBe('docProtected')
    expect(node.attrs.blockType).toBe('chart')
    const gen = node.attrs.genChart as {
      kind: string
      title: string
      series: Array<{ values: unknown[] }>
    }
    expect(gen.kind).toBe('bar')
    expect(gen.title).toBe('Quarterly Sales')
    expect(gen.series[0]!.values).toEqual([10, 20, null]) // numeric strings normalized; non-numeric → null
    expect(node.attrs.chartDisplay).toBeTruthy()
    editor.destroy()
  })

  it('rejects unsupported kind and empty data', async () => {
    const editor = blankEditor()
    const bad = await executeTool(
      editor,
      call('insert_chart', { kind: 'radar', categories: ['a'], series: [{ values: [1] }] }),
      NUM_IDS,
    )
    expect(bad.isError).toBe(true)
    const empty = await executeTool(
      editor,
      call('insert_chart', { kind: 'pie', categories: [], series: [] }),
      NUM_IDS,
    )
    expect(empty.isError).toBe(true)
    editor.destroy()
  })
})

describe('edit_chart tool', () => {
  it('edits a native chart and produces a save-time chart patch', async () => {
    const { parsed, editor } = await nativeChartEditor()
    const exec = await executeTool(
      editor,
      call('edit_chart', {
        blockIndex: 0,
        title: 'Annual Sales',
        series: [{ index: 0, name: 'East Region 1', values: [150.5, null, null] }],
      }),
      NUM_IDS,
    )
    expect(exec.isError).toBeUndefined()
    expect(exec.mutated).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.chartPatches).toHaveLength(1)
    const patch = plan.chartPatches[0]!.patch
    expect(patch.title).toBe('Annual Sales')
    expect(patch.series?.[0]?.name).toBe('East Region 1')
    expect(patch.series?.[0]?.values?.[0]).toBe(150.5)
    editor.destroy()
  })

  it('rejects writing a gap point of a native chart (unpatchable)', async () => {
    const { editor } = await nativeChartEditor()
    // The 2nd point of the South China series is null in the original chart (no c:pt to patch)
    const exec = await executeTool(
      editor,
      call('edit_chart', { blockIndex: 0, series: [{ index: 1, values: [null, 99, null] }] }),
      NUM_IDS,
    )
    expect(exec.isError).toBe(true)
    expect(exec.output).toContain('empty points of a native chart cannot be written')
    editor.destroy()
  })

  it('rejects length mismatch and non-chart blocks', async () => {
    const { editor } = await nativeChartEditor()
    const mismatch = await executeTool(
      editor,
      call('edit_chart', { blockIndex: 0, categories: ['January'] }),
      NUM_IDS,
    )
    expect(mismatch.isError).toBe(true)
    const editor2 = blankEditor()
    const notChart = await executeTool(
      editor2,
      call('edit_chart', { blockIndex: 0, title: 'x' }),
      NUM_IDS,
    )
    expect(notChart.isError).toBe(true)
    editor.destroy()
    editor2.destroy()
  })

  it('edits a freshly inserted gen chart keeping genChart spec in sync', async () => {
    const editor = blankEditor()
    await executeTool(
      editor,
      call('insert_chart', {
        kind: 'line',
        title: 'Trend',
        categories: ['Jan', 'Feb'],
        series: [{ name: 's1', values: [1, 2] }],
        afterBlockIndex: 0,
      }),
      NUM_IDS,
    )
    const exec = await executeTool(
      editor,
      call('edit_chart', {
        blockIndex: 1,
        categories: ['January', 'February'],
        series: [{ index: 0, values: [3, null] }],
      }),
      NUM_IDS,
    )
    expect(exec.isError).toBeUndefined()
    const gen = editor.state.doc.child(1).attrs.genChart as {
      categories: string[]
      series: Array<{ values: unknown[] }>
    }
    expect(gen.categories).toEqual(['January', 'February'])
    expect(gen.series[0]!.values).toEqual([3, 2]) // null = unchanged, keep original value
    editor.destroy()
  })
})
