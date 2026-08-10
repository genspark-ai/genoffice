import { Editor } from '@tiptap/core'
import { parseDocx, patchChartPartXml, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildChartDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function chartEditor() {
  const source = await buildChartDocx()
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { source, parsed, editor }
}

describe('embedded chart editing', () => {
  it('renders an SVG preview and a data grid for the chart block', async () => {
    const { editor } = await chartEditor()
    const chart = editor.view.dom.querySelector('.doc-protected-chart') as HTMLElement
    expect(chart).toBeTruthy()
    expect(chart.querySelector('.doc-chart-title')?.textContent).toBe('销售统计')
    expect(chart.querySelectorAll('.doc-chart-svg rect').length).toBe(7) // 6 slots - 1 gap + 2 legend swatches
    expect(chart.querySelectorAll('.doc-chart-cat').length).toBe(3)
    expect(chart.querySelectorAll('.doc-chart-val').length).toBe(5)
    expect(chart.querySelectorAll('.doc-chart-gap').length).toBe(1)
    editor.destroy()
  })

  it('double-click edits data and round-trips through a chart part patch', async () => {
    const { parsed, editor } = await chartEditor()
    const chart = editor.view.dom.querySelector('.doc-protected-chart') as HTMLElement
    const firstValue = chart.querySelector('.doc-chart-val') as HTMLElement
    expect(firstValue.getAttribute('contenteditable')).toBe('false')

    firstValue.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
    firstValue.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
    expect(firstValue.getAttribute('contenteditable')).toBe('true')
    firstValue.textContent = '150.5'
    const title = chart.querySelector('.doc-chart-title') as HTMLElement
    title.textContent = 'Annual <Sales> Report'
    const name = chart.querySelector('.doc-chart-name') as HTMLElement
    name.textContent = 'East Region 1'
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.chartPatches).toHaveLength(1)
    expect(plan.chartPatches[0]).toEqual({
      partPath: 'word/charts/chart1.xml',
      patch: {
        title: 'Annual <Sales> Report',
        series: [{ name: 'East Region 1', values: [150.5, null, null] }, null],
      },
    })

    const partXml = {
      'word/charts/chart1.xml': patchChartPartXml(
        parsed.extras.chartParts['word/charts/chart1.xml'],
        plan.chartPatches[0].patch,
      ),
    }
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks, { partXml }))
    const block = reparsed.blocks.find((b) => b.label === 'Chart')
    expect(block?.chartDisplay?.title).toBe('Annual <Sales> Report')
    expect(block?.chartDisplay?.series[0]).toEqual({
      name: 'East Region 1',
      values: [150.5, 88.5, 96],
    })
    expect(block?.chartDisplay?.series[1]).toEqual({ name: '华南', values: [70, null, 110] })
    editor.destroy()
  })

  it('keeps an untouched chart byte-stable', async () => {
    const { source, parsed, editor } = await chartEditor()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.chartPatches).toHaveLength(0)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
  })

  it('unparseable numbers keep the original cached value', async () => {
    const { parsed, editor } = await chartEditor()
    const chart = editor.view.dom.querySelector('.doc-protected-chart') as HTMLElement
    const firstValue = chart.querySelector('.doc-chart-val') as HTMLElement
    firstValue.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))
    firstValue.textContent = 'abc'
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.chartPatches).toHaveLength(0)
    editor.destroy()
  })
})
