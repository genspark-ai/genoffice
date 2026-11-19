import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx, type ChartDisplay, type NewChart } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const SPEC: NewChart = {
  kind: 'bar',
  title: 'Sales',
  categories: ['January', 'February'],
  series: [{ name: 'East', values: [10, 20] }],
}

describe('chart insertion from the editor', () => {
  it('genChart node saves as a real chart part and survives reparse', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const display: ChartDisplay = {
      partPath: '',
      kind: SPEC.kind,
      title: SPEC.title,
      categories: SPEC.categories,
      series: SPEC.series,
    }
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'chart',
          label: 'Chart',
          genChart: SPEC,
          chartDisplay: display,
        },
      })
      .run()
    // the freshly inserted chart renders its SVG immediately
    expect(editor.view.dom.querySelector('.doc-chart-svg, svg')).toBeTruthy()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.saveBlocks.some((b) => b.kind === 'chart')).toBe(true)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const chart = reparsed.blocks.find((b) => b.chartDisplay)
    expect(chart?.chartDisplay?.title).toBe('Sales')
    expect(chart?.chartDisplay?.categories).toEqual(['January', 'February'])
    expect(chart?.chartDisplay?.partPath).toBe('word/charts/chart1.xml')
    editor.destroy()
  })

  it('in-place data edits (chartDisplay attr) flow into the saved part', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editor
      .chain()
      .insertContentAt(editor.state.doc.content.size, {
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'chart',
          label: 'Chart',
          genChart: SPEC,
          chartDisplay: {
            partPath: '',
            kind: 'bar',
            title: 'Edited title',
            categories: ['January', 'February'],
            series: [{ name: 'East', values: [99, 20] }],
          },
        },
      })
      .run()
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(saved)
    const chart = reparsed.blocks.find((b) => b.chartDisplay)
    expect(chart?.chartDisplay?.title).toBe('Edited title')
    expect(chart?.chartDisplay?.series[0].values).toEqual([99, 20])
    editor.destroy()
  })
})
