import { TabStopType } from 'docx'
import { describe, expect, it } from 'vitest'
import { createRenderContext } from '../src/generate/render-context'
import { tabStopsFor } from '../src/generate/word-utils'

describe('tabStopsFor', () => {
  it('ignores a run that carries only measured metrics and no text', () => {
    const context = createRenderContext({})
    // the renderer hands over node.runs as analyzed: a nested run can report
    // only tabFrac/sizePx, and makeRuns in the same file guards with (r.text || '')
    const runs = [{ tabFrac: 0.4, sizePx: 12 }, { text: 'a\tb' }]
    expect(() => tabStopsFor(context, runs)).not.toThrow()
    expect(tabStopsFor(context, runs)).toHaveLength(1)
  })

  it('never emits a negative w:pos for a tab measured left of the paragraph start', () => {
    const context = createRenderContext({})
    const stops = tabStopsFor(context, [{ text: '\t', tabFrac: -0.2 }])
    expect(stops).toEqual([{ type: TabStopType.LEFT, position: 0 }])
  })

  it('returns no stops for a line without tabs', () => {
    const context = createRenderContext({})
    expect(tabStopsFor(context, [{ text: 'plain' }])).toEqual([])
  })
})
