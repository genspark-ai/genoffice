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

  it('returns no stops for a line without tabs', () => {
    const context = createRenderContext({})
    expect(tabStopsFor(context, [{ text: 'plain' }])).toEqual([])
  })

  it('clamps a negative tabFrac to a non-negative stop position', () => {
    const context = createRenderContext({})
    // a measured fraction can arrive negative (an offset column computed before
    // the left margin was known); w:pos is a signed twips measure, so the raw
    // product emitted a tab stop before the text margin that Word ignores
    const stops = tabStopsFor(context, [{ text: 'a\tb', tabFrac: -0.25 }])
    expect(stops).toHaveLength(1)
    expect(stops[0]!.position).toBe(0)
  })

  it('keeps the measured position of an in-range tabFrac and the right stop past 0.75', () => {
    const context = createRenderContext({})
    expect(tabStopsFor(context, [{ text: 'a\tb', tabFrac: 0.5 }])[0]!.position).toBe(
      Math.round(0.5 * context.contentDxa),
    )
    expect(tabStopsFor(context, [{ text: 'a\tb', tabFrac: 0.9 }])[0]!.position).toBe(
      context.contentDxa,
    )
  })
})
