import { describe, expect, it } from 'vitest'
import { tabStopsFor } from '../src/generate/word-utils'

describe('tabStopsFor (#1384)', () => {
  const context = { contentDxa: 9000 }

  it('tolerates runs without text', () => {
    const stops = tabStopsFor(context, [{ sizePx: 14 }, { text: 'a\tb', tabFrac: 0.5 }])
    expect(stops).toHaveLength(1)
    expect(stops[0]).toMatchObject({ position: 4500 })
  })

  it('returns no stops when no run carries a tab', () => {
    expect(tabStopsFor(context, [{ sizePx: 14 }])).toEqual([])
    expect(tabStopsFor(context, [{ text: 'plain' }])).toEqual([])
  })
})
