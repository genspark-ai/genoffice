import { beforeEach, describe, expect, it } from 'vitest'
import { getRecentColors, pushRecentColor } from '../src/renderer/recent-colors'

beforeEach(() => {
  localStorage.clear()
})

describe('recent colors', () => {
  it('stores valid hex as uppercase RRGGBB', () => {
    pushRecentColor('#ff0000')
    expect(getRecentColors()).toEqual(['#FF0000'])
  })

  it('rejects 8-digit RGBA and trailing garbage instead of truncating', () => {
    pushRecentColor('#FF000080')
    pushRecentColor('abcdefXYZ')
    expect(getRecentColors()).toEqual([])
  })

  it('filters legacy junk entries on read', () => {
    localStorage.setItem('slides:recent-colors', JSON.stringify(['#FF000080', '#00ff00']))
    expect(getRecentColors()).toEqual(['#00FF00'])
  })

  it('caps and dedupes oversized stores on read', () => {
    const stored = Array.from({ length: 12 }, (_, i) => `#${i.toString(16).padStart(6, '0')}`)
    localStorage.setItem('slides:recent-colors', JSON.stringify([...stored, stored[0], stored[1]]))
    const colors = getRecentColors()
    expect(colors).toHaveLength(10)
    expect(new Set(colors).size).toBe(10)
    expect(colors[0]).toBe(stored[0])
  })
})

describe('recent color session (gendesk #279)', () => {
  it('a native-picker drag records only the final color, on close', async () => {
    const { createRecentColorSession } = await import('../src/renderer/recent-colors')
    pushRecentColor('#00FF00') // the user's earlier green must survive the drag
    const session = createRecentColorSession()
    // macOS panel drag: one change event per movement across the blues
    for (const hex of ['#000011', '#000033', '#000066', '#0000aa', '#0000ff']) {
      session.preview(hex)
    }
    expect(getRecentColors()).toEqual(['#00FF00']) // nothing recorded mid-drag
    session.close()
    expect(getRecentColors()).toEqual(['#0000FF', '#00FF00'])
    session.close() // closing twice never double-records
    expect(getRecentColors()).toEqual(['#0000FF', '#00FF00'])
  })

  it('a swatch click records immediately and drops the staged preview', async () => {
    const { createRecentColorSession } = await import('../src/renderer/recent-colors')
    const session = createRecentColorSession()
    session.preview('#123456') // transient drag the user abandoned
    session.commit('#C00000') // then a deliberate standard-color click
    session.close()
    expect(getRecentColors()).toEqual(['#C00000'])
  })

  it('closing with no preview records nothing', async () => {
    const { createRecentColorSession } = await import('../src/renderer/recent-colors')
    createRecentColorSession().close()
    expect(getRecentColors()).toEqual([])
  })
})
