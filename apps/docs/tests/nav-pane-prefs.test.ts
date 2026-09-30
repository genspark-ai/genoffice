// The navigation pane remembers which sub-view was last used, how deep the
// outline was opened, and the last search, so reopening the app lands where you
// left off (#1348). The collapsed set is deliberately NOT persisted — it is
// keyed by heading text and means nothing in another document.
import { beforeEach, describe, expect, it } from 'vitest'
import { readNavPrefs } from '../src/renderer/components/NavPane'

const KEY = 'aidocs.navPrefs'

const store = (value: string | null) => {
  if (value === null) localStorage.removeItem(KEY)
  else localStorage.setItem(KEY, value)
}

describe('readNavPrefs', () => {
  beforeEach(() => {
    localStorage.clear()
  })

  it('falls back to a blank outline when nothing is stored', () => {
    expect(readNavPrefs()).toEqual({ tab: 'headings', maxLevel: 9, query: '' })
  })

  it('reads back a stored preference', () => {
    store(JSON.stringify({ tab: 'pages', maxLevel: 3, query: 'contract' }))
    expect(readNavPrefs()).toEqual({ tab: 'pages', maxLevel: 3, query: 'contract' })
  })

  it('survives a corrupt or hand-edited value', () => {
    for (const bad of ['not json', '[]', 'null', '{"tab":', '42']) {
      store(bad)
      expect(readNavPrefs()).toEqual({ tab: 'headings', maxLevel: 9, query: '' })
    }
  })

  it('rejects an unknown sub-view instead of trusting it', () => {
    store(JSON.stringify({ tab: 'evil', maxLevel: 3, query: '' }))
    expect(readNavPrefs().tab).toBe('headings')
  })

  it('clamps a level that is out of range or not an integer', () => {
    for (const maxLevel of [0, -1, 10, 2.5, 'deep', null]) {
      store(JSON.stringify({ tab: 'headings', maxLevel, query: '' }))
      expect(readNavPrefs().maxLevel).toBe(9)
    }
  })

  it('keeps a level inside the heading range', () => {
    store(JSON.stringify({ tab: 'headings', maxLevel: 1, query: '' }))
    expect(readNavPrefs().maxLevel).toBe(1)
  })

  it('drops a non-string or unbounded query rather than storing it back', () => {
    store(JSON.stringify({ tab: 'headings', maxLevel: 9, query: { evil: true } }))
    expect(readNavPrefs().query).toBe('')
    store(JSON.stringify({ tab: 'headings', maxLevel: 9, query: 'x'.repeat(5000) }))
    expect(readNavPrefs().query).toHaveLength(200)
  })
})
