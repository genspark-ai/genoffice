// The navigation pane remembers which sub-view was last used and how deep the
// outline was opened, so reopening the app lands where you left off
// (genoffice#1348). The collapsed set is deliberately NOT persisted — it is
// keyed by heading text and means nothing in another document. Neither is the
// search query: it would seed every new session, so opening any other document
// would land on Results already running the previous document's search.
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
    expect(readNavPrefs()).toEqual({ tab: 'headings', maxLevel: 9 })
  })

  it('reads back a stored preference', () => {
    store(JSON.stringify({ tab: 'pages', maxLevel: 3 }))
    expect(readNavPrefs()).toEqual({ tab: 'pages', maxLevel: 3 })
  })

  it('survives a corrupt or hand-edited value', () => {
    for (const bad of ['not json', '[]', 'null', '{"tab":', '42']) {
      store(bad)
      expect(readNavPrefs()).toEqual({ tab: 'headings', maxLevel: 9 })
    }
  })

  it('rejects an unknown sub-view instead of trusting it', () => {
    store(JSON.stringify({ tab: 'evil', maxLevel: 3 }))
    expect(readNavPrefs().tab).toBe('headings')
  })

  it('reads a stored Results tab back as Headings, since its query is not kept', () => {
    store(JSON.stringify({ tab: 'results', maxLevel: 4 }))
    expect(readNavPrefs()).toEqual({ tab: 'headings', maxLevel: 4 })
  })

  it('ignores a stored query rather than seeding a new session with it', () => {
    store(JSON.stringify({ tab: 'headings', maxLevel: 9, query: 'contract' }))
    expect(readNavPrefs()).toEqual({ tab: 'headings', maxLevel: 9 })
  })

  it('clamps a level that is out of range or not an integer', () => {
    for (const maxLevel of [0, -1, 10, 2.5, 'deep', null]) {
      store(JSON.stringify({ tab: 'headings', maxLevel }))
      expect(readNavPrefs().maxLevel).toBe(9)
    }
  })

  it('keeps a level inside the heading range', () => {
    store(JSON.stringify({ tab: 'headings', maxLevel: 1 }))
    expect(readNavPrefs().maxLevel).toBe(1)
  })
})
