import { describe, expect, it } from 'vitest'
import { currentAfterRemoval, groupSections, setAllCollapsed } from '../src/renderer/section-groups'

const sections = [
  { id: 'A', name: 'Intro', slideIndices: [1, 2] },
  { id: 'B', name: 'Body', slideIndices: [] },
  { id: 'C', name: 'Close', slideIndices: [4, 3] },
]

describe('groupSections', () => {
  it('splits the rail positionally, with a lead group for unsectioned slides', () => {
    expect(groupSections(sections, 6)).toEqual([
      { id: null, name: '', start: 0, end: 1 },
      { id: 'A', name: 'Intro', start: 1, end: 3 },
      { id: 'B', name: 'Body', start: 3, end: 3 },
      { id: 'C', name: 'Close', start: 3, end: 6 },
    ])
  })

  it('has no lead group when the first section starts on slide 0, and none at all without sections', () => {
    const groups = groupSections([{ id: 'A', name: 'All', slideIndices: [0] }], 3)
    expect(groups).toEqual([{ id: 'A', name: 'All', start: 0, end: 3 }])
    expect(groupSections([], 3)).toBeNull()
    expect(groupSections(sections, 0)).toBeNull()
  })
})

describe('setAllCollapsed', () => {
  it('collapses every section or none', () => {
    expect([...setAllCollapsed(sections, true)]).toEqual(['A', 'B', 'C'])
    expect(setAllCollapsed(sections, false).size).toBe(0)
  })
})

describe('currentAfterRemoval', () => {
  // removing slides 2-4 of 8 leaves 5
  const group = { id: 'A', name: 'A', start: 2, end: 5 }

  it('keeps a slide before the removed range', () => {
    expect(currentAfterRemoval(1, group, 5)).toBe(1)
  })

  it('lands on the slide that now fills the hole when the current slide was removed', () => {
    expect(currentAfterRemoval(3, group, 5)).toBe(2)
    // removed range was the tail: clamp to the new last slide
    expect(currentAfterRemoval(6, { id: 'B', name: 'B', start: 5, end: 8 }, 5)).toBe(4)
  })

  it('shifts a slide after the range back by the removed count', () => {
    expect(currentAfterRemoval(7, group, 5)).toBe(4)
  })
})
