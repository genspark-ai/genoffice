/** Word's Bullets and Numbering dialog: each tab offers a library the list commands can apply. */
import { beforeEach, describe, expect, it } from 'vitest'
import { LIST_TABS } from '../src/renderer/components/BulletsNumberingDialog'
import {
  bulletPresetLevels,
  customizeLibraryCard,
  libraryCardCustomized,
  listLibrary,
  resetLibraryCard,
} from '../src/renderer/list-presets'

describe('bullets and numbering dialog', () => {
  beforeEach(() => localStorage.clear())

  it('has a bulleted, a numbered and an outline tab, each with its own Customize dialog', () => {
    expect(LIST_TABS.map((tab) => [tab.kind, tab.customize])).toEqual([
      ['bullets', 'bullet'],
      ['numbers', 'number'],
      ['multi', 'multi'],
    ])
  })

  it('offers presets that define all nine levels, bullets on the bulleted tab only', () => {
    for (const tab of LIST_TABS) {
      const presets = listLibrary(tab.kind)
      expect(presets.length).toBeGreaterThan(0)
      for (const levels of presets) expect(levels).toHaveLength(9)
    }
    expect(listLibrary('bullets').every((levels) => levels[0].numFmt === 'bullet')).toBe(true)
    expect(listLibrary('numbers').some((levels) => levels[0].numFmt === 'bullet')).toBe(false)
  })

  it('keeps a customized card until Reset puts the built-in one back', () => {
    const builtIn = listLibrary('bullets')
    const star = bulletPresetLevels('\u2605', { color: 'FF0000' })
    expect(libraryCardCustomized('bullets', 2)).toBe(false)
    customizeLibraryCard('bullets', 2, star)
    expect(libraryCardCustomized('bullets', 2)).toBe(true)
    expect(listLibrary('bullets')[2]).toEqual(star)
    expect(listLibrary('bullets')[1]).toEqual(builtIn[1])
    expect(listLibrary('numbers')[2][0].numFmt).toBe('upperRoman')
    resetLibraryCard('bullets', 2)
    expect(libraryCardCustomized('bullets', 2)).toBe(false)
    expect(listLibrary('bullets')).toEqual(builtIn)
  })

  it('does not count a card customized back to its built-in definition, nor a card that does not exist', () => {
    customizeLibraryCard('multi', 0, listLibrary('numbers')[0])
    customizeLibraryCard('multi', 0, JSON.parse(JSON.stringify(listLibrary('multi')[1])))
    expect(libraryCardCustomized('multi', 0)).toBe(true)
    resetLibraryCard('multi', 0)
    const builtIn = listLibrary('multi')[0]
    customizeLibraryCard('multi', 0, JSON.parse(JSON.stringify(builtIn)))
    expect(libraryCardCustomized('multi', 0)).toBe(false)
    const numbered = listLibrary('numbers')[0]
    customizeLibraryCard(
      'numbers',
      0,
      numbered.map((l, i) => (i === 0 ? { ...l, start: 1 } : l)),
    )
    expect(libraryCardCustomized('numbers', 0)).toBe(false)
    customizeLibraryCard('multi', 99, builtIn)
    expect(listLibrary('multi')).toHaveLength(8)
    localStorage.setItem('aidocs.list.library.bullets', 'not json')
    expect(listLibrary('bullets')[0][0].lvlText).toBe('\u2022')
  })
})
