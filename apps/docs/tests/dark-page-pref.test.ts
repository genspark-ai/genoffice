import { beforeEach, describe, expect, it } from 'vitest'
import { DARK_PAGE_KEY, readDarkPagePref, writeDarkPagePref } from '../src/renderer/dark-page-pref'

describe('dark page preference', () => {
  beforeEach(() => {
    localStorage.removeItem(DARK_PAGE_KEY)
  })

  it('is unset until the user makes an explicit choice', () => {
    expect(readDarkPagePref()).toBeNull()
  })

  it('round-trips an explicit off and on', () => {
    writeDarkPagePref(false)
    expect(readDarkPagePref()).toBe(false)
    writeDarkPagePref(true)
    expect(readDarkPagePref()).toBe(true)
  })

  it('ignores values it did not write', () => {
    localStorage.setItem(DARK_PAGE_KEY, 'yes')
    expect(readDarkPagePref()).toBeNull()
  })
})
