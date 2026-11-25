import { beforeEach, describe, expect, it } from 'vitest'
import {
  STROKE_WIDTHS,
  loadDrawColor,
  loadDrawWidth,
  saveDrawColor,
  saveDrawWidth,
  type DrawRgb,
} from '../src/renderer/draw-prefs'
import { STROKE_WIDTH } from '../src/renderer/view-config'

const RED: DrawRgb = [0.86, 0.22, 0.18]

describe('pen color and thickness kept across sessions', () => {
  beforeEach(() => localStorage.clear())

  it('falls back to the defaults when nothing is stored', () => {
    expect(loadDrawColor(RED)).toEqual(RED)
    expect(loadDrawWidth()).toBe(STROKE_WIDTH)
    expect(STROKE_WIDTHS).toContain(STROKE_WIDTH)
  })

  it('returns what was saved', () => {
    saveDrawColor([0.1, 0.5, 1])
    saveDrawWidth(5)
    expect(loadDrawColor(RED)).toEqual([0.1, 0.5, 1])
    expect(loadDrawWidth()).toBe(5)
  })

  it('ignores a malformed or out-of-range entry', () => {
    localStorage.setItem('genoffice-pdf-draw-color', '[2,0,0]')
    localStorage.setItem('genoffice-pdf-draw-width', '400')
    expect(loadDrawColor(RED)).toEqual(RED)
    expect(loadDrawWidth()).toBe(STROKE_WIDTH)
    localStorage.setItem('genoffice-pdf-draw-color', '{oops')
    expect(loadDrawColor(RED)).toEqual(RED)
  })
})
