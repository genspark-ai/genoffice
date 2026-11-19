import { beforeEach, describe, expect, it } from 'vitest'
import {
  defaultNewSlideSize,
  emuToLength,
  lengthToEmu,
  matchSlideSizePreset,
  presetSize,
  saveDefaultNewSlideSize,
} from '../src/renderer/slide-size'

describe('matchSlideSizePreset', () => {
  it('recognises each preset in both orientations', () => {
    expect(matchSlideSizePreset({ cx: 12192000, cy: 6858000 })).toBe('16:9')
    expect(matchSlideSizePreset({ cx: 6858000, cy: 12192000 })).toBe('16:9')
    expect(matchSlideSizePreset({ cx: 9144000, cy: 6858000 })).toBe('4:3')
    expect(matchSlideSizePreset(presetSize('A4', true))).toBe('A4')
    expect(matchSlideSizePreset(presetSize('A3', false))).toBe('A3')
  })

  it('tells A4 from A3 although they share an aspect ratio', () => {
    expect(matchSlideSizePreset({ cx: 10692000, cy: 7560000 })).toBe('A4')
    expect(matchSlideSizePreset({ cx: 15120000, cy: 10692000 })).toBe('A3')
  })

  it('leaves other sizes custom, including PowerPoint-style "A4 paper" slides', () => {
    expect(matchSlideSizePreset({ cx: 9906000, cy: 6858000 })).toBeNull()
    expect(matchSlideSizePreset({ cx: 30240000, cy: 42768000 })).toBeNull()
  })
})

describe('presetSize', () => {
  it('swaps the sides for portrait', () => {
    expect(presetSize('A4', false)).toEqual({ cx: 10692000, cy: 7560000 })
    expect(presetSize('A4', true)).toEqual({ cx: 7560000, cy: 10692000 })
  })
})

describe('lengthToEmu / emuToLength', () => {
  it('converts typed lengths in each unit', () => {
    expect(lengthToEmu('210', 'mm')).toBe(7560000)
    expect(lengthToEmu('29,7', 'cm')).toBe(10692000)
    expect(lengthToEmu('13.333', 'in')).toBe(Math.round(13.333 * 914400))
  })

  it('rejects text, empty input and sizes outside 1–56 inches', () => {
    expect(lengthToEmu('', 'cm')).toBeNull()
    expect(lengthToEmu('abc', 'cm')).toBeNull()
    expect(lengthToEmu('0.5', 'in')).toBeNull()
    expect(lengthToEmu('57', 'in')).toBeNull()
    expect(lengthToEmu('1', 'in')).toBe(914400)
  })

  it('shows sizes without float noise', () => {
    expect(emuToLength(7560000, 'mm')).toBe('210')
    expect(emuToLength(10692000, 'cm')).toBe('29.7')
    expect(emuToLength(12192000, 'in')).toBe('13.33')
  })
})

describe('defaultNewSlideSize', () => {
  beforeEach(() => localStorage.clear())

  it('is unset until a size is saved, then returns it', () => {
    expect(defaultNewSlideSize()).toBeNull()
    saveDefaultNewSlideSize(presetSize('A4', true))
    expect(defaultNewSlideSize()).toEqual({ cx: 7560000, cy: 10692000 })
  })

  it('ignores a corrupt or out-of-range entry', () => {
    localStorage.setItem('ai-slides-default-size', '{not json')
    expect(defaultNewSlideSize()).toBeNull()
    localStorage.setItem('ai-slides-default-size', JSON.stringify({ cx: 10, cy: 7560000 }))
    expect(defaultNewSlideSize()).toBeNull()
  })
})
