import { afterEach, describe, expect, it } from 'vitest'

import { pixelsToCharacterWidth, setWorkbookMdw } from '../src/renderer/app-constants'
import { characterWidthToPixels } from '../src/renderer/univer-sync'

afterEach(() => setWorkbookMdw(7))

// Mac Excel print geometry, calib/narrow-col-width/calibri11-excel.pdf
// (Calibri 11 Normal, MDW 6pt; Calibri 12 measures identically). Column
// edges read from the PDF vector paths; values in points.
const CALIBRI_11_POINTS: ReadonlyArray<readonly [number, number]> = [
  [0.25, 1],
  [0.5, 3],
  [0.75, 4],
  [1, 6],
  [1.25, 7],
  [1.56, 9],
  [1.75, 10],
  [2, 12],
  [2.25, 13],
  [2.5, 15],
  [3, 18],
  [3.5, 21],
  [4, 24],
  [4.5, 27],
  [5, 30],
  [6, 36],
  [7, 42],
  [8.43, 51],
  [9, 54],
  [10, 60],
  [12, 72],
  [15, 90],
  [20, 120],
]

describe('characterWidthToPixels against Mac Excel print geometry', () => {
  it('matches every calibrated Calibri 11 width (1.56 chars is 12px, not 17)', () => {
    setWorkbookMdw(8)
    for (const [chars, points] of CALIBRI_11_POINTS) {
      expect(characterWidthToPixels(chars), `${chars} chars`).toBeCloseTo((points * 4) / 3, 9)
    }
  })

  it('has no fixed padding: widths scale through the origin', () => {
    setWorkbookMdw(8)
    expect(characterWidthToPixels(0)).toBe(0)
    expect(characterWidthToPixels(2)).toBeCloseTo(2 * characterWidthToPixels(1), 9)
    expect(characterWidthToPixels(15)).toBeCloseTo(3 * characterWidthToPixels(5), 9)
  })

  it('round-trips through pixelsToCharacterWidth within 1/256 char', () => {
    setWorkbookMdw(8)
    for (const chars of [1.56, 2.78, 8.43, 34.83203125, 50.86]) {
      const back = pixelsToCharacterWidth(characterWidthToPixels(chars))
      expect(characterWidthToPixels(back), `${chars} chars`).toBeCloseTo(
        characterWidthToPixels(chars),
        9,
      )
    }
  })
})
