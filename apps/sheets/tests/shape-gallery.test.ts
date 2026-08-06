import { describe, expect, it } from 'vitest'

import {
  isPillPreset,
  presetPath,
  presetPolygon,
} from '../../../packages/pptx-render/src/preset-geometry'
import { SHAPE_GALLERY_GROUPS } from '../../../packages/ui/src/shape-gallery'
import { ADDABLE_SHAPE_TYPES } from '../src/shared/shape-types'

// The ribbon gallery (ExcelShell's SHEET_SHAPE_GROUPS) is the shared groups minus Lines
const galleryPrsts = SHAPE_GALLERY_GROUPS.filter(
  (group) => group.groupKey !== 'ribbonShapeGroupLines',
).flatMap((group) => group.shapes.map((shape) => shape.prst))

describe('shape gallery', () => {
  it('every gallery prst is saveable (in ADDABLE_SHAPE_TYPES)', () => {
    const addable = new Set<string>(ADDABLE_SHAPE_TYPES)
    expect(galleryPrsts.filter((prst) => !addable.has(prst))).toEqual([])
  })

  it('every gallery prst has render geometry (no fallback chips)', () => {
    const covered = (prst: string): boolean =>
      prst === 'rect' ||
      prst === 'roundRect' ||
      prst === 'ellipse' ||
      prst === 'flowChartProcess' ||
      isPillPreset(prst) ||
      presetPolygon(prst, 100, 100) !== null ||
      presetPath(prst, 100, 100) !== null
    expect(galleryPrsts.filter((prst) => !covered(prst))).toEqual([])
  })

  it('has no duplicate prsts', () => {
    expect(new Set(galleryPrsts).size).toBe(galleryPrsts.length)
  })
})
