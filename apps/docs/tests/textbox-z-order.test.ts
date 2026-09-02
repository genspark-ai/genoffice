/**
 * Floating shape stacking: the relativeHeight rank orders overlapping floats
 * within the same z bands the image renderer uses (behind < -1 < text < 1 <=
 * front), so a full-page background shape stays under later-ranked text boxes.
 */
import { describe, expect, it } from 'vitest'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { textboxBoxStyle } from '../src/renderer/editor/protected-render'

const box = (extra: Partial<TextboxDisplay>): TextboxDisplay => ({
  paras: [],
  fill: '00FF00',
  floating: true,
  ...extra,
})

describe('textboxBoxStyle z bands', () => {
  it('front floats rank above the text layer', () => {
    expect(textboxBoxStyle(box({ z: 3 }))).toContain('z-index:5')
    // rank floor: the front band never reaches the text layer at 0
    expect(textboxBoxStyle(box({ z: -10 }))).toContain('z-index:1')
  })

  it('unranked floats keep DOM stacking (no z-index)', () => {
    expect(textboxBoxStyle(box({}))).not.toContain('z-index')
  })

  it('behind floats stay in the negative band, ordered by rank', () => {
    expect(textboxBoxStyle(box({ behind: true }))).toContain('z-index:-1000')
    expect(textboxBoxStyle(box({ behind: true, z: 4 }))).toContain('z-index:-996')
    // band ceiling: a huge rank must not lift a behind box above the text
    expect(textboxBoxStyle(box({ behind: true, z: 5000 }))).toContain('z-index:-1')
  })

  it('non-floating boxes never get a z-index', () => {
    expect(textboxBoxStyle(box({ floating: false, z: 3 }))).not.toContain('z-index')
  })
})
