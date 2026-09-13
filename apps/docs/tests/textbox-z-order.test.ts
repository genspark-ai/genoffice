/**
 * Floating shape stacking: the relativeHeight rank orders overlapping floats
 * within the same z bands the image renderer uses (behind < -1 < text < 1 <=
 * front), so a full-page background shape stays under later-ranked text boxes.
 */
import { describe, expect, it } from 'vitest'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { DOMSerializer } from '@tiptap/pm/model'
import { renderTextboxSpec, textboxBoxStyle } from '../src/renderer/editor/protected-render'

const box = (extra: Partial<TextboxDisplay>): TextboxDisplay => ({
  paras: [],
  fill: '00FF00',
  floating: true,
  ...extra,
})

describe('textboxBoxStyle border', () => {
  it('a stroked box always sets its border width (the CSS default reserves no edge)', () => {
    expect(textboxBoxStyle(box({ borderColor: '000000' }))).toContain('border-width:1px')
    expect(textboxBoxStyle(box({ borderColor: '000000', borderWidthPx: 2 }))).toContain(
      'border-width:2px',
    )
  })

  it('an unstroked fixed-size box keeps its full text rectangle', () => {
    expect(textboxBoxStyle(box({ widthPx: 62, heightPx: 16 }))).not.toContain('border-width')
  })
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

describe('renderTextboxSpec paragraph line sizing', () => {
  const paraStyles = (paras: TextboxDisplay['paras']): HTMLElement[] => {
    const { dom } = DOMSerializer.renderSpec(document, renderTextboxSpec(box({ paras })) as never)
    return Array.from((dom as HTMLElement).querySelectorAll('.doc-textbox-para'))
  }

  // Word probe 2026-09-11: a space-only paragraph is sized by its mark, never the space run
  it('a space-only paragraph takes the Latin mark factor and no run strut', () => {
    const [sized, spaceOnly, marked] = paraStyles([
      { runs: [{ text: 'ab', sizeHalfPoints: 14 }] },
      { runs: [{ text: ' ', sizeHalfPoints: 8 }] },
      { runs: [{ text: ' ', sizeHalfPoints: 8 }], emptyRunSizeHalfPoints: 8 },
    ])
    expect(sized.style.getPropertyValue('--doc-strut')).toBe('7pt')
    expect(spaceOnly.style.getPropertyValue('--doc-strut')).toBe('')
    expect(spaceOnly.style.getPropertyValue('--doc-line-factor')).toBe(
      'var(--doc-line-factor-latin,1.2)',
    )
    expect(spaceOnly.style.fontSize).toBe('')
    expect(marked.style.fontSize).toBe('4pt')
  })
})
