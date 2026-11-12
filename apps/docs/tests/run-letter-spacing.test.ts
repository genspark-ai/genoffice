/**
 * w:spacing / w:w on runs must reach the static render paths (textboxes,
 * header/footer DOM) as letter-spacing — PDF converters express word gaps as
 * whitespace-only runs with negative w:spacing.
 */
import { describe, expect, it } from 'vitest'
import { DOMSerializer } from '@tiptap/pm/model'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { renderTextboxSpec } from '../src/renderer/editor/protected-render'
import { runLetterSpacingCss } from '../src/renderer/line-metrics'

const render = (spec: unknown): HTMLElement =>
  DOMSerializer.renderSpec(document, spec as never).dom as HTMLElement

describe('runLetterSpacingCss', () => {
  it('converts w:spacing twips to pt', () => {
    expect(runLetterSpacingCss({ text: ' ', charSpacingTwips: -49 })).toBe('-2.45pt')
  })

  it('combines w:spacing and w:w into one calc()', () => {
    expect(runLetterSpacingCss({ text: 'A', charSpacingTwips: 20, charScalePct: 105 })).toBe(
      'calc(1pt + 0.026em)',
    )
  })

  it('returns null without either property', () => {
    expect(runLetterSpacingCss({ text: 'A' })).toBeNull()
  })
})

describe('static textbox render', () => {
  it('emits letter-spacing on spaced runs', () => {
    const box: TextboxDisplay = {
      paras: [
        {
          runs: [
            { text: 'END' },
            { text: ' ', charSpacingTwips: -49 },
            { text: 'OF', charScalePct: 105 },
          ],
        },
      ],
    }
    const spans = [...render(renderTextboxSpec(box)).querySelectorAll('span')].filter(
      (s) => !s.children.length,
    )
    expect(spans.map((s) => s.style.letterSpacing)).toEqual(['', '-2.45pt', '0.026em'])
  })
})
