/**
 * A run's explicit `w:spacing w:val="0"` must reach the DOM as letter-spacing:0
 * on every render path so it overrides the paragraph style's spacing (sas2 058:
 * a Titre1 heading spaced 2pt by its style wrapped to two lines and pushed a
 * page-break paragraph onto an extra page).
 */
import { Editor } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import { describe, expect, it } from 'vitest'
import type { TextboxDisplay } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { inlineToRuns, runsToInline } from '../src/renderer/editor/convert'
import { renderTextboxSpec } from '../src/renderer/editor/protected-render'
import { runLetterSpacingCss } from '../src/renderer/line-metrics'

const ZERO = /^0(px|pt)?$/

describe('explicit zero letter spacing', () => {
  it('survives the run -> mark -> run round trip', () => {
    const inline = runsToInline([{ text: 'HEAD', charSpacingTwips: 0 }])
    const marks = (inline[0] as { marks?: Array<{ type: string; attrs: Record<string, unknown> }> })
      .marks
    expect(marks?.find((m) => m.type === 'docTextStyle')?.attrs.charSpacingTwips).toBe(0)
    expect(inlineToRuns(inline)[0].charSpacingTwips).toBe(0)
  })

  it('renders letter-spacing:0 on the editable span', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'HEAD',
                marks: [{ type: 'docTextStyle', attrs: { charSpacingTwips: 0 } }],
              },
            ],
          },
        ],
      } as never,
    })
    const span = editor.view.dom.querySelector('span[style]') as HTMLElement
    expect(span.style.letterSpacing).toMatch(ZERO)
    editor.destroy()
  })

  it('reaches the static render paths', () => {
    expect(runLetterSpacingCss({ text: 'A', charSpacingTwips: 0 })).toBe('0')
    const box: TextboxDisplay = { paras: [{ runs: [{ text: 'HEAD', charSpacingTwips: 0 }] }] }
    const dom = DOMSerializer.renderSpec(document, renderTextboxSpec(box) as never)
      .dom as HTMLElement
    const span = dom.querySelector('span[style]') as HTMLElement
    expect(span.style.letterSpacing).toMatch(ZERO)
  })
})
