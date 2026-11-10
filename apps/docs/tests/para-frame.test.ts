import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { paraFrameCss, textFlowCss } from '../src/renderer/editor/para-frame'

function paraEl(attrs: Record<string, unknown>): HTMLParagraphElement {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: null, ...attrs },
          content: [{ type: 'text', text: 'x' }],
        },
      ],
    },
  })
  const p = editor.view.dom.querySelector('p')!.cloneNode(true) as HTMLParagraphElement
  editor.destroy()
  return p
}

describe('w:framePr paragraph frames', () => {
  it('a page-anchored wrap-around frame floats at its x minus the section margin', () => {
    const frame = {
      wTwips: 2851,
      hTwips: 3241,
      hRule: 'atLeast',
      xTwips: 1816,
      yTwips: 299,
      vAnchor: 'text',
      wrap: 'around',
      hSpaceTwips: 180,
    }
    const p = paraEl({ frame: JSON.stringify(frame), textDirection: 'tbRl' })
    expect(p.classList.contains('doc-para-frame')).toBe(true)
    expect(p.classList.contains('doc-para-frame-float')).toBe(true)
    expect(p.style.width).toBe('142.55pt')
    expect(p.style.height).toBe('162.05pt')
    expect(p.style.minHeight).toBe('')
    expect(p.style.cssFloat).toBe('left')
    expect(p.style.marginLeft).toBe('calc(90.80pt - var(--doc-margin-left, 0px))')
    expect(p.style.marginRight).toBe('9pt')
    expect(p.style.marginTop).toBe('14.95pt')
    expect(p.style.writingMode).toBe('vertical-rl')
  })

  it('a tight-wrap frame floats like an around-wrap one', () => {
    const p = paraEl({
      frame: JSON.stringify({ wTwips: 1200, xTwips: 0, yTwips: 0, wrap: 'tight', xAlign: 'right' }),
    })
    expect(p.classList.contains('doc-para-frame-float')).toBe(true)
    expect(p.style.cssFloat).toBe('right')
  })

  it('exact height clips; notBeside frames stay in flow', () => {
    const { classes, styles } = paraFrameCss(
      {
        wTwips: 1200,
        hTwips: 600,
        hRule: 'exact',
        xTwips: 0,
        yTwips: 0,
        wrap: 'notBeside',
        xAlign: 'right',
      },
      null,
    )
    expect(classes).toEqual(['doc-para-frame'])
    expect(styles).toContain('height:30.00pt')
    expect(styles).toContain('overflow:hidden')
    expect(styles.some((s) => s.startsWith('float:'))).toBe(false)
    expect(styles.some((s) => s.startsWith('margin-left'))).toBe(false)
  })

  it('a centered wrap-around frame takes its own line instead of floating', () => {
    const { classes, styles } = paraFrameCss(
      { wTwips: 1200, xTwips: 0, yTwips: 0, wrap: 'around', xAlign: 'center' },
      null,
    )
    expect(classes).toEqual(['doc-para-frame'])
    expect(styles).toContain('margin-left:auto')
    expect(styles.some((s) => s.startsWith('float:'))).toBe(false)
  })

  it('maps the three sideways flows to writing modes', () => {
    expect(textFlowCss('tbRl')).toEqual(['writing-mode:vertical-rl'])
    expect(textFlowCss('tbRlV')).toEqual(['writing-mode:vertical-rl', 'text-orientation:sideways'])
    expect(textFlowCss('btLr')).toEqual(['writing-mode:sideways-lr'])
    expect(textFlowCss(null)).toEqual([])
  })

  it('a paragraph without a frame carries no frame styling even with a text direction', () => {
    const p = paraEl({ textDirection: 'tbRl' })
    expect(p.classList.contains('doc-para-frame')).toBe(false)
    expect(p.style.writingMode).toBe('')
  })
})
