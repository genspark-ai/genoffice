import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { FormulaDisplay } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildProtectedDom } from '../src/renderer/editor/extensions'
import { renderFormulaSpec } from '../src/renderer/editor/protected-render'
import { runsToInline } from '../src/renderer/editor/convert'

const FORMULA: FormulaDisplay = {
  tokens: ['dy', 'dx'],
  mathml: '<math display="block"><mfrac><mi>dy</mi><mi>dx</mi></mfrac></math>',
  omml: '<m:oMath><m:f/></m:oMath>',
  sizeHalfPoints: 18,
  align: 'left',
  spaceAfterTwips: 0,
}

function protectedNode(formulaDisplay: FormulaDisplay) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docProtected',
          attrs: { docxIndex: 0, blockType: 'passthrough', label: 'Equation', formulaDisplay },
        },
      ],
    },
  })
  const node = editor.state.doc.firstChild!
  editor.destroy()
  return node
}

describe('display equation rendering', () => {
  it('the MathML host and the token strip take the run size', () => {
    const spec = renderFormulaSpec(FORMULA) as unknown[]
    const host = spec[2] as [string, Record<string, string>]
    const strip = spec[3] as [string, Record<string, string>]
    expect(host[1].class).toBe('doc-formula-math')
    expect(host[1].style).toBe('font-size:9pt')
    expect(strip[1].style).toBe('font-size:9pt')
    const plain = renderFormulaSpec({ tokens: ['x'] }) as unknown[]
    expect((plain[1] as Record<string, string>).style).toBeUndefined()
  })

  it('the wrapper follows m:jc and the direct paragraph spacing', () => {
    const el = buildProtectedDom(protectedNode(FORMULA))
    expect(el.className).toContain('doc-protected-formula-display')
    expect(el.style.textAlign).toBe('left')
    expect(el.style.marginBottom).toBe('0pt')
    const centred = buildProtectedDom(protectedNode({ ...FORMULA, align: 'centerGroup' }))
    expect(centred.style.textAlign).toBe('center')
    const bare = buildProtectedDom(protectedNode({ tokens: ['x'], mathml: '<math/>' }))
    expect(bare.getAttribute('style')).toBeNull()
  })

  it('an inline maths run passes its size to the atom node', () => {
    const [node] = runsToInline([
      { text: 'dydx', math: { omml: '<m:oMath><m:f/></m:oMath>' }, sizeHalfPoints: 18 },
    ])
    expect(node.type).toBe('docInlineMath')
    expect(node.attrs?.sizeHalfPoints).toBe(18)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', attrs: { docxIndex: 0 }, content: [node] }],
      },
    })
    const math = editor.view.dom.querySelector('span[data-inline-math]') as HTMLElement
    expect(math.style.fontSize).toBe('9pt')
    editor.destroy()
  })
})
