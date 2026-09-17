import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'

interface JsonNode {
  type: string
  attrs?: Record<string, unknown>
  content?: JsonNode[]
  text?: string
  marks?: { type: string; attrs?: Record<string, unknown> }[]
}

const ARABIC = '\u0639\u0631\u0628\u064a\u0629'

function render(node: JsonNode): HTMLElement {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [node] },
  })
  const dom = editor.view.dom.firstElementChild as HTMLElement
  editor.destroy()
  return dom
}

const para = (
  attrs: Record<string, unknown>,
  content: JsonNode[],
  type = 'docParagraph',
): JsonNode => ({ type, attrs: { docxIndex: null, ...attrs }, content })

describe('w:bidi paragraphs order runs without w:rtl left-to-right (Word)', () => {
  it('isolates the text as one LTR item when no run is rtl', () => {
    const el = render(para({ bidi: true }, [{ type: 'text', text: `${ARABIC} (LATIN)` }]))
    expect(el.style.direction).toBe('rtl')
    const wrap = el.firstElementChild as HTMLElement
    expect(wrap.className).toBe('doc-ltr-runs')
    expect(wrap.textContent).toBe(`${ARABIC} (LATIN)`)
  })

  it('keeps the RTL base when a run is an rtl run', () => {
    const el = render(
      para({ bidi: true }, [
        { type: 'text', text: ARABIC, marks: [{ type: 'docTextStyle', attrs: { cs: true } }] },
        { type: 'text', text: ' (LATIN)' },
      ]),
    )
    expect(el.querySelector('.doc-ltr-runs')).toBeNull()
  })

  it('leaves LTR paragraphs and empty bidi paragraphs alone', () => {
    expect(
      render(para({}, [{ type: 'text', text: 'abc' }])).querySelector('.doc-ltr-runs'),
    ).toBeNull()
    expect(render(para({ bidi: true }, [])).querySelector('.doc-ltr-runs')).toBeNull()
  })

  it('applies to headings and list items too', () => {
    const h = render(para({ bidi: true, level: 2 }, [{ type: 'text', text: 'T' }], 'docHeading'))
    expect(h.tagName).toBe('H2')
    expect(h.firstElementChild?.className).toBe('doc-ltr-runs')
    const li = render(
      para({ bidi: true, kind: 'bullet', ilvl: 0 }, [{ type: 'text', text: 'T' }], 'docListItem'),
    )
    expect(li.querySelector('.doc-ltr-runs')).not.toBeNull()
  })
})
