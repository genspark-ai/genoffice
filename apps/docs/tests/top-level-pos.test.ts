import { describe, expect, it } from 'vitest'
import { Schema } from '@tiptap/pm/model'
import { EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet, EditorView } from '@tiptap/pm/view'
import { TopLevelPositions } from '../src/renderer/editor/top-level-pos'

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { content: 'text*', group: 'block', toDOM: () => ['p', 0] },
    quote: { content: 'paragraph+', group: 'block', toDOM: () => ['blockquote', 0] },
    text: {},
  },
})

describe('TopLevelPositions', () => {
  it('matches posAtDOM for every top-level block, with widgets between them', () => {
    const blocks = []
    for (let i = 0; i < 40; i++) {
      const p = schema.node('paragraph', null, schema.text(`para ${i} ${'x'.repeat(i % 7)}`))
      blocks.push(
        i % 5 === 0
          ? schema.node('quote', null, [p, schema.node('paragraph', null, schema.text('q'))])
          : p,
      )
    }
    const doc = schema.node('doc', null, blocks)
    const widgets: Decoration[] = []
    doc.forEach((_, offset, i) => {
      if (i % 3 === 0)
        widgets.push(Decoration.widget(offset, () => document.createElement('hr'), { side: -1 }))
    })
    const set = DecorationSet.create(doc, widgets)
    const view = new EditorView(document.body.appendChild(document.createElement('div')), {
      state: EditorState.create({ doc }),
      decorations: () => set,
    })
    const positions = new TopLevelPositions(view)
    let checked = 0
    for (const el of Array.from(view.dom.children)) {
      if (el.tagName === 'HR') {
        expect(positions.of(el)).toBeNull()
        continue
      }
      const $pos = view.state.doc.resolve(view.posAtDOM(el, 0))
      expect(positions.of(el)).toEqual({ from: $pos.before(1), to: $pos.after(1) })
      // nested elements resolve to their top-level block
      const inner = el.querySelector('p') ?? el
      expect(positions.of(inner)).toEqual({ from: $pos.before(1), to: $pos.after(1) })
      checked++
    }
    expect(checked).toBe(40)
    expect(positions.of(document.createElement('p'))).toBeNull()
    view.destroy()
  })
})
