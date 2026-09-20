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

  it('refreshes cached positions after the doc changes', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, schema.text('first')),
      schema.node('paragraph', null, schema.text('second')),
    ])
    const view = new EditorView(document.body.appendChild(document.createElement('div')), {
      state: EditorState.create({ doc }),
    })
    const positions = new TopLevelPositions(view)
    const first = view.dom.children[0] as Element
    expect(positions.of(first)).toEqual({ from: 0, to: 7 })
    view.dispatch(view.state.tr.insert(0, schema.node('paragraph', null, schema.text('new'))))
    const shifted = blockOffset(view.state.doc, 1)
    const second = view.dom.children[1] as Element
    // a never-invalidated walk would miss the new DOM and report stale spots
    expect(positions.of(second)).toEqual({ from: shifted, to: shifted + 7 })
    const head = view.dom.children[0] as Element
    expect(positions.of(head)).toEqual({ from: 0, to: 5 })
    view.destroy()
  })

  it('invalidate() drops the cached walk and the next lookup rebuilds', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, schema.text('first')),
      schema.node('paragraph', null, schema.text('second')),
    ])
    const view = new EditorView(document.body.appendChild(document.createElement('div')), {
      state: EditorState.create({ doc }),
    })
    const positions = new TopLevelPositions(view)
    expect(positions.of(view.dom.children[0] as Element)).toEqual({ from: 0, to: 7 })
    positions.invalidate()
    expect(positions.of(view.dom.children[1] as Element)).toEqual({ from: 7, to: 15 })
    view.dispatch(view.state.tr.insert(0, schema.node('paragraph', null, schema.text('new'))))
    positions.invalidate()
    expect(positions.of(view.dom.children[0] as Element)).toEqual({ from: 0, to: 5 })
    // detached elements never resolve to a block
    expect(positions.of(document.createElement('p'))).toBeNull()
    view.destroy()
  })

  it('falls back to posAtDOM when the child-desc walk misses an element', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, schema.text('first')),
      schema.node('paragraph', null, schema.text('second')),
    ])
    const container = document.createElement('div')
    const first = document.createElement('p')
    const widget = document.createElement('hr')
    container.append(first, widget)
    const throwing = document.createElement('p')
    container.append(throwing)
    // an empty child-desc walk: every lookup must go through the fallback
    const fake = {
      dom: container,
      state: { doc },
      docView: { children: [] },
      posAtDOM(el: Element) {
        if (el === throwing) throw new RangeError('no position')
        return el === first ? 1 : 8
      },
      nodeDOM(pos: number) {
        return pos === 0 ? first : undefined
      },
    } as unknown as EditorView
    const positions = new TopLevelPositions(fake)
    expect(positions.of(first)).toEqual({ from: 0, to: 7 })
    // a widget resolves inside a block it does not own: still null
    expect(positions.of(widget)).toBeNull()
    // posAtDOM failures stay null instead of throwing
    expect(positions.of(throwing)).toBeNull()
  })
})

const blockOffset = (doc: { child(i: number): { nodeSize: number } }, index: number): number => {
  let offset = 0
  for (let i = 0; i < index; i++) offset += doc.child(i).nodeSize
  return offset
}
