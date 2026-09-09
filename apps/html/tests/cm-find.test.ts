import { describe, expect, it, vi } from 'vitest'
import { EditorState } from '@codemirror/state'
import { EditorView } from '@codemirror/view'
import { collectMatches, findHighlight } from '../src/renderer/source/cm-find'
import { cmFindTarget } from '../src/renderer/source/find-target'

const opts = { matchCase: false, wholeWord: false }

describe('collectMatches', () => {
  it('honors matchCase and wholeWord', () => {
    const doc = EditorState.create({ doc: '<p class="a">Cat cats CAT</p>' }).doc
    expect(collectMatches(doc, 'cat', opts)).toHaveLength(3)
    expect(collectMatches(doc, 'cat', { matchCase: false, wholeWord: true })).toHaveLength(2)
    expect(collectMatches(doc, 'CAT', { matchCase: true, wholeWord: false })).toHaveLength(1)
    expect(collectMatches(doc, '', opts)).toEqual([])
  })

  it('treats the query literally', () => {
    const doc = EditorState.create({ doc: 'a.b axb a.b' }).doc
    expect(collectMatches(doc, 'a.b', opts)).toEqual([
      { from: 0, to: 3 },
      { from: 8, to: 11 },
    ])
  })
})

describe('cmFindTarget', () => {
  function setup(text: string) {
    const changed = new Set<() => void>()
    const onChange = vi.fn()
    const view = new EditorView({
      state: EditorState.create({ doc: text, extensions: [findHighlight] }),
      parent: document.body,
      dispatchTransactions: (trs, v) => {
        v.update(trs)
        if (trs.some((tr) => tr.docChanged)) {
          onChange(v.state.doc.toString())
          for (const l of changed) l()
        }
      },
    })
    const target = cmFindTarget(view, (l) => {
      changed.add(l)
      return () => changed.delete(l)
    })
    return { view, target, onChange }
  }

  it('paints hits and replaces through regular transactions', () => {
    const { view, target, onChange } = setup('<h1>Hello</h1>\n<p>hello hello</p>')
    expect(target.search('hello', opts, 0)).toBe(3)
    expect(onChange).not.toHaveBeenCalled()
    let hits = 0
    view.state.field(findHighlight).between(0, view.state.doc.length, () => {
      hits++
    })
    expect(hits).toBe(3)

    target.replaceOne(0, 'Bye')
    expect(view.state.doc.toString()).toBe('<h1>Bye</h1>\n<p>hello hello</p>')
    expect(onChange).toHaveBeenCalledTimes(1)

    expect(target.search('hello', opts, 0)).toBe(2)
    target.replaceAll('x')
    expect(view.state.doc.toString()).toBe('<h1>Bye</h1>\n<p>x x</p>')
    expect(target.search('hello', opts, 0)).toBe(0)
    view.destroy()
  })

  it('notifies doc changes once per edit and stops after unsubscribe', () => {
    const { view, target } = setup('abc')
    const listener = vi.fn()
    const off = target.onDocChanged(listener)
    target.search('abc', opts, 0)
    target.activate(0)
    expect(listener).not.toHaveBeenCalled()
    view.dispatch({ changes: { from: 0, insert: 'z' } })
    expect(listener).toHaveBeenCalledTimes(1)
    off()
    view.dispatch({ changes: { from: 0, insert: 'z' } })
    expect(listener).toHaveBeenCalledTimes(1)
    view.destroy()
  })
})
