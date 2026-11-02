// @vitest-environment jsdom
import { Editor } from '@tiptap/core'
import { Decoration } from '@tiptap/pm/view'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  tabReflowMeta,
  tabReflowStart,
  tabStopPluginKey,
} from '../src/renderer/editor/decoration-extensions'
import { editorExtensions } from '../src/renderer/editor/extensions'

describe('tabReflowStart', () => {
  it('reports the lowest position whose paragraph height moved', () => {
    const before = new Map([
      [40, 20],
      [10, 20],
      [25, 20],
    ])
    const after = new Map([
      [40, 40],
      [10, 20.3],
      [25, 41],
    ])
    expect(tabReflowStart(before, (pos) => after.get(pos) ?? null)).toBe(25)
  })

  it('stays quiet for sub-pixel drift and unmeasurable paragraphs', () => {
    const before = new Map([
      [10, 20],
      [25, 20],
    ])
    expect(tabReflowStart(before, (pos) => (pos === 10 ? 20.4 : null))).toBeNull()
    expect(tabReflowStart(new Map(), () => 99)).toBeNull()
  })
})

describe('tab decoration transactions signal the paginator', () => {
  const realRect = HTMLElement.prototype.getBoundingClientRect
  let paraHeight = 20
  let host: HTMLDivElement
  let editor: Editor
  let signals: unknown[]

  beforeEach(() => {
    HTMLElement.prototype.getBoundingClientRect = function (this: HTMLElement) {
      const h = this.tagName === 'P' ? paraHeight : 0
      return {
        x: 0,
        y: 0,
        left: 0,
        top: 0,
        right: 600,
        bottom: h,
        width: 600,
        height: h,
      } as DOMRect
    }
    host = document.createElement('div')
    document.body.appendChild(host)
    editor = new Editor({
      element: host,
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: 'plain' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'a\tb' }] },
        ],
      } as never,
    })
    signals = []
    editor.on('transaction', ({ transaction }) => {
      const m = transaction.getMeta(tabReflowMeta)
      if (m !== undefined) signals.push(m)
    })
  })

  afterEach(() => {
    editor.destroy()
    host.remove()
    HTMLElement.prototype.getBoundingClientRect = realRect
  })

  const decorate = () =>
    editor.view.dispatch(
      editor.state.tr.setMeta(tabStopPluginKey, [
        Decoration.inline(9, 10, { class: 'doc-tab', style: 'tab-size:48px' }),
      ]),
    )

  it('emits the tab paragraph position once its height changed', () => {
    paraHeight = 40
    decorate()
    expect(signals).toEqual([7])
  })

  it('stays silent when the decorations left the height alone', () => {
    decorate()
    expect(signals).toEqual([])
  })

  it('does not fire for document edits, which the paginator already sees', () => {
    paraHeight = 40
    editor.commands.insertContentAt(1, 'x')
    expect(signals).toEqual([])
  })
})
