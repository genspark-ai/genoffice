import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { describe, expect, it, vi } from 'vitest'
import { editorExtensions, hasTextSelection } from '../src/renderer/editor/extensions'

const CSS = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), '../src/renderer/styles.css'),
  'utf8',
)

// jsdom has no elementsFromPoint; the run-picture drag probe reads it on every
// mousedown that reaches ProseMirror's plugins.
if (typeof document.elementsFromPoint !== 'function') {
  Object.defineProperty(document, 'elementsFromPoint', {
    configurable: true,
    value: () => [] as Element[],
  })
}

const TABLE_MODEL = {
  rows: [[{ paras: ['Alpha Beta Gamma Delta'] }, { paras: ['Epsilon Zeta'] }]],
}

function makeEditor() {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'before' }] },
        { type: 'docProtected', attrs: { blockType: 'table', label: 'Table', table: TABLE_MODEL } },
        {
          type: 'docProtected',
          attrs: {
            blockType: 'passthrough',
            label: 'Excel worksheet (embedded)',
            previewText: 'Quarterly totals 2026',
          },
        },
      ],
    },
  })
}

const mouse = (type: string, init: MouseEventInit = {}) =>
  new MouseEvent(type, { bubbles: true, cancelable: true, button: 0, ...init })

/** a fake non-collapsed browser selection, or an empty one. ProseMirror's
 *  selectionToDOM uses these methods after a dispatch, so provide real no-ops. */
function mockSelection(text: string) {
  const range = document.createRange()
  range.selectNodeContents(document.body)
  const addRange = vi.fn()
  const removeAllRanges = vi.fn()
  const spy = vi.spyOn(window, 'getSelection').mockReturnValue({
    isCollapsed: text.length === 0,
    toString: () => text,
    rangeCount: text.length === 0 ? 0 : 1,
    getRangeAt: () => range,
    addRange,
    removeAllRanges,
    collapse: vi.fn(),
    extend: vi.fn(),
    setBaseAndExtent: vi.fn(),
  } as unknown as Selection)
  return Object.assign(spy, { addRange, removeAllRanges })
}

describe('hasTextSelection', () => {
  it('is false for an empty or collapsed selection, true for real characters', () => {
    mockSelection('')
    expect(hasTextSelection()).toBe(false)
    mockSelection('   ')
    expect(hasTextSelection()).toBe(false)
    mockSelection('Alpha Beta')
    expect(hasTextSelection()).toBe(true)
    vi.restoreAllMocks()
  })
})

describe('selecting text inside an embedded object', () => {
  it.each([
    ['a protected table cell', '.doc-protected[data-doc-protected="table"] td'],
    ['a read-only OLE caption', '.doc-protected-passthrough .doc-protected-label'],
  ])(
    'a drag across %s keeps a partial text selection instead of selecting the object',
    (_name, selector) => {
      const editor = makeEditor()
      const dom = editor.view.dom.querySelector('.doc-protected') as HTMLElement
      const target = editor.view.dom.querySelector(selector) as HTMLElement
      const owner = target.closest('.doc-protected') as HTMLElement
      const drag = mockSelection('Alpha Beta')

      target.dispatchEvent(mouse('mousedown'))
      // the browser owns the gesture: HTML5 object dragging must stay off until release
      expect(owner.draggable).toBe(false)
      expect(editor.state.selection instanceof NodeSelection).toBe(false)

      document.dispatchEvent(mouse('mouseup'))
      expect(editor.state.selection instanceof NodeSelection).toBe(false)
      expect(owner.draggable).toBe(true)
      drag.mockRestore()
      editor.destroy()
      expect(dom).toBeTruthy()
    },
  )

  it('a text drag after a prior click on the object clears the editor selection', () => {
    const editor = makeEditor()
    const label = editor.view.dom.querySelector('.doc-protected-label') as HTMLElement
    const owner = label.closest('.doc-protected') as HTMLElement
    let pos = -1
    editor.state.doc.descendants((node, nodePos) => {
      if (node.type.name === 'docProtected' && node.attrs.blockType === 'passthrough') {
        pos = nodePos
      }
    })
    expect(pos).toBeGreaterThan(-1)
    // The common flow: click the frame (NodeSelection), then drag a few chars.
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    expect(editor.state.selection).toBeInstanceOf(NodeSelection)

    const drag = mockSelection('Excel workshee')
    label.dispatchEvent(mouse('mousedown'))
    document.dispatchEvent(mouse('mouseup'))

    // Copy and Delete must act on the dragged characters, not on the atom.
    expect(editor.state.selection.empty).toBe(true)
    expect(editor.state.selection instanceof NodeSelection).toBe(false)
    // The dispatch's selectionToDOM must not wipe the visible character range.
    expect(drag.removeAllRanges).toHaveBeenCalled()
    expect(drag.addRange).toHaveBeenCalled()
    expect(owner.draggable).toBe(true)
    drag.mockRestore()
    editor.destroy()
  })

  it('a plain click with no characters still selects the whole object', () => {
    const editor = makeEditor()
    const label = editor.view.dom.querySelector('.doc-protected-label') as HTMLElement
    const owner = label.closest('.doc-protected') as HTMLElement
    const empty = mockSelection('')

    label.dispatchEvent(mouse('mousedown'))
    document.dispatchEvent(mouse('mouseup'))
    expect(editor.state.selection instanceof NodeSelection).toBe(true)
    expect(owner.draggable).toBe(true)
    empty.mockRestore()
    editor.destroy()
  })

  it('double-click still enters text-edit mode on a protected table cell', () => {
    const editor = makeEditor()
    const td = editor.view.dom.querySelector('.doc-protected td') as HTMLElement
    expect(td.getAttribute('contenteditable')).toBe('false')
    td.dispatchEvent(mouse('mousedown'))
    td.dispatchEvent(mouse('dblclick'))
    expect(td.getAttribute('contenteditable')).toBe('true')
    editor.destroy()
  })

  it('the move handle still selects the whole object', () => {
    const editor = makeEditor()
    const handle = editor.view.dom.querySelector('.doc-move-handle') as HTMLElement
    // capture listeners still see a non-bubbling press; ProseMirror's
    // posAtCoords probe (which needs jsdom's missing elementFromPoint) does not
    handle.dispatchEvent(mouse('mousedown', { bubbles: false }))
    expect(editor.state.selection instanceof NodeSelection).toBe(true)
    editor.destroy()
  })
})

describe('embedded-object text CSS', () => {
  it('lets the reader select text inside a non-editing object', () => {
    const rule = /\.doc-protected\s+:is\([\s\S]*?\}/.exec(CSS)?.[0] ?? ''
    expect(rule).toContain('user-select: text')
    expect(rule).toContain('.doc-protected-label')
    expect(rule).toContain('td')
  })

  it('no longer disables selection on non-editing text boxes and table cells', () => {
    const rule =
      /\.doc-protected-textboxes:not\(\.doc-content-editing\)[\s\S]*?\}/.exec(CSS)?.[0] ?? ''
    expect(rule).toContain('cursor: text')
    expect(rule).not.toMatch(/user-select:\s*none/)
  })
})
