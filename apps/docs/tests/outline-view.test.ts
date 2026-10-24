import { afterEach, describe, expect, it } from 'vitest'
import { Editor, type JSONContent } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { outlineViewKey } from '../src/renderer/editor/outline-view'
import { hiddenOutlineHeadings, outlineHasChildren } from '../src/renderer/editor/outline-view'
import { headingKey } from '../src/renderer/editor/nav-outline'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))

const para = (text: string): JSONContent => ({
  type: 'docParagraph',
  attrs: { docxIndex: null },
  content: text ? [{ type: 'text', text }] : [],
})
const heading = (level: number, text: string): JSONContent => ({
  type: 'docHeading',
  attrs: { level },
  content: [{ type: 'text', text }],
})

const refs = (levels: number[]): Array<{ level: number; text: string }> =>
  levels.map((level, i) => ({ level, text: `H${i}` }))

const OUTLINE_DOC = [
  heading(1, 'Alpha'),
  para('alpha body'),
  heading(2, 'Beta'),
  para('beta body'),
  heading(3, 'Gamma'),
  para('gamma body'),
  heading(1, 'Delta'),
  para('delta body'),
]

function createEditor(content: JSONContent[]): Editor {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
  editors.push(editor)
  return editor
}

const mousedown = (el: Element) =>
  el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))

describe('hiddenOutlineHeadings', () => {
  it('hides the descendant heading rows of a folded heading', () => {
    // Alpha(1) > Beta(2) > Gamma(3), Delta(1): folding Alpha hides Beta and Gamma
    expect(
      hiddenOutlineHeadings(refs([1, 2, 3, 1]), new Set([headingKey({ level: 1, text: 'H0' })])),
    ).toEqual(new Set([1, 2]))
  })

  it('hides only below the fold boundary (same or higher level ends the subtree)', () => {
    // folding the level-2 heading hides only the level-3 row under it
    expect(
      hiddenOutlineHeadings(refs([1, 2, 3, 1, 2]), new Set([headingKey({ level: 2, text: 'H1' })])),
    ).toEqual(new Set([2]))
  })

  it('stacks folds: an ancestor fold hides rows its own fold would keep', () => {
    const both = new Set([
      headingKey({ level: 1, text: 'H0' }),
      headingKey({ level: 2, text: 'H1' }),
    ])
    expect(hiddenOutlineHeadings(refs([1, 2, 3, 1, 2]), both)).toEqual(new Set([1, 2]))
    expect(
      hiddenOutlineHeadings(refs([1, 2, 3, 1, 2]), new Set([headingKey({ level: 2, text: 'H1' })])),
    ).toEqual(new Set([2]))
  })

  it('keys on level + text, so equal headings fold together', () => {
    const headings = [
      { level: 1, text: 'A' },
      { level: 2, text: 'X' },
      { level: 1, text: 'A' },
      { level: 2, text: 'Y' },
    ]
    expect(hiddenOutlineHeadings(headings, new Set([headingKey({ level: 1, text: 'A' })]))).toEqual(
      new Set([1, 3]),
    )
  })

  it('hides nothing when nothing is folded, and treats a deeper first heading as a root', () => {
    expect(hiddenOutlineHeadings(refs([1, 2, 1]), new Set())).toEqual(new Set())
    expect(
      hiddenOutlineHeadings(refs([3, 1]), new Set([headingKey({ level: 3, text: 'H0' })])),
    ).toEqual(new Set())
  })
})

describe('outlineHasChildren', () => {
  it('is true only while deeper headings sit below before the subtree ends', () => {
    const levels = refs([1, 2, 1, 2])
    expect(outlineHasChildren(levels, 0)).toBe(true)
    expect(outlineHasChildren(levels, 1)).toBe(false)
    expect(outlineHasChildren(levels, 2)).toBe(true)
    expect(outlineHasChildren(levels, 3)).toBe(false)
  })
})

describe('OutlineViewExtension', () => {
  it('draws handles only while outline view is on', () => {
    const editor = createEditor(OUTLINE_DOC)
    expect(editor.view.dom.querySelectorAll('.outline-handle').length).toBe(0)
    expect(editor.commands.setOutlineViewActive(true)).toBe(true)
    expect(editor.view.dom.querySelectorAll('.outline-handle').length).toBe(4)
    expect(editor.commands.setOutlineViewActive(false)).toBe(true)
    expect(editor.view.dom.querySelectorAll('.outline-handle').length).toBe(0)
  })

  it('expanded headings show ⊖, childless ones a dimmed leaf handle', () => {
    const editor = createEditor(OUTLINE_DOC)
    editor.commands.setOutlineViewActive(true)
    const handles = [...editor.view.dom.querySelectorAll('.outline-handle')]
    // Alpha, Beta and Gamma have children; Delta is a leaf
    expect(handles.map((h) => h.textContent)).toEqual(['⊖', '⊖', '⊖', '⊖'])
    expect(handles[3].classList.contains('is-leaf')).toBe(true)
    expect(handles[0].classList.contains('is-leaf')).toBe(false)
  })

  it('clicking the handle folds and unfolds the subtree', () => {
    const editor = createEditor(OUTLINE_DOC)
    editor.commands.setOutlineViewActive(true)
    const dom = () => editor.view.dom
    const alphaHandle = dom().querySelectorAll('.outline-handle')[0]
    mousedown(alphaHandle) // preventDefault makes dispatchEvent resolve false; the DOM is the assertion
    // Alpha's subtree rows (Beta h2, Gamma h3) hide; Delta stays visible
    const headings = () => [...dom().querySelectorAll('h1, h2, h3')]
    expect(headings().map((h) => h.classList.contains('outline-collapsed'))).toEqual([
      false,
      true,
      true,
      false,
    ])
    // the folded handle swaps to ⊕
    expect(dom().querySelectorAll('.outline-handle')[0].textContent).toBe('⊕')
    mousedown(dom().querySelectorAll('.outline-handle')[0])
    expect(headings().map((h) => h.classList.contains('outline-collapsed'))).toEqual([
      false,
      false,
      false,
      false,
    ])
    expect(dom().querySelectorAll('.outline-handle')[0].textContent).toBe('⊖')
  })

  it('a leaf handle click is inert', () => {
    const editor = createEditor(OUTLINE_DOC)
    editor.commands.setOutlineViewActive(true)
    const delta = editor.view.dom.querySelectorAll('.outline-handle')[3]
    expect(delta.classList.contains('is-leaf')).toBe(true)
    mousedown(delta)
    expect(editor.view.dom.querySelectorAll('.outline-collapsed').length).toBe(0)
  })

  it('clicking a nested heading folds only its own rows', () => {
    const editor = createEditor(OUTLINE_DOC)
    editor.commands.setOutlineViewActive(true)
    // Beta (h2) has one child row: Gamma (h3)
    mousedown(editor.view.dom.querySelectorAll('h2 .outline-handle')[0])
    const headings = [...editor.view.dom.querySelectorAll('h1, h2, h3')]
    expect(headings.map((h) => h.classList.contains('outline-collapsed'))).toEqual([
      false,
      false,
      true,
      false,
    ])
  })

  it('folds survive an unrelated edit (keyed by level + text, not position)', () => {
    const editor = createEditor(OUTLINE_DOC)
    editor.commands.setOutlineViewActive(true)
    mousedown(editor.view.dom.querySelectorAll('.outline-handle')[0])
    expect(editor.view.dom.querySelectorAll('.outline-collapsed').length).toBe(2)
    editor.view.dispatch(editor.state.tr.insertText('x', 10))
    expect(editor.view.dom.querySelectorAll('.outline-collapsed').length).toBe(2)
  })

  it('leaving outline view clears the folds and the plugin state', () => {
    const editor = createEditor(OUTLINE_DOC)
    editor.commands.setOutlineViewActive(true)
    mousedown(editor.view.dom.querySelectorAll('.outline-handle')[0])
    editor.commands.setOutlineViewActive(false)
    expect(editor.view.dom.querySelectorAll('.outline-handle, .outline-collapsed').length).toBe(0)
    const state = outlineViewKey.getState(editor.state)
    expect(state?.active).toBe(false)
    expect(state?.collapsed.size).toBe(0)
    // returning to outline view starts fully expanded
    editor.commands.setOutlineViewActive(true)
    expect(editor.view.dom.querySelectorAll('.outline-collapsed').length).toBe(0)
    expect(editor.view.dom.querySelectorAll('.outline-handle').length).toBe(4)
  })
})
