import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setColumnLayout } from '../src/renderer/editor/column-layout'

const editors: Editor[] = []
// The repro mixes portrait A4 (11906 twips) and landscape A4 (16838 twips),
// with 1134-twip side margins. Portrait text needs this shared-paper offset.
const pageDx = (16838 - 11906) / 30
const contentWidth = (11906 - 2 * 1134) / 15
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

function openEditor(texts = ['first paragraph', 'next paragraph']) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: texts.map((text) => ({
        type: 'docParagraph',
        content: text ? [{ type: 'text', text }] : [],
      })),
    },
  })
  editors.push(editor)
  setColumnLayout(
    editor.view,
    Array.from(editor.view.dom.children, (el) => ({
      el: el as HTMLElement,
      widthPx: contentWidth,
      contentWPx: contentWidth,
      marginLeftPx: 75.6,
      marginRightPx: 75.6,
      pageDx,
      dx: 0,
      dy: 0,
    })),
  )
  return editor
}

function expectPageFrame(el: Element) {
  expect(el.classList.contains('doc-col-block')).toBe(true)
  expect((el as HTMLElement).style.getPropertyValue('--page-cx')).toBe('164.4px')
  expect((el as HTMLElement).style.getPropertyValue('--col-w')).toBe('642.53px')
}

describe('column and mixed-paper layout between an edit and the next pagination pass', () => {
  it.each([0, 5, 15])('keeps both paragraph frames when Enter splits at offset %i', (offset) => {
    const editor = openEditor()
    for (const el of Array.from(editor.view.dom.children)) expectPageFrame(el)
    editor.commands.setTextSelection(1 + offset)
    expect(editor.commands.splitBlock()).toBe(true)
    expect(editor.state.doc.childCount).toBe(3)
    // No setColumnLayout call: these styles must survive immediately, before
    // App's debounced pagination can measure and publish a replacement.
    for (const el of Array.from(editor.view.dom.children)) expectPageFrame(el)
  })

  it('keeps the frame when Enter splits a fresh empty paragraph again', () => {
    const editor = openEditor()
    editor.commands.setTextSelection(16)
    editor.commands.splitBlock()
    editor.commands.splitBlock()
    expect(editor.state.doc.childCount).toBe(4)
    for (const el of Array.from(editor.view.dom.children)) expectPageFrame(el)
  })

  it('keeps the frame across undo and redo of a paragraph split', () => {
    const editor = openEditor()
    editor.commands.setTextSelection(6)
    editor.commands.splitBlock()
    expect(editor.commands.undo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(2)
    for (const el of Array.from(editor.view.dom.children)) expectPageFrame(el)
    expect(editor.commands.redo()).toBe(true)
    expect(editor.state.doc.childCount).toBe(3)
    for (const el of Array.from(editor.view.dom.children)) expectPageFrame(el)
  })

  it('keeps the first paragraph frame when joining differently placed paragraphs', () => {
    const editor = openEditor()
    setColumnLayout(editor.view, [
      { el: editor.view.dom.children[0] as HTMLElement, widthPx: 300, dx: 40, dy: -20 },
      { el: editor.view.dom.children[1] as HTMLElement, widthPx: 500, dx: 90, dy: -50 },
    ])
    editor.view.dispatch(editor.state.tr.join(editor.state.doc.firstChild!.nodeSize))
    expect(editor.state.doc.childCount).toBe(1)
    const el = editor.view.dom.firstElementChild as HTMLElement
    expect(el.style.getPropertyValue('--col-w')).toBe('300px')
    expect(el.style.getPropertyValue('--col-dx')).toBe('40px')
    expect(el.style.getPropertyValue('--col-dy')).toBe('-20px')
  })

  it('does not transfer a deleted paragraph frame onto the following paragraph', () => {
    const editor = openEditor()
    setColumnLayout(editor.view, [
      { el: editor.view.dom.children[0] as HTMLElement, widthPx: 300, dx: 40, dy: 0 },
      { el: editor.view.dom.children[1] as HTMLElement, widthPx: 500, dx: 90, dy: 0 },
    ])
    editor.view.dispatch(editor.state.tr.delete(0, editor.state.doc.firstChild!.nodeSize))
    expect(editor.state.doc.childCount).toBe(1)
    const el = editor.view.dom.firstElementChild as HTMLElement
    expect(el.style.getPropertyValue('--col-w')).toBe('500px')
    expect(el.style.getPropertyValue('--col-dx')).toBe('90px')
  })

  it('keeps the frame across paragraph formatting and ordinary typing', () => {
    const editor = openEditor()
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, {
        ...editor.state.doc.firstChild!.attrs,
        align: 'center',
      }),
    )
    editor.commands.setTextSelection(6)
    editor.commands.insertContent('new text')
    for (const el of Array.from(editor.view.dom.children)) expectPageFrame(el)
  })

  it('lets the next pagination pass replace or clear the provisional frame', () => {
    const editor = openEditor()
    editor.commands.setTextSelection(6)
    editor.commands.splitBlock()
    setColumnLayout(editor.view, [
      { el: editor.view.dom.children[1] as HTMLElement, widthPx: 300, dx: 200, dy: -40 },
    ])
    const el = editor.view.dom.children[1] as HTMLElement
    expect(el.style.getPropertyValue('--col-w')).toBe('300px')
    expect(el.style.getPropertyValue('--col-dx')).toBe('200px')
    expect(el.style.getPropertyValue('--page-cx')).toBe('')
    setColumnLayout(editor.view, [])
    for (const node of Array.from(editor.view.dom.children)) {
      expect(node.classList.contains('doc-col-block')).toBe(false)
      expect((node as HTMLElement).style.getPropertyValue('--col-w')).toBe('')
    }
  })

  it.each([0, 5, 15])(
    'keeps a section-start clear only once after Enter at offset %i',
    (offset) => {
      const editor = openEditor()
      setColumnLayout(editor.view, [
        {
          el: editor.view.dom.children[0] as HTMLElement,
          widthPx: contentWidth,
          pageDx,
          dx: 0,
          dy: 0,
          sectionPageStart: true,
        },
      ])
      editor.commands.setTextSelection(1 + offset)
      editor.commands.splitBlock()
      expect(editor.view.dom.children[0].classList.contains('doc-sect-page-start')).toBe(true)
      expect(editor.view.dom.children[1].classList.contains('doc-sect-page-start')).toBe(false)
    },
  )

  it('leaves size-dependent vertical placement on the existing pagination path', () => {
    const editor = openEditor()
    setColumnLayout(editor.view, [
      {
        el: editor.view.dom.children[0] as HTMLElement,
        widthPx: 500,
        dx: 300,
        dy: 0,
        vertical: { mode: 'tbRl', marginBottomPx: -470, firstOnPage: true },
      },
    ])
    editor.commands.setTextSelection(6)
    editor.commands.splitBlock()
    // Vertical offsets and negative flow margins depend on each resulting
    // paragraph's measured extent; horizontal-frame recovery must not copy them.
    for (const el of Array.from(editor.view.dom.children)) {
      expect(el.classList.contains('doc-vert-first')).toBe(false)
      expect(el.classList.contains('doc-vert-block')).toBe(false)
      expect((el as HTMLElement).style.getPropertyValue('--vert-mb')).toBe('')
    }
  })

  it('does not copy a measured split-column float shape into new paragraphs', () => {
    const editor = openEditor()
    setColumnLayout(editor.view, [
      {
        el: editor.view.dom.children[0] as HTMLElement,
        widthPx: 500,
        dx: 300,
        dy: -100,
        split: { floatPx: 100, heightPx: 40, insetPx: 0 },
      },
    ])
    editor.commands.setTextSelection(6)
    editor.commands.splitBlock()
    expect(editor.view.dom.querySelectorAll('.doc-col-block')).toHaveLength(0)
    expect(editor.view.dom.querySelectorAll('.doc-col-split-float').length).toBeLessThanOrEqual(1)
  })

  it('keeps an undecorated first paragraph neutral when joining a decorated second one', () => {
    const editor = openEditor()
    // Mixed character-grid sections decorate only the blocks with nonzero
    // spacing. The preceding undecorated paragraph is a real layout choice.
    setColumnLayout(editor.view, [
      { el: editor.view.dom.children[1] as HTMLElement, charSpacePt: 2, dx: 0, dy: 0 },
    ])
    editor.view.dispatch(editor.state.tr.join(editor.state.doc.firstChild!.nodeSize))
    expect(editor.state.doc.childCount).toBe(1)
    const el = editor.view.dom.firstElementChild as HTMLElement
    expect(el.classList.contains('doc-charspace-block')).toBe(false)
    expect(el.style.getPropertyValue('--doc-char-space')).toBe('')
  })

  it('keeps an empty undecorated first paragraph neutral on join', () => {
    const editor = openEditor(['', 'next paragraph'])
    setColumnLayout(editor.view, [
      { el: editor.view.dom.children[1] as HTMLElement, charSpacePt: 2, dx: 0, dy: 0 },
    ])
    editor.view.dispatch(editor.state.tr.join(2))
    expect(editor.state.doc.textContent).toBe('next paragraph')
    const el = editor.view.dom.firstElementChild as HTMLElement
    expect(el.classList.contains('doc-charspace-block')).toBe(false)
    expect(el.style.getPropertyValue('--doc-char-space')).toBe('')
  })

  it('does not reintroduce a section-start clear when a continuation is split again', () => {
    const editor = openEditor()
    setColumnLayout(editor.view, [
      {
        el: editor.view.dom.children[0] as HTMLElement,
        widthPx: 300,
        dx: 40,
        dy: 0,
        sectionPageStart: true,
      },
    ])
    editor.commands.setTextSelection(6)
    editor.commands.splitBlock()
    editor.commands.splitBlock()
    expect(editor.state.doc.childCount).toBe(4)
    expect(editor.view.dom.querySelectorAll('.doc-sect-page-start')).toHaveLength(1)
    expect(editor.view.dom.children[0].classList.contains('doc-sect-page-start')).toBe(true)
    for (const el of Array.from(editor.view.dom.children).slice(0, 3)) {
      expect((el as HTMLElement).style.getPropertyValue('--col-w')).toBe('300px')
    }
  })

  it('keeps separate owners across multiple splits in a single transaction', () => {
    const editor = openEditor()
    setColumnLayout(editor.view, [
      { el: editor.view.dom.children[0] as HTMLElement, widthPx: 300, dx: 40, dy: 0 },
      { el: editor.view.dom.children[1] as HTMLElement, widthPx: 500, dx: 90, dy: 0 },
    ])
    const second = editor.state.doc.firstChild!.nodeSize
    const tr = editor.state.tr.split(6)
    tr.split(tr.mapping.map(second + 6))
    editor.view.dispatch(tr)
    expect(editor.state.doc.childCount).toBe(4)
    const widths = Array.from(editor.view.dom.children, (el) =>
      (el as HTMLElement).style.getPropertyValue('--col-w'),
    )
    expect(widths).toEqual(['300px', '300px', '500px', '500px'])
  })

  it('carries the containing frame through a multi-paragraph paste without duplicating its start flag', () => {
    const editor = openEditor()
    setColumnLayout(editor.view, [
      {
        el: editor.view.dom.children[0] as HTMLElement,
        widthPx: 300,
        dx: 40,
        dy: 0,
        sectionPageStart: true,
      },
      { el: editor.view.dom.children[1] as HTMLElement, widthPx: 500, dx: 90, dy: 0 },
    ])
    editor.commands.setTextSelection(6)
    editor.commands.insertContent([
      { type: 'docParagraph', content: [{ type: 'text', text: 'pasted A' }] },
      { type: 'docParagraph', content: [{ type: 'text', text: 'pasted B' }] },
    ])
    const blocks = Array.from(editor.view.dom.children) as HTMLElement[]
    expect(blocks.length).toBeGreaterThan(2)
    expect(editor.state.doc.lastChild!.textContent).toBe('next paragraph')
    for (const el of blocks.slice(0, -1)) {
      expect(el.style.getPropertyValue('--col-w')).toBe('300px')
    }
    expect(blocks.at(-1)!.style.getPropertyValue('--col-w')).toBe('500px')
    expect(editor.view.dom.querySelectorAll('.doc-sect-page-start')).toHaveLength(1)
  })

  it('does not lend removed document frames to a replacement document', () => {
    const editor = openEditor()
    editor.commands.setContent({
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'replacement' }] }],
    })
    expect(editor.state.doc.textContent).toBe('replacement')
    expect(editor.view.dom.querySelectorAll('.doc-col-block')).toHaveLength(0)
  })

  it('keeps table layout on the standard mapping path when Enter splits a cell paragraph', () => {
    const editor = openEditor()
    editor.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docTable',
          content: [
            {
              type: 'docTableRow',
              content: [
                {
                  type: 'docTableCell',
                  content: [
                    { type: 'docParagraph', content: [{ type: 'text', text: 'cell text' }] },
                  ],
                },
              ],
            },
          ],
        },
        { type: 'docParagraph', content: [{ type: 'text', text: 'after table' }] },
      ],
    })
    setColumnLayout(editor.view, [
      { el: editor.view.dom.children[0] as HTMLElement, contentWPx: 500, dx: 40, dy: 0 },
    ])
    editor.commands.setTextSelection(8)
    expect(editor.commands.splitBlock()).toBe(true)
    expect(editor.state.doc.childCount).toBe(2)
    expect(editor.state.doc.firstChild!.firstChild!.firstChild!.childCount).toBe(2)
    expect(editor.view.dom.children[0].classList.contains('doc-col-block')).toBe(true)
    expect(editor.view.dom.querySelectorAll('td .doc-col-block')).toHaveLength(0)
  })
})
