import { describe, expect, it } from 'vitest'
import { restoreEditSelection, saveEditSelection } from '../src/renderer/TextEditOverlay'

// jsdom doesn't derive isContentEditable from the attribute
function makeEditor(text: string): HTMLDivElement {
  const div = document.createElement('div')
  div.contentEditable = 'true'
  div.tabIndex = 0
  Object.defineProperty(div, 'isContentEditable', { value: true })
  div.textContent = text
  document.body.appendChild(div)
  return div
}

function selectChars(div: HTMLElement, start: number, end: number) {
  const range = document.createRange()
  range.setStart(div.firstChild!, start)
  range.setEnd(div.firstChild!, end)
  const sel = window.getSelection()!
  sel.removeAllRanges()
  sel.addRange(range)
}

describe('edit selection handoff for ribbon controls', () => {
  it('restores the saved range and re-focuses the editor after focus moved to a control', () => {
    const div = makeEditor('AAABBB')
    div.focus()
    selectChars(div, 0, 3)
    saveEditSelection()

    const input = document.createElement('input')
    document.body.appendChild(input)
    input.focus()
    window.getSelection()!.removeAllRanges()

    expect(restoreEditSelection()).toBe(true)
    expect(document.activeElement).toBe(div)
    const r = window.getSelection()!.getRangeAt(0)
    expect(r.startContainer).toBe(div.firstChild)
    expect(r.startOffset).toBe(0)
    expect(r.endOffset).toBe(3)

    div.remove()
    input.remove()
  })

  it('restore fails once the editor is gone (overlay committed/unmounted)', () => {
    const div = makeEditor('AAABBB')
    div.focus()
    selectChars(div, 1, 4)
    saveEditSelection()
    div.remove()
    expect(restoreEditSelection()).toBe(false)
  })

  it('save is a no-op when focus is not inside a contenteditable', () => {
    const div = makeEditor('stale')
    div.focus()
    selectChars(div, 0, 2)
    saveEditSelection()
    div.remove()

    const input = document.createElement('input')
    document.body.appendChild(input)
    input.focus()
    saveEditSelection() // must not overwrite with a non-editor selection nor throw
    expect(restoreEditSelection()).toBe(false)
    input.remove()
  })
})
