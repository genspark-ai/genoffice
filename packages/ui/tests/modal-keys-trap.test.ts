/** @vitest-environment jsdom */
import { describe, expect, it } from 'vitest'
import { trapTab } from '../src/modal-keys'

const tab = (shiftKey: boolean) => ({ key: 'Tab', shiftKey, preventDefault: () => {} })

/** A modal backdrop that can hold focus itself, like the ImageViewer mask. */
function modal(): HTMLElement {
  document.body.innerHTML = ''
  const backdrop = document.createElement('div')
  backdrop.tabIndex = -1
  for (const name of ['First', 'Last']) {
    const button = document.createElement('button')
    button.textContent = name
    backdrop.append(button)
  }
  document.body.append(backdrop)
  return backdrop
}

describe('shared trapTab', () => {
  it('wraps backwards from the backdrop itself to the last control', () => {
    const container = modal()
    container.focus()
    expect(document.activeElement).toBe(container)
    const e = tab(true)
    trapTab(container, e)
    expect(container.contains(document.activeElement)).toBe(true)
    expect(document.activeElement!.textContent).toBe('Last')
  })

  it('wraps forwards from the backdrop itself to the first control', () => {
    const container = modal()
    container.focus()
    const e = tab(false)
    trapTab(container, e)
    expect(document.activeElement!.textContent).toBe('First')
  })

  it('keeps Tab from escaping through an in-modal link', () => {
    const container = modal()
    const link = document.createElement('a')
    link.href = 'https://example.com'
    link.textContent = 'Link'
    container.append(link)
    const first = container.querySelector('button')!
    link.focus()
    trapTab(container, tab(false))
    expect(document.activeElement).toBe(first)
  })
})
