/** @vitest-environment jsdom */
import { act, createElement, useRef, useState } from 'react'
import type { ChangeEvent, KeyboardEvent as ReactKeyDown } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { LANGS } from '@genoffice/i18n'
import { AI_SLASH_LABELS } from '../src/strings-ai-slash'
import { useAiSlash } from '../src/useAiSlash'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const SKILLS = [
  { name: 'tables', description: 'Header rows and totals' },
  { name: 'tone', description: 'Write plainly' },
]

let host: HTMLDivElement
let root: Root

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})
afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

async function settle(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

/**
 * The parent the hook is written for: a controlled textarea whose state it
 * owns, an Enter that sends, and an attach handler that strikes the handed
 * range out of the draft. The palette must never edit the draft itself.
 */
interface Harness {
  attach: ReturnType<typeof vi.fn>
  send: ReturnType<typeof vi.fn>
  /** what the last keydown returned: true means the palette ate the key */
  consumed: boolean | null
  type: (text: string, caret?: number) => void
  press: (key: string) => KeyboardEvent
  clickCaret: (caret: number) => void
  pop: () => HTMLElement | null
  textarea: () => HTMLTextAreaElement
}

function mount(listSkills = vi.fn(async () => SKILLS)): Harness {
  const attach = vi.fn()
  const send = vi.fn()
  const h: Harness = {
    attach,
    send,
    consumed: null,
    type(text, caret = text.length) {
      const ta = host.querySelector('textarea')!
      act(() => {
        const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
        setter.call(ta, text)
        ta.setSelectionRange(caret, caret)
        ta.dispatchEvent(new Event('input', { bubbles: true }))
      })
    },
    press(key) {
      const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true })
      act(() => host.querySelector('textarea')!.dispatchEvent(event))
      return event
    },
    clickCaret(caret) {
      const ta = host.querySelector('textarea')!
      act(() => {
        ta.setSelectionRange(caret, caret)
        ta.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      })
    },
    pop: () => host.querySelector<HTMLElement>('.ai-slash-pop'),
    textarea: () => host.querySelector('textarea')!,
  }

  function Composer() {
    const [value, setValue] = useState('')
    const inputRef = useRef<HTMLTextAreaElement | null>(null)
    const slash = useAiSlash({
      lang: 'en',
      listSkills,
      inputRef,
      value,
      onAttachSkill: (name, range) => {
        attach(name, range)
        setValue(value.slice(0, range.start) + value.slice(range.end))
      },
    })
    return createElement(
      'div',
      null,
      slash.element,
      createElement('textarea', {
        ref: inputRef,
        value,
        onChange: (e: ChangeEvent<HTMLTextAreaElement>) => setValue(e.target.value),
        onKeyDown: (e: ReactKeyDown<HTMLTextAreaElement>) => {
          h.consumed = slash.onKeyDown(e)
          if (h.consumed) return
          if (e.key === 'Enter' && !e.shiftKey) send()
        },
      }),
    )
  }

  act(() => root.render(createElement(Composer)))
  return h
}

const names = (h: Harness) =>
  [...(h.pop()?.querySelectorAll('.ai-slash-skill-name') ?? [])].map((n) => n.textContent)
const selected = (h: Harness) =>
  h.pop()?.querySelector('[aria-selected="true"] .ai-slash-skill-name')?.textContent

describe('useAiSlash', () => {
  it('labels every language the skills strings cover', () => {
    expect(Object.keys(AI_SLASH_LABELS).sort()).toEqual([...LANGS].sort())
  })

  it('opens above the composer on a leading slash and lazy-loads the skills', async () => {
    const listSkills = vi.fn(async () => SKILLS)
    const h = mount(listSkills)
    expect(h.pop()).toBeNull()
    expect(listSkills).not.toHaveBeenCalled()
    h.type('/t')
    expect(listSkills).toHaveBeenCalledTimes(1)
    await settle()
    const pop = h.pop()!
    expect(pop.getAttribute('role')).toBe('listbox')
    // pinned above the textarea the way the "+" menu is: innerHeight - top + 6
    expect(pop.style.bottom).toBe(`${window.innerHeight + 6}px`)
    expect(names(h)).toEqual(['tables', 'tone'])
    expect(pop.querySelector('.ai-slash-avatar')?.textContent).toBe('T')
    expect(selected(h)).toBe('tables')
  })

  it('narrows the rows as the query grows', async () => {
    const h = mount()
    h.type('/wri')
    await settle()
    expect(names(h)).toEqual(['tone'])
  })

  it('moves the active row with the arrows and picks with Enter, striking the whole token', async () => {
    const h = mount()
    h.type('/t')
    await settle()
    h.press('ArrowDown')
    expect(selected(h)).toBe('tone')
    h.press('ArrowUp')
    expect(selected(h)).toBe('tables')
    h.press('ArrowDown')
    expect(selected(h)).toBe('tone')
    const event = h.press('Enter')
    expect(event.defaultPrevented).toBe(true)
    expect(h.consumed).toBe(true)
    expect(h.attach).toHaveBeenCalledWith('tone', { start: 0, end: 2 })
    expect(h.textarea().value).toBe('')
    expect(h.send).not.toHaveBeenCalled()
    expect(h.pop()).toBeNull()
  })

  it('picks with Tab and with a mouse click, striking a token mid-draft', async () => {
    const h = mount()
    h.type('please /tone')
    await settle()
    h.press('Tab')
    expect(h.attach).toHaveBeenCalledWith('tone', { start: 7, end: 12 })
    expect(h.textarea().value).toBe('please ')
    expect(h.pop()).toBeNull()

    h.type('run /tab')
    await settle()
    act(() => h.pop()!.querySelector('.ai-slash-skill')!.click())
    expect(h.attach).toHaveBeenCalledWith('tables', { start: 4, end: 8 })
    expect(h.textarea().value).toBe('run ')
  })

  it('closes on Escape and stays closed while the same token lives on', async () => {
    const h = mount()
    h.type('/t')
    await settle()
    const event = h.press('Escape')
    expect(event.defaultPrevented).toBe(true)
    expect(h.pop()).toBeNull()
    h.type('/to')
    expect(h.pop()).toBeNull()
    h.type('')
    h.type('/')
    await settle()
    expect(h.pop()).not.toBeNull()
  })

  it('lets Enter fall through to send when nothing matches', async () => {
    const h = mount()
    h.type('/zebra')
    await settle()
    const event = h.press('Enter')
    expect(event.defaultPrevented).toBe(false)
    expect(h.consumed).toBe(false)
    expect(h.send).toHaveBeenCalledTimes(1)
    expect(h.attach).not.toHaveBeenCalled()
  })

  it('returns false while closed so the composer keeps its own keys', async () => {
    const h = mount()
    h.type('hello')
    await settle()
    h.press('Enter')
    expect(h.consumed).toBe(false)
    expect(h.send).toHaveBeenCalledTimes(1)
    expect(h.attach).not.toHaveBeenCalled()
  })

  it('opens only for a slash at the start or after whitespace, and a space ends the token', async () => {
    const h = mount()
    h.type('plain/text')
    await settle()
    expect(h.pop()).toBeNull()
    h.type('hi /to')
    await settle()
    expect(h.pop()).not.toBeNull()
    h.type('hi /to now')
    await settle()
    expect(h.pop()).toBeNull()
  })

  it('distinguishes no skills installed from no match', async () => {
    const empty = mount(vi.fn(async () => []))
    empty.type('/')
    await settle()
    expect(empty.pop()?.querySelector('.ai-slash-empty')?.textContent).toBe(
      'No skills installed yet',
    )
    const h = mount()
    h.type('/zebra')
    await settle()
    expect(h.pop()?.querySelector('.ai-slash-empty')?.textContent).toBe('No matching skills')
  })

  it('dismisses on an outside press and reopens when the caret re-enters a token', async () => {
    const h = mount()
    h.type('/to')
    await settle()
    expect(h.pop()).not.toBeNull()
    act(() => document.body.dispatchEvent(new Event('pointerdown', { bubbles: true })))
    expect(h.pop()).toBeNull()
    h.clickCaret(0)
    expect(h.pop()).toBeNull()
    h.clickCaret(4)
    await settle()
    expect(h.pop()).not.toBeNull()
  })

  it('refreshes the skill list on every open', async () => {
    const listSkills = vi.fn(async () => SKILLS)
    const h = mount(listSkills)
    h.type('/')
    await settle()
    expect(listSkills).toHaveBeenCalledTimes(1)
    h.press('Escape')
    h.type('')
    h.type('/')
    await settle()
    expect(listSkills).toHaveBeenCalledTimes(2)
    expect(names(h)).toEqual(['tables', 'tone'])
  })
})
