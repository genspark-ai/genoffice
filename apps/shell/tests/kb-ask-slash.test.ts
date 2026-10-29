// @vitest-environment jsdom
/**
 * Slash-palette wiring in the home knowledge-base ask panel. The composer is
 * the shared AiComposer, whose textarea keydown the panel cannot hook into,
 * so the palette's keys are gated on the way down through the panel root and
 * stopped before the composer's Enter-send/Escape-stop handling. These tests
 * pin that contract against the real panel: typing "/" opens the listbox,
 * Enter picks — attaching the skill and striking the token without sending —
 * and an already-attached skill stays attached when picked again.
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'
import { KbAskPanel } from '../src/renderer/src/KbAskPanel'
import { LocaleProvider } from '../src/renderer/src/locale'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const SKILLS = [
  { name: 'tables', description: 'Header rows and totals' },
  { name: 'tone', description: 'Write plainly' },
]

let host: HTMLDivElement
let root: Root
let skillsList: ReturnType<typeof vi.fn>

beforeAll(() => {
  // jsdom implements no scrolling; the thread's follow-stream effect calls scrollTo
  if (!Element.prototype.scrollTo) Element.prototype.scrollTo = () => {}
})

beforeEach(() => {
  skillsList = vi.fn(async () => SKILLS)
  // the panel tolerates a bridge without kb:* channels (defensive branch in
  // its mount effect), which also skips the agent loop: only the calls its
  // mount and the palette make are stubbed, the rest would only run on use
  ;(window as unknown as { aiOffice: unknown }).aiOffice = {
    skillsList,
    aiGetSettings: async () => null,
    aiGskStatus: async () => ({ loggedIn: true }),
    onAiSettingsChanged: () => () => {},
  }
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  delete (window as unknown as { aiOffice?: unknown }).aiOffice
})

function render(): void {
  act(() =>
    root.render(
      createElement(
        LocaleProvider,
        { initial: 'en' },
        createElement(KbAskPanel, {
          onOpenPath: () => {},
          onGoToRecents: () => {},
          onOpenModelSettings: () => {},
        }),
      ),
    ),
  )
}

async function settle(): Promise<void> {
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

function type(text: string): void {
  const ta = host.querySelector('textarea')!
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
    setter.call(ta, text)
    ta.setSelectionRange(text.length, text.length)
    ta.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function press(key: string): void {
  act(() => {
    host
      .querySelector('textarea')!
      .dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
  })
}

const pop = (): HTMLElement | null => host.querySelector<HTMLElement>('.ai-slash-pop')
const names = (): (string | null)[] =>
  [...(pop()?.querySelectorAll('.ai-slash-skill-name') ?? [])].map((n) => n.textContent)
const chips = (): (string | null)[] =>
  [...host.querySelectorAll('.ai-skill-chip-name')].map((n) => n.textContent)

describe('KbAskPanel slash palette', () => {
  it('opens the skill listbox above the composer for a leading slash token', async () => {
    render()
    await settle()
    expect(pop()).toBeNull()
    expect(skillsList).not.toHaveBeenCalled()
    type('/t')
    await settle()
    expect(skillsList).toHaveBeenCalledTimes(1)
    expect(pop()!.getAttribute('role')).toBe('listbox')
    expect(names()).toEqual(['tables', 'tone'])
  })

  it('picks with Enter: attaches the skill, strikes the token and keeps Enter from sending', async () => {
    render()
    await settle()
    type('please /tone')
    await settle()
    press('Enter')
    await settle()
    expect(host.querySelector('textarea')!.value).toBe('please ')
    expect(chips()).toEqual(['tone'])
    expect(pop()).toBeNull()
    // nothing was asked: no message entered the thread
    expect(host.querySelectorAll('.kb-ask-msg').length).toBe(0)
    // the pick refocuses the textarea and leaves the caret where the token began
    await act(async () => {
      await new Promise((resolve) => requestAnimationFrame(() => resolve(null)))
    })
    const ta = host.querySelector('textarea')!
    expect(document.activeElement).toBe(ta)
    expect(ta.selectionStart).toBe('please '.length)
  })

  it('picking an already-attached skill keeps it attached once', async () => {
    render()
    await settle()
    type('/tone')
    await settle()
    press('Enter')
    await settle()
    type('/tone')
    await settle()
    press('Enter')
    await settle()
    // one chip; a blind toggle would have detached it on the second pick
    expect(chips()).toEqual(['tone'])
    expect(host.querySelector('textarea')!.value).toBe('')
  })

  it('closes on Escape with the draft and thread untouched', async () => {
    render()
    await settle()
    type('/tone')
    await settle()
    press('Escape')
    await settle()
    expect(pop()).toBeNull()
    expect(host.querySelector('textarea')!.value).toBe('/tone')
    expect(host.querySelectorAll('.kb-ask-msg').length).toBe(0)
  })
})
