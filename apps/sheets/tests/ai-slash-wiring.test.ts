// @vitest-environment jsdom
/**
 * Slash-palette wiring in the sheets AI panel. The composer here is the shared
 * AiComposer, whose textarea keydown the panel cannot hook into, so the
 * palette's keys are intercepted on the way down through .ai-composer and
 * stopped before the composer's Enter-send/Escape-stop handling. These tests
 * pin that contract against the real panel: typing "/" opens the listbox,
 * Enter picks — attaching the skill and striking the token without sending —
 * and an already-attached skill stays attached when picked again.
 *
 * Written with `createElement` rather than JSX because this package's vitest
 * include pattern only picks up `.test.ts`, so a `.tsx` here would not run.
 */
import { act, createElement, useState } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest'

import { AiChatPanel } from '../src/renderer/ai/AiChatPanel'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const SKILLS = [
  { name: 'tables', description: 'Header rows and totals' },
  { name: 'tone', description: 'Write plainly' },
]

beforeAll(() => {
  // jsdom implements no scrolling; the chat log's follow-stream effect calls scrollTo on mount
  if (!Element.prototype.scrollTo) Element.prototype.scrollTo = () => {}
})

let host: HTMLDivElement
let root: Root
let skillsList: ReturnType<typeof vi.fn>
const toggled: string[] = []
const sent: string[] = []

beforeEach(() => {
  toggled.length = 0
  sent.length = 0
  skillsList = vi.fn(async () => SKILLS)
  // the panel reads the bridge off window.desktopApi; only the calls its mount
  // and the palette make are stubbed, the rest would only run on interaction
  ;(window as unknown as { desktopApi: unknown }).desktopApi = {
    skillsList,
    getAiSettings: async () => null,
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
  delete (window as unknown as { desktopApi?: unknown }).desktopApi
})

/** the shell's slice of the panel contract: the draft and attached skills are owned above the panel */
function Panel() {
  const [prompt, setPrompt] = useState('')
  const [attached, setAttached] = useState<readonly string[]>([])
  return createElement(AiChatPanel, {
    isOpen: true,
    hasContent: true,
    chat: [],
    attachments: [],
    attachedSkills: attached,
    onToggleSkill: (name: string) => {
      toggled.push(name)
      setAttached((prev) =>
        prev.includes(name) ? prev.filter((n) => n !== name) : [...prev, name],
      )
    },
    attachNotice: null,
    onPickAttachments: () => {},
    onAddAttachmentPaths: () => {},
    onAddPastedImage: () => {},
    onRemoveAttachment: () => {},
    prompt,
    onPromptChange: setPrompt,
    preview: null,
    aiBusy: false,
    onSend: (instruction?: string) => {
      sent.push(instruction ?? prompt)
    },
    onStop: () => {},
    onNewChat: () => {},
    onUndo: () => {},
    scopeRange: null,
    scopeColumns: null,
    scopeLocked: false,
    onScopeDismiss: () => {},
    onCitation: () => {},
    onExpand: () => {},
    onCollapse: () => {},
  })
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

describe('AiChatPanel slash palette', () => {
  it('opens the skill listbox above the composer for a leading slash token', async () => {
    act(() => root.render(createElement(Panel)))
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
    act(() => root.render(createElement(Panel)))
    await settle()
    type('please /tone')
    await settle()
    press('Enter')
    await settle()
    expect(toggled).toEqual(['tone'])
    expect(host.querySelector('textarea')!.value).toBe('please ')
    expect(sent).toEqual([])
    expect(pop()).toBeNull()
    // the pick refocuses the textarea and leaves the caret where the token began
    await act(async () => {
      await new Promise((resolve) => requestAnimationFrame(() => resolve(null)))
    })
    const ta = host.querySelector('textarea')!
    expect(document.activeElement).toBe(ta)
    expect(ta.selectionStart).toBe('please '.length)
  })

  it('picking an already-attached skill keeps it attached', async () => {
    act(() => root.render(createElement(Panel)))
    await settle()
    type('/tone')
    await settle()
    press('Enter')
    await settle()
    type('/tone')
    await settle()
    press('Enter')
    await settle()
    // one attach; a blind toggle would have detached it on the second pick
    expect(toggled).toEqual(['tone'])
    expect(host.querySelector('textarea')!.value).toBe('')
  })

  it('still sends on Enter when no slash token is live', async () => {
    act(() => root.render(createElement(Panel)))
    await settle()
    type('hello')
    await settle()
    press('Enter')
    expect(sent).toEqual(['hello'])
  })
})
