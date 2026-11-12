/**
 * Turning off Settings > "show the AI-edit chip on selection" also dismisses a
 * popover the chip opened in the Markdown window (Docs parity).
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { Editor } from '@tiptap/core'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { applyAiPanelPrefs } from '@genoffice/ui'
import { AiAskPopover } from '../src/renderer/components/AiAskPopover'

let coords: { left: number; right: number; top: number; bottom: number }
let listeners: Map<string, () => void>
let noop: () => void

function makeEditor(): Editor {
  const tr: Record<string, unknown> = {}
  tr.setMeta = () => tr
  const state = {
    selection: { empty: false, from: 1, to: 5, $from: { depth: 0 } },
    doc: { textBetween: () => '\u673aIF\u6539' },
    tr,
  }
  const listenersMap = listeners
  return {
    state,
    isFocused: true,
    isDestroyed: false,
    view: {
      coordsAtPos: () => coords,
      dispatch: noop,
    },
    on: (event: string, handler: () => void) => listenersMap.set(event, handler),
    off: (event: string) => listenersMap.delete(event),
  } as unknown as Editor
}

let host: HTMLDivElement
let root: Root
let editor: Editor

const chip = (): HTMLButtonElement | null =>
  host.querySelector<HTMLButtonElement>('.ai-ask-trigger')

function mount(): void {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
  editor = makeEditor()
  act(() =>
    root.render(
      createElement(AiAskPopover, {
        editor,
        queueFull: false,
        getItem: () => undefined,
        onSendNow: noop,
        onQueueAdd: noop,
        onQueueUpdate: noop,
        onQueueRemove: noop,
      }),
    ),
  )
}

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  applyAiPanelPrefs({})
  coords = { left: 10, right: 20, top: 30, bottom: 45 }
  listeners = new Map()
  noop = () => {}
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  applyAiPanelPrefs({})
  vi.unstubAllGlobals()
})

describe('AiAskPopover selection trigger switch (markdown)', () => {
  it('dismisses a chip-opened popover when the switch is turned off', () => {
    mount()
    act(() => chip()!.click())
    expect(host.querySelector('.ai-ask-pop')).not.toBeNull()
    act(() => applyAiPanelPrefs({ showAiAskTrigger: false }))
    expect(host.querySelector('.ai-ask-pop')).toBeNull()
  })
})
