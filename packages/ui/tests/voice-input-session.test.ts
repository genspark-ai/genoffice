/** @vitest-environment jsdom */
import { act, createElement, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { useVoiceInput, type VoiceTranscriber } from '../src/use-voice-input'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

/** Enough of MediaRecorder for the session machine; the blob is never decoded. */
class FakeMediaRecorder {
  static last: FakeMediaRecorder | null = null
  state: 'inactive' | 'recording' = 'inactive'
  mimeType = 'audio/webm'
  ondataavailable: ((event: { data: Blob }) => void) | null = null
  onstop: (() => void) | null = null

  constructor(_stream: unknown) {
    FakeMediaRecorder.last = this
  }

  start(_timeslice?: number): void {
    this.state = 'recording'
  }

  stop(): void {
    if (this.state === 'inactive') return
    this.state = 'inactive'
    this.ondataavailable?.({ data: new Blob(['pcm'], { type: 'audio/webm' }) })
    this.onstop?.()
  }
}

function Harness({
  initial,
  transcribe,
  request,
}: {
  initial: string
  transcribe: VoiceTranscriber
  request: () => Promise<boolean>
}) {
  const [draft, setDraft] = useState(initial)
  const ref = useRef<HTMLTextAreaElement | null>(null)
  const voice = useVoiceInput({
    draft,
    onChange: setDraft,
    textareaRef: ref,
    transcribe,
    requestMicrophoneAccess: request,
  })
  return createElement(
    'div',
    null,
    createElement('textarea', { id: 'draft', ref, value: draft, readOnly: true }),
    createElement('span', { id: 'phase' }, voice.phase),
    createElement('span', { id: 'error' }, voice.error ?? ''),
    createElement('button', { id: 'toggle', onClick: voice.toggle }),
    createElement('button', { id: 'discard', onClick: voice.discard }),
    createElement('button', { id: 'type', onClick: () => setDraft(`${draft}!`) }),
  )
}

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>

beforeEach(() => {
  FakeMediaRecorder.last = null
  vi.stubGlobal('MediaRecorder', FakeMediaRecorder)
  vi.stubGlobal('requestAnimationFrame', (callback: () => void) => {
    callback()
    return 0
  })
  Object.defineProperty(navigator, 'mediaDevices', {
    configurable: true,
    value: { getUserMedia: async () => ({ getTracks: () => [{ stop: () => {} }] }) },
  })
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.unstubAllGlobals()
})

const text = (id: string): string => host.querySelector(`#${id}`)?.textContent ?? ''
const draft = (): string => (host.querySelector('#draft') as HTMLTextAreaElement).value
const click = async (id: string): Promise<void> => {
  await act(async () => {
    ;(host.querySelector(`#${id}`) as HTMLButtonElement).click()
  })
}

const render = (element: ReactElement): void => {
  act(() => root.render(element))
}

/** jsdom parks a textarea's caret at 0; a session starts wherever the user left it. */
const placeCaret = (at: number): void => {
  ;(host.querySelector('#draft') as HTMLTextAreaElement).setSelectionRange(at, at)
}

describe('useVoiceInput', () => {
  it('replaces its interim transcript with the final one', async () => {
    const transcribe: VoiceTranscriber = async (_audio, { onPartial }) => {
      onPartial?.('hi')
      return 'hi there'
    }
    render(createElement(Harness, { initial: 'say: ', transcribe, request: async () => true }))
    placeCaret(5)

    await click('toggle')
    expect(text('phase')).toBe('recording')

    await click('toggle')
    expect(draft()).toBe('say: hi there')
    expect(text('phase')).toBe('idle')
    expect(text('error')).toBe('')
  })

  it('reports a refused microphone and never opens a session', async () => {
    render(
      createElement(Harness, {
        initial: 'say: ',
        transcribe: async () => 'unused',
        request: async () => false,
      }),
    )

    await click('toggle')
    expect(text('phase')).toBe('idle')
    expect(text('error')).toBe('denied')
    expect(FakeMediaRecorder.last).toBeNull()
  })

  it('takes the interim transcript back out when the session is discarded', async () => {
    const gate: { release?: (value: string) => void } = {}
    const transcribe: VoiceTranscriber = (_audio, { onPartial }) =>
      new Promise<string>((resolve) => {
        onPartial?.('half')
        gate.release = resolve
      })
    render(createElement(Harness, { initial: 'say: ', transcribe, request: async () => true }))
    placeCaret(5)

    await click('toggle')
    await click('toggle')
    expect(text('phase')).toBe('transcribing')
    expect(draft()).toBe('say: half')

    await click('discard')
    expect(draft()).toBe('say: ')
    expect(text('phase')).toBe('idle')

    gate.release?.('half')
  })

  it('ends the session and keeps the interim transcript when the user types', async () => {
    const steps: Array<() => void> = []
    const transcribe: VoiceTranscriber = (_audio, { onPartial }) =>
      new Promise<string>((resolve) => {
        steps.push(
          () => onPartial?.('hi'),
          () => onPartial?.('hi there'),
          () => resolve('hi there'),
        )
      })
    render(createElement(Harness, { initial: 'say: ', transcribe, request: async () => true }))
    placeCaret(5)

    await click('toggle')
    await click('toggle')
    act(() => steps[0]?.())
    expect(draft()).toBe('say: hi')

    await click('type')
    expect(draft()).toBe('say: hi!')

    act(() => steps[1]?.())
    expect(text('phase')).toBe('idle')
    expect(draft()).toBe('say: hi!')
  })

  it('keeps the typed draft when the final transcript arrives after the user types', async () => {
    const gate: { release?: (value: string) => void } = {}
    const transcribe: VoiceTranscriber = (_audio, { onPartial }) =>
      new Promise<string>((resolve) => {
        onPartial?.('hi')
        gate.release = resolve
      })
    render(createElement(Harness, { initial: 'say: ', transcribe, request: async () => true }))
    placeCaret(5)

    await click('toggle')
    await click('toggle')
    expect(draft()).toBe('say: hi')

    await click('type')
    expect(draft()).toBe('say: hi!')

    await act(async () => {
      gate.release?.('hi there')
    })
    expect(text('phase')).toBe('idle')
    expect(draft()).toBe('say: hi!')
  })

  it('closes the microphone when the recorder cannot be constructed', async () => {
    const stop = vi.fn()
    Object.defineProperty(navigator, 'mediaDevices', {
      configurable: true,
      value: { getUserMedia: async () => ({ getTracks: () => [{ stop }] }) },
    })
    vi.stubGlobal(
      'MediaRecorder',
      class {
        constructor() {
          throw new Error('no recorder')
        }
      },
    )
    render(
      createElement(Harness, {
        initial: 'say: ',
        transcribe: async () => 'x',
        request: async () => true,
      }),
    )

    await click('toggle')
    expect(text('phase')).toBe('idle')
    expect(text('error')).toBe('unavailable')
    expect(stop).toHaveBeenCalled()
  })
})
