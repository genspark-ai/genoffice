import { act } from 'react'
import type { ReactElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { SendRequest } from '@genoffice/mail-protocols'
import { ComposeHost } from '../src/renderer/compose/ComposeHost'
import type { ComposeHostProps } from '../src/renderer/compose/ComposeHost'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { FakeMailApi } from '../src/renderer/devFakes'

/**
 * Lane C compose tests over IPC fakes: draft restore, the unsent-draft
 * dirty guard, two-step send with the data-fallback attachment path,
 * client-side oversize rejection, and read-only mode.
 */

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

let containers: HTMLElement[] = []

function render(node: ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  containers.push(container)
  const root = createRoot(container)
  act(() => {
    root.render(<LocaleProvider initial="en">{node}</LocaleProvider>)
  })
  return container
}

async function settle(rounds = 12): Promise<void> {
  await act(async () => {
    for (let i = 0; i < rounds; i += 1) {
      await Promise.resolve()
    }
  })
}

function buttonByText(container: HTMLElement, text: string): HTMLButtonElement | undefined {
  return [...container.querySelectorAll('button')].find((button) => button.textContent === text) as
    HTMLButtonElement | undefined
}

function setInputValue(input: HTMLInputElement, value: string): void {
  const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set
  if (setter) setter.call(input, value)
  else input.value = value
  input.dispatchEvent(new Event('input', { bubbles: true }))
}

function setBody(container: HTMLElement, html: string): void {
  const body = container.querySelector('.mail-compose-body')
  if (!body) throw new Error('missing compose body')
  body.innerHTML = html
  body.dispatchEvent(new Event('input', { bubbles: true }))
}

function mountCompose(over?: Partial<ComposeHostProps>): {
  container: HTMLElement
  onDirtyChange: ReturnType<typeof vi.fn>
  onSend: ReturnType<typeof vi.fn>
  api: FakeMailApi
} {
  const api = new FakeMailApi()
  ;(window as unknown as { mailApi: FakeMailApi }).mailApi = api
  const onDirtyChange = vi.fn()
  const onSend = vi.fn()
  const container = render(
    <ComposeHost
      accountId="acc-1"
      threadKey="thread-1"
      onDirtyChange={onDirtyChange}
      onSend={onSend}
      {...over}
    />,
  )
  return { container, onDirtyChange, onSend, api }
}

beforeEach(() => {
  ;(window as unknown as { mailCompose?: unknown }).mailCompose = undefined
})

afterEach(() => {
  for (const container of containers) container.remove()
  containers = []
})

describe('ComposeHost', () => {
  it('starts clean and restores an initial draft as dirty', async () => {
    const clean = mountCompose()
    await settle()
    expect(clean.onDirtyChange).toHaveBeenCalledWith(false)

    const restored = mountCompose({ initialDraft: 'hello draft' })
    await settle()
    expect(restored.container.textContent).toContain('Draft restored.')
    expect(restored.onDirtyChange).toHaveBeenCalledWith(true)
  })

  it('marks dirty on typing and autosaves the draft', async () => {
    const { container, onDirtyChange, api } = mountCompose()
    await settle()
    const spy = vi.spyOn(api, 'saveDraft')
    const to = container.querySelector('#mail-compose-to')
    if (!(to instanceof HTMLInputElement)) throw new Error('missing to field')
    await act(async () => {
      setInputValue(to, 'a@example.com')
    })
    setBody(container, 'hello')
    await settle()
    expect(onDirtyChange).toHaveBeenCalledWith(true)
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 1000))
    })
    await settle()
    expect(spy).toHaveBeenCalledWith('acc-1', 'thread-1', expect.stringContaining('a@example.com'))
    expect(onDirtyChange).toHaveBeenLastCalledWith(false)
  })

  it('blocks send without a valid recipient', async () => {
    const { container, onSend } = mountCompose()
    await settle()
    const send = buttonByText(container, 'Send')
    expect(send?.disabled).toBe(true)
    await act(async () => {
      send?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await settle()
    expect(onSend).not.toHaveBeenCalled()
  })

  it('sends in two steps with a SendRequest payload', async () => {
    const { container, onSend, onDirtyChange } = mountCompose()
    await settle()
    const to = container.querySelector('#mail-compose-to')
    const subject = container.querySelector('#mail-compose-subject')
    if (!(to instanceof HTMLInputElement) || !(subject instanceof HTMLInputElement)) {
      throw new Error('missing compose fields')
    }
    await act(async () => {
      setInputValue(to, 'a@example.com, b@example.com')
      setInputValue(subject, 'hello')
    })
    setBody(container, 'see <b>you</b>')
    await settle()
    const send = buttonByText(container, 'Send')
    expect(send?.disabled).toBe(false)
    await act(async () => {
      send?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await settle()
    expect(buttonByText(container, 'Send this message?')).not.toBeUndefined()
    expect(onSend).not.toHaveBeenCalled()
    const confirm = buttonByText(container, 'Send this message?')
    await act(async () => {
      confirm?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await settle()
    expect(onSend).toHaveBeenCalledTimes(1)
    const req = onSend.mock.calls[0]?.[0] as SendRequest
    expect(req.to).toEqual(['a@example.com', 'b@example.com'])
    expect(req.subject).toBe('hello')
    expect(req.text).toContain('see you')
    expect(onDirtyChange).toHaveBeenLastCalledWith(false)
  })

  it('carries attachments on the data-fallback path', async () => {
    const { container, onSend } = mountCompose()
    await settle()
    const to = container.querySelector('#mail-compose-to')
    if (!(to instanceof HTMLInputElement)) throw new Error('missing to field')
    await act(async () => {
      setInputValue(to, 'a@example.com')
    })
    const picker = container.querySelector('.mail-compose-attach input[type="file"]')
    if (!(picker instanceof HTMLInputElement)) throw new Error('missing file picker')
    const file = new File(['hello'], 'hi.txt', { type: 'text/plain' })
    Object.defineProperty(picker, 'files', { configurable: true, value: [file] })
    await act(async () => {
      picker.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await settle()
    expect(container.textContent).toContain('hi.txt')
    const send = buttonByText(container, 'Send')
    await act(async () => {
      send?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await settle()
    const confirm = buttonByText(container, 'Send this message?')
    await act(async () => {
      confirm?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await settle()
    expect(onSend).toHaveBeenCalledTimes(1)
    const req = onSend.mock.calls[0]?.[0] as SendRequest
    expect(req.attachments).toHaveLength(1)
    expect(req.attachments?.[0].filename).toBe('hi.txt')
    expect(req.attachments?.[0].path.startsWith('memory:')).toBe(true)
    expect(req.attachments?.[0].data).toBeInstanceOf(Uint8Array)
  })

  it('stages through the compose bridge when present', async () => {
    const { container, onSend } = mountCompose()
    await settle()
    ;(window as unknown as { mailCompose?: unknown }).mailCompose = {
      stageAttachment: async (filename: string) => ({
        filename,
        mime: 'text/plain',
        size: 5,
        path: '/staged/abc-hi.txt',
      }),
      outboxList: async () => [],
      openExternal: async () => ({ opened: true }),
    }
    const to = container.querySelector('#mail-compose-to')
    if (!(to instanceof HTMLInputElement)) throw new Error('missing to field')
    await act(async () => {
      setInputValue(to, 'a@example.com')
    })
    const picker = container.querySelector('.mail-compose-attach input[type="file"]')
    if (!(picker instanceof HTMLInputElement)) throw new Error('missing file picker')
    Object.defineProperty(picker, 'files', {
      configurable: true,
      value: [new File(['hello'], 'hi.txt', { type: 'text/plain' })],
    })
    await act(async () => {
      picker.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await settle()
    const send = buttonByText(container, 'Send')
    await act(async () => {
      send?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await settle()
    const confirm = buttonByText(container, 'Send this message?')
    await act(async () => {
      confirm?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await settle()
    const req = onSend.mock.calls[0]?.[0] as SendRequest
    expect(req.attachments?.[0].path).toBe('/staged/abc-hi.txt')
    expect(req.attachments?.[0]).not.toHaveProperty('data')
  })

  it('rejects oversize files client-side', async () => {
    const { container, onSend } = mountCompose()
    await settle()
    const picker = container.querySelector('.mail-compose-attach input[type="file"]')
    if (!(picker instanceof HTMLInputElement)) throw new Error('missing file picker')
    Object.defineProperty(picker, 'files', {
      configurable: true,
      value: [new File([new Uint8Array(51 * 1024 * 1024)], 'huge.bin')],
    })
    await act(async () => {
      picker.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await settle()
    expect(container.textContent).toContain('Attachments exceed the size cap.')
    expect(container.querySelectorAll('.mail-compose-attach-row')).toHaveLength(0)
    expect(onSend).not.toHaveBeenCalled()
  })

  it('disables editing in read-only mode', async () => {
    const { container } = mountCompose({ readOnly: true })
    await settle()
    const to = container.querySelector('#mail-compose-to')
    expect(to?.hasAttribute('disabled')).toBe(true)
    expect(buttonByText(container, 'Send')?.disabled).toBe(true)
  })

  it('populates the body DOM from an initial draft', async () => {
    const { container } = mountCompose({ initialDraft: 'hello draft' })
    await settle()
    const body = container.querySelector('.mail-compose-body')
    if (!(body instanceof HTMLElement)) throw new Error('missing compose body')
    expect(body.innerHTML).toContain('hello draft')
  })

  it('keeps typed body order across re-renders (caret regression)', async () => {
    const { container } = mountCompose()
    await settle()
    const body = container.querySelector('.mail-compose-body')
    if (!(body instanceof HTMLElement)) throw new Error('missing compose body')
    // Simulate human keystrokes: the browser appends each char at the caret
    // (end) and fires input; every setState re-renders. The old
    // dangerouslySetInnerHTML rewrite reset the caret to offset 0, so each
    // char landed at the start ("Hello World" -> "dlroW olleH").
    for (const char of 'Hello World') {
      await act(async () => {
        body.innerHTML += char
        body.dispatchEvent(new Event('input', { bubbles: true }))
      })
      await settle()
    }
    expect(body.innerHTML).toBe('Hello World')
  })

  it('does not rewrite the body DOM on unrelated re-renders', async () => {
    const { container } = mountCompose()
    await settle()
    setBody(container, 'hello')
    await settle()
    const body = container.querySelector('.mail-compose-body')
    if (!(body instanceof HTMLElement)) throw new Error('missing compose body')
    // An in-progress edit the state mirror has not seen yet (the caret is
    // mid-text in a real browser): re-renders from other fields must not
    // clobber it back to the last mirrored state.
    body.innerHTML = 'hello + unsynced edit'
    const to = container.querySelector('#mail-compose-to')
    if (!(to instanceof HTMLInputElement)) throw new Error('missing to field')
    await act(async () => {
      setInputValue(to, 'a@example.com')
    })
    await settle()
    expect(body.innerHTML).toBe('hello + unsynced edit')
  })
})
