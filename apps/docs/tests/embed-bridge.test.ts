import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  DATAFLARE_EMBED_PROTOCOL,
  installDataflareEmbedBridge,
  requestDataflareParent,
} from '../src/shared/embed-bridge'

const parentWindow = {
  postMessage: vi.fn(),
}

const parentOrigin = 'http://dataflare.test'

function setEmbeddedWindow(): void {
  Object.defineProperty(window, 'parent', { configurable: true, value: parentWindow })
  Object.defineProperty(document, 'referrer', { configurable: true, value: `${parentOrigin}/ai/office` })
}

afterEach(() => {
  vi.clearAllMocks()
  Object.defineProperty(window, 'parent', { configurable: true, value: window })
  Object.defineProperty(document, 'referrer', { configurable: true, value: '' })
})

describe('Dataflare embed bridge', () => {
  it('binds requests and responses to the initialized session', async () => {
    setEmbeddedWindow()
    const onCommand = vi.fn()
    const dispose = installDataflareEmbedBridge(onCommand)

    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'command',
        payload: { type: 'init', sessionId: 'session-1', context: { documentType: 'docx' } },
      },
    }))

    const requestPromise = requestDataflareParent({
      type: 'http-request',
      requestId: 'request-1',
      sessionId: 'session-1',
      method: 'GET',
      path: '/crmapi/knowledge/office/1',
    })
    const requestEnvelope = parentWindow.postMessage.mock.calls.at(-1)?.[0] as {
      sessionId?: string
      payload?: { requestId?: string; sessionId?: string }
    }
    expect(requestEnvelope.sessionId).toBe('session-1')
    expect(requestEnvelope.payload?.sessionId).toBe('session-1')

    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'response',
        sessionId: 'wrong-session',
        payload: {
          type: 'http-response',
          requestId: 'request-1',
          sessionId: 'wrong-session',
          status: 200,
          headers: {},
          body: new ArrayBuffer(0),
        },
      },
    }))

    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'response',
        sessionId: 'session-1',
        payload: {
          type: 'http-response',
          requestId: 'request-1',
          sessionId: 'session-1',
          status: 204,
          headers: { etag: '"1"' },
          body: new ArrayBuffer(0),
        },
      },
    }))

    await expect(requestPromise).resolves.toMatchObject({ status: 204, sessionId: 'session-1' })

    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'command',
        sessionId: 'wrong-session',
        payload: { type: 'save' },
      },
    }))
    expect(onCommand).toHaveBeenCalledTimes(1)

    dispose()
  })

  it('delivers global-state-update commands only to the globalState handler and posts ready+request', () => {
    setEmbeddedWindow()
    const onCommand = vi.fn()
    const onGlobalState = vi.fn()
    const dispose = installDataflareEmbedBridge({ onCommand, onGlobalState })

    // init sets the session and is also dispatched to onCommand (kept for backward compatibility)
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'command',
        payload: { type: 'init', sessionId: 'session-2', context: { documentType: 'docx' } },
      },
    }))
    expect(onCommand).toHaveBeenCalledTimes(1)
    expect(onCommand.mock.calls[0][0].type).toBe('init')
    onCommand.mockClear()

    // global-state-update should not call onCommand (intercepted by bridge)
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'command',
        sessionId: 'session-2',
        payload: { type: 'global-state-update', state: { theme: 'dark', tenantId: 't-1' }, revision: 3 },
      },
    }))
    expect(onCommand).not.toHaveBeenCalled()
    expect(onGlobalState).toHaveBeenCalledWith({ theme: 'dark', tenantId: 't-1' }, 3)

    // wrong-session global-state-update should be ignored
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'command',
        sessionId: 'wrong-session',
        payload: { type: 'global-state-update', state: { theme: 'light' } },
      },
    }))
    expect(onGlobalState).toHaveBeenCalledTimes(1)

    // ready + global-state-request should have been posted to parent
    const postedPayloads = parentWindow.postMessage.mock.calls.map(call => call[0]?.payload?.type)
    expect(postedPayloads).toContain('ready')
    expect(postedPayloads).toContain('global-state-request')

    dispose()
  })

  it('accepts documentRevision in global-state-update without invoking onCommand', () => {
    setEmbeddedWindow()
    const onCommand = vi.fn()
    const onGlobalState = vi.fn()
    const dispose = installDataflareEmbedBridge({ onCommand, onGlobalState })

    // init sets the session
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'command',
        payload: { type: 'init', sessionId: 'session-rev', context: { documentType: 'docx' } },
      },
    }))
    onCommand.mockClear()

    // Host pushes a newer documentRevision (e.g. another user just saved)
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'command',
        sessionId: 'session-rev',
        payload: { type: 'global-state-update', state: { documentRevision: '42' }, revision: 2 },
      },
    }))
    expect(onCommand).not.toHaveBeenCalled()
    expect(onGlobalState).toHaveBeenCalledWith({ documentRevision: '42' }, 2)

    dispose()
  })

  it('rejects commands from a non-initialized state', () => {
    setEmbeddedWindow()
    const onCommand = vi.fn()
    const dispose = installDataflareEmbedBridge({ onCommand })

    // dispatching without prior init must not invoke onCommand
    window.dispatchEvent(new MessageEvent('message', {
      source: parentWindow as unknown as Window,
      origin: parentOrigin,
      data: {
        protocol: DATAFLARE_EMBED_PROTOCOL,
        kind: 'command',
        sessionId: 'never-init',
        payload: { type: 'save' },
      },
    }))
    expect(onCommand).not.toHaveBeenCalled()
    dispose()
  })
})
