// Slides: the renderer transport surfaces reasoning chunks to the app's
// thinking-block hook while keeping the AgentStreamCallbacks contract intact.
import { describe, expect, it, vi } from 'vitest'
import type { AgentStreamCallbacks } from '@genoffice/agent-core'
import { createElectronTransport } from '../src/renderer/ai/transport'
import { AI_PROVIDERS, type AiSettings } from '../src/shared/ipc'

const settings: AiSettings = {
  provider: 'anthropic',
  providers: Object.fromEntries(
    AI_PROVIDERS.map((p) => [p.id, { apiKey: '', model: p.defaultModel }]),
  ) as AiSettings['providers'],
}

/** slidesApi bridge mock: capture the stream listener so tests can push chunks */
function installSlidesApi(): { push: (chunk: Record<string, unknown>) => void } {
  let listener: ((chunk: Record<string, unknown>) => void) | null = null
  const requestIds: string[] = []
  Object.defineProperty(window, 'slidesApi', {
    configurable: true,
    value: {
      onAiStream: (l: (chunk: Record<string, unknown>) => void) => {
        listener = l
        return () => {
          listener = null
        }
      },
      aiStream: (request: { requestId: string }) => {
        requestIds.push(request.requestId)
        return Promise.resolve()
      },
      aiStreamCancel: vi.fn(),
    },
  })
  return {
    push: (chunk: Record<string, unknown>) =>
      listener?.({ requestId: requestIds.at(-1), ...chunk }),
  }
}

function makeCallbacks(onReasoning?: (text: string) => void): AgentStreamCallbacks {
  return {
    onDelta: vi.fn(),
    onToolCall: vi.fn(),
    onDone: vi.fn(),
    onError: vi.fn(),
    ...(onReasoning ? { onReasoning } : {}),
  }
}

describe('slides AI transport reasoning tap', () => {
  it('forwards reasoning chunks to the app hook', () => {
    const api = installSlidesApi()
    const appReasoning = vi.fn()
    const transport = createElectronTransport(() => settings, appReasoning)
    const cb = makeCallbacks()
    transport.stream({ system: 's', messages: [], tools: [] }, cb)

    api.push({ type: 'reasoning', text: 'because' })
    api.push({ type: 'done' })

    expect(appReasoning).toHaveBeenCalledWith('because')
    expect(cb.onDelta).not.toHaveBeenCalled()
  })

  it('keeps a consumer-provided onReasoning working without the app hook', () => {
    const api = installSlidesApi()
    const cbReasoning = vi.fn()
    const transport = createElectronTransport(() => settings)
    const cb = makeCallbacks(cbReasoning)
    transport.stream({ system: 's', messages: [], tools: [] }, cb)

    api.push({ type: 'reasoning', text: 'model thinking' })
    api.push({ type: 'done' })

    expect(cbReasoning).toHaveBeenCalledWith('model thinking')
  })

  it('does not disturb the delta stream through the wrapper', () => {
    const api = installSlidesApi()
    const appReasoning = vi.fn()
    const cbReasoning = vi.fn()
    const transport = createElectronTransport(() => settings, appReasoning)
    const cb = makeCallbacks(cbReasoning)
    transport.stream({ system: 's', messages: [], tools: [] }, cb)

    api.push({ type: 'delta', text: 'hel' })
    api.push({ type: 'delta', text: 'lo' })
    api.push({ type: 'reasoning', text: 'quietly' })
    api.push({ type: 'done' })

    expect(cb.onDelta).toHaveBeenCalledTimes(2)
    expect(cb.onDelta).toHaveBeenNthCalledWith(1, 'hel')
    expect(cb.onDelta).toHaveBeenNthCalledWith(2, 'lo')
    expect(cbReasoning).toHaveBeenCalledWith('quietly')
    expect(appReasoning).toHaveBeenCalledWith('quietly')
  })
})
