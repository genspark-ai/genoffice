import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentStreamCallbacks, IpcStreamChunk } from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { createElectronTransport, withReasoningTap } from '../src/renderer/ai/transport'

const noTools = { onToolCall: () => {} }

describe('withReasoningTap', () => {
  it('taps reasoning deltas while the loop keeps its own copy', () => {
    const tapped: string[] = []
    const loopSeen: string[] = []
    const onDelta = vi.fn()
    const callbacks: AgentStreamCallbacks = {
      onDelta,
      onDone: () => {},
      onError: () => {},
      ...noTools,
      onReasoning: (delta) => loopSeen.push(delta),
    }
    const wrapped = withReasoningTap(callbacks, (delta) => tapped.push(delta))
    wrapped.onReasoning?.('think ')
    wrapped.onReasoning?.('harder')
    expect(tapped).toEqual(['think ', 'harder'])
    expect(loopSeen).toEqual(['think ', 'harder'])
    expect(wrapped.onDelta).toBe(onDelta)
  })

  it('still taps when the loop registers no onReasoning of its own', () => {
    const tapped: string[] = []
    const wrapped = withReasoningTap(
      { onDelta: () => {}, onDone: () => {}, onError: () => {}, ...noTools },
      (delta) => tapped.push(delta),
    )
    expect(() => wrapped.onReasoning?.('x')).not.toThrow()
    expect(tapped).toEqual(['x'])
  })
})

describe('createElectronTransport reasoning forwarding', () => {
  // the transport resolves window.desktopApi lazily inside stream(); stub it for the node harness
  const g = globalThis as { window?: unknown }

  afterEach(() => {
    delete g.window
  })

  it('fires the app tap and the loop callback for reasoning chunks', () => {
    // holder object: closure assignments must not narrow the call site to null
    const sink: { push: ((chunk: IpcStreamChunk) => void) | null } = { push: null }
    let requestId = ''
    g.window = {
      desktopApi: {
        onAiStream: (listener: (chunk: IpcStreamChunk) => void) => {
          sink.push = listener
          return () => {
            sink.push = null
          }
        },
        aiStream: (request: { requestId: string }) => {
          requestId = request.requestId
        },
        aiStreamCancel: () => {},
      },
    }
    const tapped: string[] = []
    const loopSeen: string[] = []
    const transport = createElectronTransport(() => ({}) as unknown as AiSettings, {
      onReasoning: (delta) => tapped.push(delta),
    })
    transport.stream(
      { system: '', messages: [], tools: [] },
      {
        onDelta: () => {},
        onDone: () => {},
        onError: () => {},
        ...noTools,
        onReasoning: (delta) => loopSeen.push(delta),
      },
    )
    sink.push?.({ requestId, type: 'reasoning', text: 'step 1' })
    sink.push?.({ requestId, type: 'delta', text: 'answer' })
    sink.push?.({ requestId, type: 'reasoning', text: ' step 2' })
    sink.push?.({ requestId, type: 'done' })
    expect(tapped).toEqual(['step 1', ' step 2'])
    expect(loopSeen).toEqual(['step 1', ' step 2'])
  })
})
