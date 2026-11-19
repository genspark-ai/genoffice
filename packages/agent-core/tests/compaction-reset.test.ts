import { describe, expect, it, vi } from 'vitest'
import {
  AgentLoop,
  type AgentSkill,
  type AgentStreamCallbacks,
  type AgentStreamRequest,
  type AgentTransport,
} from '../src'

const flush = () => new Promise<void>((resolve) => setTimeout(resolve, 0))

function controlledTransport() {
  const calls: Array<{
    request: AgentStreamRequest
    callbacks: AgentStreamCallbacks
    cancel: ReturnType<typeof vi.fn>
  }> = []
  const transport: AgentTransport = {
    stream(request, callbacks) {
      // Like the Electron transport, cancellation completes asynchronously.
      const cancel = vi.fn(() => queueMicrotask(() => callbacks.onDone()))
      calls.push({ request, callbacks, cancel })
      return { cancel }
    },
  }
  return { transport, calls }
}

async function startCompaction(advance: () => Promise<unknown> = flush) {
  const { transport, calls } = controlledTransport()
  const skill: AgentSkill = {
    id: 'test',
    systemPrompt: 'Test system prompt',
    tools: [{ name: 'read', description: 'Read', inputSchema: { type: 'object' } }],
    executeTool: () => ({ output: 'ok', summary: 'Read completed' }),
  }
  const onDone = vi.fn()
  const loop = new AgentLoop({
    transport,
    skill,
    events: { onDone },
    compaction: { maxBytes: 500, keepRecentBytes: 100 },
  })

  // Build history through public calls. The long first reply exceeds the budget;
  // a second user turn supplies a boundary at which that history can be folded.
  loop.run('Old conversation instruction')
  await advance()
  calls[0]!.callbacks.onDelta('Old answer '.repeat(80))
  calls[0]!.callbacks.onDone()
  loop.run('Recent question')
  await advance()
  calls[1]!.callbacks.onDelta('Recent answer')
  calls[1]!.callbacks.onDone()

  onDone.mockClear()
  loop.run('Continue the old conversation')
  await advance()
  expect(calls).toHaveLength(3)
  expect(calls[2]!.request.tools).toEqual([])
  expect(calls[2]!.request.messages).toContainEqual({
    role: 'user',
    text: 'Old conversation instruction',
  })
  expect(loop.busy).toBe(true)
  return { loop, calls, onDone, summary: calls[2]! }
}

describe('AgentLoop reset during compaction', () => {
  it('keeps history empty when reset cancels a pending summary', async () => {
    const { loop, calls, onDone, summary } = await startCompaction()

    loop.reset()
    expect(summary.cancel).toHaveBeenCalledOnce()
    expect(loop.messages).toEqual([])
    await flush()

    // Cancellation without summary text must not reintroduce a mechanical digest.
    expect(loop.messages).toEqual([])
    expect(loop.busy).toBe(false)
    expect(calls).toHaveLength(3)
    expect(onDone).not.toHaveBeenCalled()
  })

  it('cancels a timed-out summary before starting the real turn', async () => {
    vi.useFakeTimers()
    try {
      const { calls, loop, summary } = await startCompaction(() => vi.advanceTimersByTimeAsync(0))
      await vi.advanceTimersByTimeAsync(30_000)
      expect(summary.cancel).toHaveBeenCalledOnce()
      expect(calls).toHaveLength(4)
      const continuation = calls[3]!
      expect(continuation.request.messages[0]).toEqual({
        role: 'user',
        text: expect.stringContaining('[Summary of earlier conversation'),
      })
      continuation.callbacks.onDelta('Continued answer')
      continuation.callbacks.onDone()
      expect(loop.busy).toBe(false)
    } finally {
      vi.useRealTimers()
    }
  })

  it('preserves a new conversation started immediately after reset', async () => {
    const { loop, calls, onDone, summary } = await startCompaction()
    summary.callbacks.onDelta('Summary belonging only to the old conversation')

    loop.reset()
    loop.run('Fresh conversation instruction')
    await flush()
    expect(calls).toHaveLength(4)
    const freshTurn = calls[3]!
    freshTurn.callbacks.onDelta('Fresh conversation answer')
    freshTurn.callbacks.onDone()

    expect(loop.messages).toEqual([
      { role: 'user', text: 'Fresh conversation instruction' },
      { role: 'assistant', text: 'Fresh conversation answer' },
    ])
    expect(freshTurn.request.messages).toEqual([
      { role: 'user', text: 'Fresh conversation instruction' },
    ])
    expect(loop.busy).toBe(false)
    expect(onDone).toHaveBeenCalledExactlyOnceWith({
      text: 'Fresh conversation answer',
      cancelled: false,
      turnLimit: false,
    })
  })

  it('retains the summary and recent messages when the conversation is not reset', async () => {
    const { loop, calls, summary } = await startCompaction()
    summary.callbacks.onDelta('Old conversation summary')
    summary.callbacks.onDone()
    await flush()

    expect(calls).toHaveLength(4)
    const continuation = calls[3]!
    expect(continuation.request.messages[0]).toEqual({
      role: 'user',
      text: expect.stringContaining('Old conversation summary'),
    })
    expect(continuation.request.messages.slice(2)).toEqual([
      { role: 'user', text: 'Recent question' },
      { role: 'assistant', text: 'Recent answer' },
      { role: 'user', text: 'Continue the old conversation' },
    ])
    continuation.callbacks.onDelta('Continued answer')
    continuation.callbacks.onDone()
    expect(loop.messages.at(-1)).toEqual({ role: 'assistant', text: 'Continued answer' })
    expect(loop.busy).toBe(false)
  })
})
