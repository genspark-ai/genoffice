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
      const cancel = vi.fn(() => queueMicrotask(() => callbacks.onDone()))
      calls.push({ request, callbacks, cancel })
      return { cancel }
    },
  }
  return { transport, calls }
}

describe('AgentLoop compaction summary billing', () => {
  it('folds the summary turn usage into the run total and reports it', async () => {
    const { transport, calls } = controlledTransport()
    const skill: AgentSkill = {
      id: 'test',
      systemPrompt: 'Test system prompt',
      tools: [{ name: 'read', description: 'Read', inputSchema: { type: 'object' } }],
      executeTool: () => ({ output: 'ok', summary: 'Read completed' }),
    }
    const onDone = vi.fn()
    const onUsage = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill,
      events: { onDone, onUsage },
      compaction: { maxBytes: 500, keepRecentBytes: 100 },
    })

    loop.run('Old conversation instruction')
    await flush()
    calls[0]!.callbacks.onDelta('Old answer '.repeat(80))
    calls[0]!.callbacks.onDone()
    loop.run('Recent question')
    await flush()
    calls[1]!.callbacks.onDelta('Recent answer')
    calls[1]!.callbacks.onDone()

    onUsage.mockClear()
    loop.run('Continue the old conversation')
    await flush()
    // the over-budget history triggers a summary request before the real turn
    expect(calls).toHaveLength(3)
    const summary = calls[2]!
    expect(summary.request.tools).toEqual([])
    summary.callbacks.onUsage?.({ inputTokens: 900, outputTokens: 50 })
    summary.callbacks.onDelta('- Old conversation summarized')
    summary.callbacks.onDone()
    await flush()

    // the summary turn's billing is surfaced like any other turn's
    expect(onUsage).toHaveBeenCalledWith({ inputTokens: 900, outputTokens: 50 })

    expect(calls).toHaveLength(4)
    const continuation = calls[3]!
    continuation.callbacks.onUsage?.({ inputTokens: 100, outputTokens: 10 })
    continuation.callbacks.onDelta('Continued answer')
    continuation.callbacks.onDone()
    await flush()

    // run total = summary turn + real turn
    expect(onDone).toHaveBeenCalledWith({
      text: 'Continued answer',
      cancelled: false,
      turnLimit: false,
      usage: { inputTokens: 1000, outputTokens: 60 },
    })
  })
})
