import { describe, expect, it, vi } from 'vitest'
import {
  AgentLoop,
  TOOL_ABORTED_OUTPUT,
  TOOL_DENIED_OUTPUT,
  type AgentMessage,
  type AgentSkill,
  type AgentStreamCallbacks,
  type AgentToolCall,
  type AgentToolDef,
  type AgentTransport,
  type ToolExecution,
} from '../src'

/** transport scripted turn by turn; exposes the callbacks for manual driving */
function scriptedTransport(script: Array<(cb: AgentStreamCallbacks) => void>): AgentTransport {
  let turn = 0
  return {
    stream(_request, cb) {
      const step = script[turn++]
      if (step) queueMicrotask(() => step(cb))
      return {
        cancel: () => {
          queueMicrotask(() => cb.onDone())
        },
      }
    },
  }
}

function makeSkill(
  execute?: (call: AgentToolCall) => ToolExecution,
  risk?: AgentToolDef['risk'],
): AgentSkill {
  return {
    id: 'test',
    systemPrompt: 'system',
    tools: [
      {
        name: 'do_thing',
        description: 'd',
        inputSchema: { type: 'object' },
        ...(risk ? { risk } : {}),
      },
    ],
    buildContext: () => '',
    executeTool: execute ?? (() => ({ output: 'tool ran', summary: 'done', mutated: true })),
  }
}

const flush = () => new Promise((r) => setTimeout(r, 0))

function toolResultOf(messages: readonly AgentMessage[]): { output: string; isError?: boolean } {
  const toolMsg = messages.find((m) => m.role === 'tool')
  if (!toolMsg || toolMsg.role !== 'tool') throw new Error('no tool message in history')
  return toolMsg.results[0]!
}

describe('AgentLoop tool confirmation (confirmTool on risk:high tools)', () => {
  it('executes the high-risk tool when the host allows it', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: { a: 1 } })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('all done')
        cb.onDone()
      },
    ])
    const executed: AgentToolCall[] = []
    const confirmTool = vi.fn(async (_call: AgentToolCall, _tool: AgentToolDef) => true)
    const loop = new AgentLoop({
      transport,
      skill: makeSkill((call) => {
        executed.push(call)
        return { output: 'tool ran', summary: 'done', mutated: true }
      }, 'high'),
      confirmTool,
    })
    const handle = loop.run('make a change')
    await flush()
    await flush()

    expect(confirmTool).toHaveBeenCalledTimes(1)
    // the hook sees the call and the tool definition it gated
    expect(confirmTool.mock.calls[0]![0]).toMatchObject({ id: 't1', name: 'do_thing' })
    expect(confirmTool.mock.calls[0]![1].risk).toBe('high')
    expect(executed).toHaveLength(1)
    expect(toolResultOf(loop.messages).output).toBe('tool ran')
    await expect(handle.done).resolves.toEqual({
      text: 'all done',
      cancelled: false,
      turnLimit: false,
    })
  })

  it('does not execute when the host denies: the model gets the denial as an error result and continues', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('okay, I will not do that then')
        cb.onDone()
      },
    ])
    const executed: AgentToolCall[] = []
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill((call) => {
        executed.push(call)
        return { output: 'tool ran', summary: 'done', mutated: true }
      }, 'high'),
      confirmTool: async () => false,
      events: { onToolExecuted },
    })
    const handle = loop.run('make a change')
    await flush()
    await flush()

    // the tool never ran
    expect(executed).toHaveLength(0)
    // the model sees a paired error result it can act on
    const result = toolResultOf(loop.messages)
    expect(result.output).toBe(TOOL_DENIED_OUTPUT)
    expect(result.isError).toBe(true)
    // UIs can render the denial distinctly from a failure
    expect(onToolExecuted).toHaveBeenCalledWith(
      expect.objectContaining({ execution: expect.objectContaining({ denied: true }) }),
    )
    // the run carried on and answered
    await expect(handle.done).resolves.toEqual({
      text: 'okay, I will not do that then',
      cancelled: false,
      turnLimit: false,
    })
  })

  it('skips confirmation entirely when the host provides no hook (behavior unchanged)', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const executed: AgentToolCall[] = []
    const loop = new AgentLoop({
      transport,
      skill: makeSkill((call) => {
        executed.push(call)
        return { output: 'tool ran', summary: 'done', mutated: true }
      }, 'high'),
    })
    loop.run('make a change')
    await flush()
    await flush()
    expect(executed).toHaveLength(1)
    expect(toolResultOf(loop.messages).output).toBe('tool ran')
  })

  it('never calls the hook for a low-risk tool', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const confirmTool = vi.fn(async () => true)
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(undefined, 'low'),
      confirmTool,
    })
    loop.run('x')
    await flush()
    await flush()
    expect(confirmTool).not.toHaveBeenCalled()
  })

  it('treats a throwing hook as a denial (a broken host must not run a dangerous tool)', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const executed: AgentToolCall[] = []
    const loop = new AgentLoop({
      transport,
      skill: makeSkill((call) => {
        executed.push(call)
        return { output: 'tool ran', summary: 'done', mutated: true }
      }, 'high'),
      confirmTool: () => Promise.reject(new Error('host hook exploded')),
    })
    const onDone = vi.fn()
    const events = { onDone }
    const handle = loop.run('x')
    void events
    await flush()
    await flush()
    expect(executed).toHaveLength(0)
    expect(toolResultOf(loop.messages).output).toBe(TOOL_DENIED_OUTPUT)
    // the run itself survived the throw and finished
    await expect(handle.done).resolves.toMatchObject({ cancelled: false })
  })

  it('a cancelled run ends the pending confirmation instead of hanging on the dialog', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const executed: AgentToolCall[] = []
    const loop = new AgentLoop({
      transport,
      skill: makeSkill((call) => {
        executed.push(call)
        return { output: 'tool ran', summary: 'done', mutated: true }
      }, 'high'),
      // never settles: the dialog is left open
      confirmTool: () => new Promise<boolean>(() => {}),
    })
    const handle = loop.run('x')
    await flush() // turn 1 delivered, finishTurn reached the gate
    loop.cancel()
    await flush()
    await flush()

    expect(executed).toHaveLength(0)
    expect(toolResultOf(loop.messages).output).toBe(TOOL_ABORTED_OUTPUT)
    expect(toolResultOf(loop.messages).isError).toBe(true)
    await expect(handle.done).resolves.toMatchObject({ cancelled: true })
    expect(loop.busy).toBe(false)
  })

  it('emits awaiting-confirm on the run handle with the call and the tool definition', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const seen: Array<{ call: AgentToolCall; tool: AgentToolDef }> = []
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(undefined, 'high'),
      confirmTool: async () => true,
    })
    const handle = loop.run('x')
    handle.on('awaiting-confirm', (payload) => seen.push(payload))
    await flush()
    await flush()
    expect(seen).toHaveLength(1)
    expect(seen[0]!.call.name).toBe('do_thing')
    expect(seen[0]!.tool.risk).toBe('high')
  })

  it('a denied call counts as not-ok for verifyResponse (the claimed-action guard)', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('I saved it for you')
        cb.onDone()
      },
    ])
    const corrections: string[] = []
    const skill: AgentSkill = {
      ...makeSkill(undefined, 'high'),
      verifyResponse: (finalText, executedCalls) => {
        const correction = executedCalls.some((c) => c.ok)
          ? null
          : 'the tool did not run; admit that nothing happened'
        if (correction) corrections.push(correction)
        return null
      },
    }
    const loop = new AgentLoop({ transport, skill, confirmTool: async () => false })
    loop.run('x')
    await flush()
    await flush()
    expect(corrections).toHaveLength(1)
    expect(toolResultOf(loop.messages).output).toBe(TOOL_DENIED_OUTPUT)
  })
})
