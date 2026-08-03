import { describe, expect, it, vi } from 'vitest'
import {
  AgentLoop,
  composeSkills,
  type AgentMessage,
  type AgentSkill,
  type AgentStreamCallbacks,
  type AgentToolCall,
  type AgentTransport,
  type ToolExecution,
} from '../src'

/** transport scripted turn by turn; exposes the callbacks for manual driving */
function scriptedTransport(script: Array<(cb: AgentStreamCallbacks) => void>): AgentTransport & {
  requests: Array<{ messageCount: number; toolCount: number }>
  cancels: number
} {
  let turn = 0
  const transport = {
    requests: [] as Array<{ messageCount: number; toolCount: number }>,
    cancels: 0,
    lastCallbacks: null as AgentStreamCallbacks | null,
    stream(request: { messages: AgentMessage[]; tools: unknown[] }, cb: AgentStreamCallbacks) {
      transport.requests.push({
        messageCount: request.messages.length,
        toolCount: request.tools.length,
      })
      transport.lastCallbacks = cb
      const step = script[turn++]
      if (step) queueMicrotask(() => step(cb))
      return {
        cancel: () => {
          transport.cancels++
          queueMicrotask(() => cb.onDone())
        },
      }
    },
  }
  return transport
}

function makeSkill(execute?: (call: AgentToolCall) => ToolExecution): AgentSkill {
  return {
    id: 'test',
    systemPrompt: 'system',
    tools: [{ name: 'do_thing', description: 'd', inputSchema: { type: 'object' } }],
    buildContext: () => 'CTX',
    executeTool: execute ?? (() => ({ output: 'ok', summary: 'done', mutated: true })),
  }
}

const flush = () => new Promise((r) => setTimeout(r, 0))

describe('AgentLoop', () => {
  it('runs a plain-text turn to completion', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('Hello')
        cb.onDelta(', world')
        cb.onDone()
      },
    ])
    const onDone = vi.fn()
    const onText = vi.fn()
    const loop = new AgentLoop({ transport, skill: makeSkill(), events: { onDone, onText } })
    loop.run('question')
    await flush()
    expect(onText).toHaveBeenLastCalledWith('Hello, world')
    expect(onDone).toHaveBeenCalledWith({
      text: 'Hello, world',
      cancelled: false,
      turnLimit: false,
    })
    expect(loop.busy).toBe(false)
    // user message carries the skill context
    expect(loop.messages[0]).toEqual({ role: 'user', text: 'question\n\nCTX' })
    expect(loop.messages[1]).toEqual({ role: 'assistant', text: 'Hello, world' })
  })

  it('attaches images to the user message (and omits the field without any)', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('I can see it')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({ transport, skill: makeSkill() })
    const images = [{ base64: 'AAAA', mime: 'image/png' }]
    loop.run('describe the image', images)
    await flush()
    expect(loop.messages[0]).toEqual({ role: 'user', text: 'describe the image\n\nCTX', images })
    expect('images' in (loop.messages[0] as object)).toBe(true)
  })

  it('executes tools and loops back to the model', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('Let me edit the document first')
        cb.onToolCall({ id: 't1', name: 'do_thing', input: { a: 1 } })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('All done')
        cb.onDone()
      },
    ])
    const executed: AgentToolCall[] = []
    const skill = makeSkill((call) => {
      executed.push(call)
      return { output: 'result-1', summary: 'changed 1 spot', mutated: true }
    })
    const onToolExecuted = vi.fn()
    const onDone = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill,
      captureSnapshot: () => 'SNAP',
      events: { onToolExecuted, onDone },
    })
    loop.run('make a change')
    await flush()
    await flush()

    expect(executed).toHaveLength(1)
    // first mutation carries the pre-tool snapshot
    expect(onToolExecuted).toHaveBeenCalledWith(expect.objectContaining({ snapshotBefore: 'SNAP' }))
    expect(onDone).toHaveBeenCalledWith({ text: 'All done', cancelled: false, turnLimit: false })
    // history: user / assistant+tools / tool results / assistant
    expect(loop.messages.map((m) => m.role)).toEqual(['user', 'assistant', 'tool', 'assistant'])
    const toolMsg = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMsg.results[0]).toEqual({
      id: 't1',
      name: 'do_thing',
      output: 'result-1',
      isError: undefined,
    })
    // second request included the tool round-trip
    expect(transport.requests[1].messageCount).toBe(3)
  })

  it('emits onToolStart before each execution, paired with onToolExecuted', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: { a: 1 } })
        cb.onToolCall({ id: 't2', name: 'do_thing', input: { a: 2 } })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const order: string[] = []
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => {
        order.push('exec')
        return { output: 'ok', summary: 's' }
      }),
      events: {
        onToolStart: (call) => order.push(`start:${call.id}`),
        onToolExecuted: ({ call }) => order.push(`done:${call.id}`),
      },
    })
    loop.run('x')
    await flush()
    await flush()
    expect(order).toEqual(['start:t1', 'exec', 'done:t1', 'start:t2', 'exec', 'done:t2'])
  })

  it('only the first mutating tool carries a snapshot', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onToolCall({ id: 't2', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      captureSnapshot: () => 'SNAP',
      events: { onToolExecuted },
    })
    loop.run('x')
    await flush()
    await flush()
    expect(onToolExecuted).toHaveBeenCalledTimes(2)
    expect(onToolExecuted.mock.calls[0][0].snapshotBefore).toBe('SNAP')
    expect(onToolExecuted.mock.calls[1][0].snapshotBefore).toBeUndefined()
  })

  it('after maxTurns, adds a final tool-less turn that yields a partial answer', async () => {
    const alwaysTool = (cb: AgentStreamCallbacks) => {
      cb.onToolCall({ id: 'x', name: 'do_thing', input: {} })
      cb.onDone()
    }
    const finalize = (cb: AgentStreamCallbacks) => {
      cb.onDelta('partial conclusion')
      cb.onDone()
    }
    const transport = scriptedTransport([alwaysTool, alwaysTool, finalize])
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill: makeSkill(), maxTurns: 2, events: { onDone } })
    loop.run('x')
    await flush()
    await flush()
    await flush()
    await flush()
    // The third turn is the finalizing one: no tools, and a system note was inserted into history
    expect(transport.requests).toHaveLength(3)
    expect(transport.requests[2].toolCount).toBe(0)
    const note = loop.messages.find((m) => m.role === 'user' && m.text.includes('turn limit'))
    expect(note).toBeDefined()
    expect(onDone).toHaveBeenCalledWith({
      text: 'partial conclusion',
      cancelled: false,
      turnLimit: true,
    })
  })

  it('cancel drops pending tool calls and finalizes the run', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('partial')
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        // no onDone: waits for cancel
      },
    ])
    const executed: AgentToolCall[] = []
    const skill = makeSkill((call) => {
      executed.push(call)
      return { output: 'ok', summary: 's' }
    })
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })
    loop.run('x')
    await flush()
    loop.cancel()
    await flush()
    expect(transport.cancels).toBe(1)
    expect(executed).toHaveLength(0)
    expect(onDone).toHaveBeenCalledWith({ text: 'partial', cancelled: true, turnLimit: false })
    // assistant message stored without toolCalls (no results would follow)
    expect(loop.messages[1]).toEqual({ role: 'assistant', text: 'partial' })
  })

  it('cancel during tool execution aborts the signal, skips remaining tools and starts no new turn', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onToolCall({ id: 't2', name: 'do_thing', input: {} })
        cb.onDone()
      },
      // The second turn should never be requested (no return to the model after cancel)
      (cb) => cb.onDone(),
    ])
    const seen: Array<{ id: string; abortedAfterCancel: boolean }> = []
    let loop: AgentLoop | null = null
    const skill = makeSkill()
    skill.executeTool = (call, signal) => {
      // Simulate the user hitting stop while a long tool is running
      loop?.cancel()
      seen.push({ id: call.id, abortedAfterCancel: signal?.aborted === true })
      return { output: 'ok', summary: 's', mutated: true }
    }
    const onDone = vi.fn()
    loop = new AgentLoop({ transport, skill, events: { onDone } })
    loop.run('x')
    await flush()
    await flush()
    // Only the first tool ran; after cancel the signal is aborted right away (long tools' inner loops break on it)
    expect(seen).toEqual([{ id: 't1', abortedAfterCancel: true }])
    // The second tool didn't run, but a paired error result was added (tool_use/tool_result stay paired)
    const toolMsg = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMsg.results.map((r) => r.id)).toEqual(['t1', 't2'])
    expect(toolMsg.results[1].isError).toBe(true)
    // No further model request; the run finishes as cancelled
    expect(transport.requests).toHaveLength(1)
    expect(onDone).toHaveBeenCalledTimes(1)
    expect(onDone).toHaveBeenCalledWith({ text: '', cancelled: true, turnLimit: false })
    expect(loop.busy).toBe(false)
  })

  it('executeTool receives a live (non-aborted) signal during normal runs', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    let sawSignal: AbortSignal | undefined
    const skill = makeSkill()
    skill.executeTool = (_call, signal) => {
      sawSignal = signal
      return { output: 'ok', summary: 's' }
    }
    const loop = new AgentLoop({ transport, skill })
    loop.run('x')
    await flush()
    await flush()
    expect(sawSignal).toBeInstanceOf(AbortSignal)
    expect(sawSignal?.aborted).toBe(false)
  })

  it('trims history at user boundaries', async () => {
    const script = Array.from({ length: 4 }, () => (cb: AgentStreamCallbacks) => {
      cb.onDelta('answer')
      cb.onDone()
    })
    const transport = scriptedTransport(script)
    const loop = new AgentLoop({ transport, skill: makeSkill(), maxHistory: 3 })
    for (let i = 0; i < 4; i++) {
      loop.run(`question${i}`)
      await flush()
    }
    // trimmed to start at a user message
    expect(loop.messages.length).toBeLessThanOrEqual(4)
    expect(loop.messages[0].role).toBe('user')
  })

  it('restore seeds history and the next run sends it to the model', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('picking up from before')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({ transport, skill: makeSkill() })
    loop.restore([
      { role: 'user', text: 'previous question' },
      { role: 'assistant', text: 'previous answer' },
    ])
    expect(loop.messages).toHaveLength(2)

    loop.run('follow-up')
    await flush()
    // The request carries the 2 restored messages + the new user message
    expect(transport.requests[0].messageCount).toBe(3)
    expect(loop.messages[0]).toEqual({ role: 'user', text: 'previous question' })
  })

  it('restore is a no-op when history exists or the loop is busy', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('answer')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({ transport, skill: makeSkill() })
    loop.run('opening message')
    await flush()
    const before = loop.messages.length
    loop.restore([{ role: 'user', text: 'should-not-be-injected' }])
    expect(loop.messages.length).toBe(before)
    expect(
      loop.messages.some(
        (m) => m.role === 'user' && 'text' in m && m.text === 'should-not-be-injected',
      ),
    ).toBe(false)
  })

  it('a failed run rolls its user message back out of history', async () => {
    const transport = scriptedTransport([
      (cb) => cb.onError('Not signed in'),
      (cb) => {
        cb.onDelta('answer to the second question')
        cb.onDone()
      },
    ])
    const onError = vi.fn()
    const loop = new AgentLoop({ transport, skill: makeSkill(), events: { onError } })
    loop.run('change all headings to red')
    await flush()
    expect(onError).toHaveBeenCalledWith('Not signed in')
    // the failed instruction is gone, so it can't be re-executed by the next run
    expect(loop.messages).toHaveLength(0)

    loop.run('what does this report propose?')
    await flush()
    // the next request carries only the new user message — no adjacent user turns
    expect(transport.requests[1].messageCount).toBe(1)
    expect(loop.messages.map((m) => m.role)).toEqual(['user', 'assistant'])
  })

  it('a mid-run failure after tool turns rolls back the whole run', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onError('network dropped'),
    ])
    const onError = vi.fn()
    const loop = new AgentLoop({ transport, skill: makeSkill(), events: { onError } })
    loop.restore([
      { role: 'user', text: 'earlier question' },
      { role: 'assistant', text: 'earlier answer' },
    ])
    loop.run('do more work')
    await flush()
    await flush()
    expect(onError).toHaveBeenCalledWith('network dropped')
    // history is back to the pre-run state: no dangling user/assistant/tool from the failed run
    expect(loop.messages).toEqual([
      { role: 'user', text: 'earlier question' },
      { role: 'assistant', text: 'earlier answer' },
    ])
  })

  it('restore drops unanswered user messages (trailing and adjacent)', () => {
    const transport = scriptedTransport([])
    const loop = new AgentLoop({ transport, skill: makeSkill() })
    loop.restore([
      { role: 'user', text: 'failed and never answered' },
      { role: 'user', text: 'answered question' },
      { role: 'assistant', text: 'the answer' },
      { role: 'user', text: 'trailing unanswered' },
    ])
    expect(loop.messages).toEqual([
      { role: 'user', text: 'answered question' },
      { role: 'assistant', text: 'the answer' },
    ])
  })

  it('restore trims oversized history at a user boundary', () => {
    const transport = scriptedTransport([])
    const loop = new AgentLoop({ transport, skill: makeSkill(), maxHistory: 2 })
    loop.restore([
      { role: 'user', text: 'q1' },
      { role: 'assistant', text: 'a1' },
      { role: 'user', text: 'q2' },
      { role: 'assistant', text: 'a2' },
    ])
    expect(loop.messages.length).toBeLessThanOrEqual(2)
    expect(loop.messages[0]).toEqual({ role: 'user', text: 'q2' })
  })
})

describe('AgentLoop compaction', () => {
  it('over budget before run: folds old conversation via LLM summary; new request carries the summary, not the originals', async () => {
    const bigAnswer = 'old reply'.padEnd(800, 'y')
    const transport = scriptedTransport([
      // Two real turns first: turn 1 produces a long over-budget reply
      (cb) => {
        cb.onDelta(bigAnswer)
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('second-turn reply')
        cb.onDone()
      },
      // The 3rd run triggers compaction: this stream call is the summary request
      (cb) => {
        cb.onDelta('Goal: build a deck; 5 slides done')
        cb.onDone()
      },
      // Only then comes the actual model turn for run 3
      (cb) => {
        cb.onDelta('continuing')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      compaction: { maxBytes: 500, keepRecentBytes: 100 },
    })
    loop.run('old instruction')
    await flush()
    loop.run('second question')
    await flush()
    loop.run('follow-up')
    await flush()
    expect(transport.requests).toHaveLength(4)
    const msgs = loop.messages
    expect(msgs[0]!.role).toBe('user')
    expect((msgs[0] as { text: string }).text).toContain('[Summary of earlier conversation')
    expect((msgs[0] as { text: string }).text).toContain('5 slides done')
    // The folded original text is gone from history
    expect(msgs.some((m) => 'text' in m && m.text === bigAnswer)).toBe(false)
    // The new user message comes after the summary
    expect(msgs.some((m) => m.role === 'user' && 'text' in m && m.text.includes('follow-up'))).toBe(
      true,
    )
  })

  it('falls back to a mechanical digest when the LLM summary fails; the turn proceeds normally', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('done'.padEnd(800, 'y'))
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('second turn')
        cb.onDone()
      },
      (cb) => cb.onError('summary failed'),
      (cb) => {
        cb.onDelta('answer')
        cb.onDone()
      },
    ])
    const onDone = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      compaction: { maxBytes: 500, keepRecentBytes: 100 },
      events: { onDone },
    })
    loop.run('key instruction')
    await flush()
    loop.run('second question')
    await flush()
    loop.run('continue')
    await flush()
    expect(onDone).toHaveBeenLastCalledWith({ text: 'answer', cancelled: false, turnLimit: false })
    // The mechanical digest kept the gist of the user instruction
    expect((loop.messages[0] as { text: string }).text).toContain('key instruction')
    expect((loop.messages[0] as { text: string }).text).toContain(
      '[Summary of earlier conversation',
    )
  })

  it('restore over budget folds mechanically (no LLM request)', () => {
    const transport = scriptedTransport([])
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      compaction: { maxBytes: 500, keepRecentBytes: 200 },
    })
    loop.restore([
      { role: 'user', text: 'very old instruction'.padEnd(600, 'x') },
      { role: 'assistant', text: 'early reply' },
      { role: 'user', text: 'recent question' },
      { role: 'assistant', text: 'recent answer' },
    ])
    expect(transport.requests).toHaveLength(0)
    expect((loop.messages[0] as { text: string }).text).toContain(
      '[Summary of earlier conversation',
    )
    expect((loop.messages[0] as { text: string }).text).toContain('very old instruction')
    expect(loop.messages.some((m) => 'text' in m && m.text === 'recent question')).toBe(true)
  })

  it('over budget mid-run truncates stale tool outputs (the 2 most recent keep the full text)', async () => {
    const big = 'z'.repeat(3_000)
    const script = Array.from({ length: 3 }, (_, i) => (cb: AgentStreamCallbacks) => {
      cb.onToolCall({ id: `t${i}`, name: 'do_thing', input: {} })
      cb.onDone()
    })
    script.push((cb) => {
      cb.onDelta('done')
      cb.onDone()
    })
    const transport = scriptedTransport(script)
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => ({ output: big, summary: 'big output', mutated: false })),
      compaction: { maxBytes: 5_000, keepRecentBytes: 2_000, disableLlmSummary: true },
    })
    loop.run('do the work')
    await flush()
    await flush()
    await flush()
    await flush()
    const toolMsgs = loop.messages.filter((m) => m.role === 'tool')
    expect(toolMsgs).toHaveLength(3)
    const first = toolMsgs[0] as { results: Array<{ output: string }> }
    const last = toolMsgs[2] as { results: Array<{ output: string }> }
    expect(first.results[0]!.output).toContain('…(output truncated: too long)')
    expect(first.results[0]!.output.length).toBeLessThan(1_200)
    expect(last.results[0]!.output).toBe(big)
  })

  it('compaction: false disables both folding and truncation', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('answer')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({ transport, skill: makeSkill(), compaction: false })
    const bigText = 'x'.repeat(200_000)
    loop.restore([
      { role: 'user', text: bigText },
      { role: 'assistant', text: 'ok' },
    ])
    loop.run('continue')
    await flush()
    // Only the real model turn, no summary request; the original text was not folded
    expect(transport.requests).toHaveLength(1)
    expect((loop.messages[0] as { text: string }).text).toBe(bigText)
  })

  it('tool executor exceptions become error results and the loop continues', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('OK')
        cb.onDone()
      },
    ])
    const skill = makeSkill(() => {
      throw new Error('boom')
    })
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })
    loop.run('x')
    await flush()
    await flush()
    const toolMsg = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMsg.results[0].isError).toBe(true)
    expect(toolMsg.results[0].output).toBe('boom')
    expect(onDone).toHaveBeenCalledWith({ text: 'OK', cancelled: false, turnLimit: false })
  })

  it('calls with inputError are not executed; an is_error result is fed back so the model can retry', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {}, inputError: 'bad json' })
        cb.onDone()
      },
      (cb) => {
        cb.onToolCall({ id: 't2', name: 'do_thing', input: { a: 1 } })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('done')
        cb.onDone()
      },
    ])
    const executed: AgentToolCall[] = []
    const skill = makeSkill((call) => {
      executed.push(call)
      return { output: 'ok', summary: 'ok' }
    })
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })
    loop.run('x')
    await flush()
    await flush()
    await flush()
    expect(executed.map((c) => c.id)).toEqual(['t2'])
    const toolMsg = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMsg.results[0].isError).toBe(true)
    expect(toolMsg.results[0].output).toContain('bad json')
    expect(onDone).toHaveBeenCalledWith({ text: 'done', cancelled: false, turnLimit: false })
  })

  it('terminates the run after consecutive input-parse failures hit the limit', async () => {
    const badTurn = (cb: AgentStreamCallbacks) => {
      cb.onToolCall({ id: 't', name: 'do_thing', input: {}, inputError: 'bad json' })
      cb.onDone()
    }
    const transport = scriptedTransport([badTurn, badTurn, badTurn, badTurn])
    const executed: AgentToolCall[] = []
    const skill = makeSkill((call) => {
      executed.push(call)
      return { output: 'ok', summary: 'ok' }
    })
    const onError = vi.fn()
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onError, onDone } })
    loop.run('x')
    for (let i = 0; i < 6; i++) await flush()
    expect(executed).toHaveLength(0)
    expect(transport.requests).toHaveLength(3)
    expect(onError).toHaveBeenCalledWith(expect.stringContaining('retries stopped'))
    expect(onDone).not.toHaveBeenCalled()
    expect(loop.busy).toBe(false)
  })
})

describe('composeSkills', () => {
  it('merges prompts, tools and context, and routes execution', async () => {
    const a: AgentSkill = {
      id: 'a',
      systemPrompt: 'PA',
      tools: [{ name: 'tool_a', description: '', inputSchema: {} }],
      buildContext: () => 'CA',
      executeTool: () => ({ output: 'from-a', summary: 'a' }),
    }
    const b: AgentSkill = {
      id: 'b',
      systemPrompt: 'PB',
      tools: [{ name: 'tool_b', description: '', inputSchema: {} }],
      executeTool: () => ({ output: 'from-b', summary: 'b' }),
    }
    const merged = composeSkills('ab', 'INTRO', [a, b])
    expect(merged.systemPrompt).toBe('INTRO\n\nPA\n\nPB')
    expect(merged.tools.map((t) => t.name)).toEqual(['tool_a', 'tool_b'])
    expect(merged.buildContext?.()).toBe('CA')
    expect((await merged.executeTool({ id: '1', name: 'tool_b', input: {} })).output).toBe('from-b')
    const unknown = await merged.executeTool({ id: '2', name: 'nope', input: {} })
    expect(unknown.isError).toBe(true)
  })

  it('rejects duplicate tool names', () => {
    const tool = { name: 'same', description: '', inputSchema: {} }
    const make = (id: string): AgentSkill => ({
      id,
      systemPrompt: '',
      tools: [tool],
      executeTool: () => ({ output: '', summary: '' }),
    })
    expect(() => composeSkills('x', '', [make('a'), make('b')])).toThrow(/duplicate/)
  })
})
