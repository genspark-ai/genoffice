import { describe, expect, it, vi } from 'vitest'
import {
  AgentLoop,
  type AgentMessage,
  type AgentSkill,
  type AgentStreamCallbacks,
  type AgentToolCall,
  type AgentTransport,
  type ToolExecution,
} from '../src'

/**
 * Transport whose turns complete only when the test drives them, so a run can
 * be held in flight (queue/pause tests need the loop to stay busy on demand).
 */
function heldTransport() {
  const requests: Array<{ messages: AgentMessage[]; tools: unknown[] }> = []
  const pending: AgentStreamCallbacks[] = []
  let cancels = 0
  const transport: AgentTransport = {
    stream(request, cb) {
      requests.push({ messages: [...request.messages], tools: [...request.tools] })
      pending.push(cb)
      return {
        cancel: () => {
          cancels++
          queueMicrotask(() => cb.onDone())
        },
      }
    },
  }
  return {
    transport,
    requests,
    cancels: () => cancels,
    /** deliver the model's answer for the oldest in-flight turn */
    turn: (deliver: (cb: AgentStreamCallbacks) => void) => {
      const cb = pending.shift()
      if (cb) deliver(cb)
    },
  }
}

function makeSkill(
  execute?: (call: AgentToolCall) => ToolExecution | Promise<ToolExecution>,
): AgentSkill {
  return {
    id: 'test',
    systemPrompt: 'system',
    tools: [{ name: 'do_thing', description: 'd', inputSchema: { type: 'object' } }],
    buildContext: () => '',
    executeTool: execute ?? (() => ({ output: 'ok', summary: 'done', mutated: true })),
  }
}

const flush = () => new Promise((r) => setTimeout(r, 0))

const lastText = (messages: readonly AgentMessage[]): string => {
  const last = messages.at(-1)
  return last && 'text' in last ? last.text : ''
}

describe('AgentLoop run handles', () => {
  it('run() returns a handle whose done resolves with the run result', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const handle = loop.run('question')
    expect(handle.id).toBe('run-1')
    await flush() // beginRun awaits compaction before the first stream call
    t.turn((cb) => {
      cb.onDelta('answer')
      cb.onDone()
    })
    await flush()
    await expect(handle.done).resolves.toEqual({
      text: 'answer',
      cancelled: false,
      turnLimit: false,
    })
    expect(loop.busy).toBe(false)
  })

  it('an empty instruction never reaches the model; its handle settles cancelled', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const handle = loop.run('')
    await expect(handle.done).resolves.toEqual({
      text: '',
      cancelled: true,
      turnLimit: false,
    })
    expect(t.requests).toHaveLength(0)
    expect(loop.busy).toBe(false)
  })

  it('queues an instruction that arrives while a run is in flight, in FIFO order', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const first = loop.run('first question')
    const second = loop.run('second question')
    // nothing dropped: the second instruction waits behind the first
    expect(loop.busy).toBe(true)
    expect(loop.queuedCount).toBe(1)
    await flush() // the first run's stream is in flight now

    t.turn((cb) => {
      cb.onDelta('answer one')
      cb.onDone()
    })
    await flush()
    await expect(first.done).resolves.toEqual({
      text: 'answer one',
      cancelled: false,
      turnLimit: false,
    })
    // the queued run started on its own once the loop freed
    expect(t.requests).toHaveLength(2)
    expect(lastText(t.requests[1]!.messages)).toBe('second question')

    t.turn((cb) => {
      cb.onDelta('answer two')
      cb.onDone()
    })
    await flush()
    await expect(second.done).resolves.toEqual({
      text: 'answer two',
      cancelled: false,
      turnLimit: false,
    })
    expect(loop.busy).toBe(false)
    expect(loop.messages.map((m) => m.role)).toEqual(['user', 'assistant', 'user', 'assistant'])
  })

  it('pause holds the run between turns; a streaming turn finishes first', async () => {
    const t = heldTransport()
    // one controllable tool execution: the test releases it to move the run on
    let releaseTool: ((execution: ToolExecution) => void) | null = null
    const loop = new AgentLoop({
      transport: t.transport,
      skill: makeSkill(
        (call) =>
          new Promise<ToolExecution>((resolve) => {
            releaseTool = (execution) => resolve(execution)
          }),
      ),
    })
    const handle = loop.run('long task')
    await flush() // the first turn is streaming
    // turn 1: a tool call; the tool promise is pending
    t.turn((cb) => {
      cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
      cb.onDone()
    })
    await flush()
    expect(handle.id).toBe('run-1')
    // pause while the tool is still running
    handle.pause()
    releaseTool!({ output: 'ok', summary: 'done', mutated: true })
    await flush()
    await flush()
    // the tool turn ended, but the next model request is held
    expect(t.requests).toHaveLength(1)
    expect(loop.busy).toBe(true)

    handle.resume()
    await flush()
    expect(t.requests).toHaveLength(2)
    t.turn((cb) => {
      cb.onDelta('finished')
      cb.onDone()
    })
    await flush()
    await expect(handle.done).resolves.toEqual({
      text: 'finished',
      cancelled: false,
      turnLimit: false,
    })
  })

  it('a paused queued run stays at the head of the queue until resumed', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const first = loop.run('first')
    const second = loop.run('second')
    second.pause()
    await flush() // the first run's stream is in flight now

    t.turn((cb) => {
      cb.onDelta('one')
      cb.onDone()
    })
    await flush()
    await expect(first.done).resolves.toMatchObject({ text: 'one' })
    // the loop freed, but the paused head holds the queue (strict FIFO)
    expect(loop.busy).toBe(false)
    expect(loop.queuedCount).toBe(1)
    expect(t.requests).toHaveLength(1)

    second.resume()
    await flush()
    expect(t.requests).toHaveLength(2)
    t.turn((cb) => {
      cb.onDelta('two')
      cb.onDone()
    })
    await flush()
    await expect(second.done).resolves.toMatchObject({ text: 'two' })
  })

  it('cancel on the active run settles done as cancelled and frees the loop', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const handle = loop.run('question')
    await flush() // the turn is streaming now
    t.turn((cb) => cb.onDelta('partial answer'))
    // no onDone yet: the turn is still streaming when the user hits stop
    handle.cancel()
    await flush()
    await flush()
    await expect(handle.done).resolves.toEqual({
      text: 'partial answer',
      cancelled: true,
      turnLimit: false,
    })
    expect(t.cancels()).toBe(1)
    expect(loop.busy).toBe(false)
  })

  it('cancel on a queued run drops it: it never starts and the queue moves on', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const first = loop.run('first')
    const second = loop.run('second')
    second.cancel()
    await flush() // the first run's stream is in flight now
    await expect(second.done).resolves.toEqual({
      text: '',
      cancelled: true,
      turnLimit: false,
    })
    expect(loop.queuedCount).toBe(0)

    t.turn((cb) => {
      cb.onDelta('one')
      cb.onDone()
    })
    await flush()
    await expect(first.done).resolves.toMatchObject({ text: 'one' })
    // the cancelled run never reached the model
    expect(t.requests).toHaveLength(1)
    expect(loop.busy).toBe(false)
  })

  it('a failed run rejects done and emits error on the handle', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const handle = loop.run('question')
    const errors: string[] = []
    handle.on('error', (message) => errors.push(message))
    await flush() // the turn is streaming now
    t.turn((cb) => cb.onError('provider is down'))
    await flush()
    await expect(handle.done).rejects.toThrow('provider is down')
    expect(errors).toEqual(['provider is down'])
    expect(loop.busy).toBe(false)
  })

  it('reset settles the active run and every queued run as cancelled', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const first = loop.run('first')
    const second = loop.run('second')
    await flush() // the first turn is streaming now
    t.turn((cb) => cb.onDelta('partial'))
    loop.reset()
    await expect(first.done).resolves.toMatchObject({ cancelled: true })
    await expect(second.done).resolves.toMatchObject({ cancelled: true, text: '' })
    expect(loop.busy).toBe(false)
    expect(loop.queuedCount).toBe(0)
    expect(loop.messages).toHaveLength(0)
  })

  it('the handle streams text, tool and finished events; on() unsubscribes', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const handle = loop.run('question')
    const seen: string[] = []
    const offText = handle.on('text', (text) => seen.push(`text:${text}`))
    handle.on('tool', ({ call, execution }) =>
      seen.push(`tool:${call.name}:${execution ? execution.output : 'start'}`),
    )
    handle.on('finished', (result) => seen.push(`finished:${result.text}`))
    await flush() // the turn is streaming now

    t.turn((cb) => {
      cb.onDelta('Let me check')
      cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
      cb.onDone()
    })
    await flush()
    offText()
    t.turn((cb) => {
      cb.onDelta('more text')
      cb.onDone()
    })
    await flush()

    expect(seen).toEqual([
      'text:Let me check',
      'tool:do_thing:start',
      'tool:do_thing:ok',
      'finished:more text',
    ])
    // the unsubscribed listener missed the second turn's text
    expect(seen.filter((s) => s.startsWith('text:'))).toEqual(['text:Let me check'])
  })

  it('a queued run emits nothing until it starts', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const first = loop.run('first')
    const second = loop.run('second')
    const seen: string[] = []
    second.on('text', (text) => seen.push(text))
    second.on('finished', (result) => seen.push(result.text))
    await flush() // the first run's stream is in flight now
    t.turn((cb) => {
      cb.onDelta('one')
      cb.onDone()
    })
    await flush()
    expect(seen).toEqual([]) // still queued: no events belong to it yet
    t.turn((cb) => {
      cb.onDelta('two')
      cb.onDone()
    })
    await flush()
    // one text event and the finished event, both after the run started
    expect(seen).toEqual(['two', 'two'])
    await expect(first.done).resolves.toMatchObject({ text: 'one' })
    await expect(second.done).resolves.toMatchObject({ text: 'two' })
  })

  it('done never surfaces as an unhandled rejection when nobody awaits it', async () => {
    const t = heldTransport()
    const loop = new AgentLoop({ transport: t.transport, skill: makeSkill() })
    const handle = loop.run('question')
    const onError = vi.fn()
    // no catch attached to handle.done anywhere: the internal sink must own it
    await flush() // the turn is streaming now
    t.turn((cb) => cb.onError('boom'))
    await flush()
    await flush()
    expect(onError).not.toHaveBeenCalled()
    // still settled exactly once
    await expect(handle.done).rejects.toThrow('boom')
  })
})
