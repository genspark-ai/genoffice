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

function scriptedTransport(script: Array<(cb: AgentStreamCallbacks) => void>): AgentTransport {
  let turn = 0
  return {
    stream(
      _request: { system: string; messages: AgentMessage[]; tools: unknown[] },
      cb: AgentStreamCallbacks,
    ) {
      const step = script[turn++]
      if (step) queueMicrotask(() => step(cb))
      return { cancel: () => queueMicrotask(() => cb.onDone()) }
    },
  }
}

function makeSkill(execute: (call: AgentToolCall) => ToolExecution): AgentSkill {
  return {
    id: 'test',
    systemPrompt: 'system',
    tools: [{ name: 'read_file', description: 'd', inputSchema: { type: 'object' } }],
    executeTool: execute,
  }
}

const flush = () => new Promise((r) => setTimeout(r, 0))

describe('tool output sanitize-on-store', () => {
  it('stores sanitized tool output in history so file secrets never reach the model', async () => {
    const secretOutput = 'config dump:\npassword=hunter2\napi_key=AKIA1234567890\nok'
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'read_file', input: { path: 'config.env' } })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('done')
        cb.onDone()
      },
    ])
    const onToolExecuted = vi.fn()
    const onDone = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => ({ output: secretOutput, summary: 'read' })),
      events: { onToolExecuted, onDone },
    })
    loop.run('read the config')
    await flush()
    await flush()

    const toolMsg = loop.messages.find((m) => m.role === 'tool') as Extract<
      AgentMessage,
      { role: 'tool' }
    >
    expect(toolMsg).toBeDefined()
    const stored = toolMsg.results[0]!.output
    expect(stored).not.toContain('hunter2')
    expect(stored).not.toContain('AKIA1234567890')
    expect(stored).toContain('[REDACTED_SECURE_TOKEN]')
    expect(onDone).toHaveBeenCalledWith({ text: 'done', cancelled: false, turnLimit: false })
  })
})
