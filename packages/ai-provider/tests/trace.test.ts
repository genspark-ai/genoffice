import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { startAiTrace } from '../src/trace'

const dirs: string[] = []
afterEach(() => {
  for (const d of dirs) rmSync(d, { recursive: true, force: true })
  dirs.length = 0
})

const request = {
  requestId: 'r1',
  provider: 'genspark',
  model: 'claude-opus-5-5',
  maxTokens: 32768,
  system: 'sys',
  messages: [{ role: 'user' as const, text: 'hi' }],
  tools: [],
}

describe('startAiTrace', () => {
  it('is off without GENOFFICE_AI_TRACE_DIR', () => {
    expect(startAiTrace(request, {})).toBeNull()
  })

  it('appends one JSON line per turn with the streamed response', () => {
    const dir = mkdtempSync(join(tmpdir(), 'ai-trace-'))
    dirs.push(dir)
    const env = { GENOFFICE_AI_TRACE_DIR: dir }
    const a = startAiTrace(request, env)!
    a.onReasoningDelta('think ')
    a.onDelta('Hel')
    a.onDelta('lo')
    a.onToolCall({ id: 't1', name: 'read_blocks', input: { from: 0 } })
    a.onStopReason('tool_use')
    a.onUsage({ inputTokens: 10, outputTokens: 5 })
    a.finish()
    a.finish()
    const b = startAiTrace({ ...request, requestId: 'r2' }, env)!
    b.finish('boom')

    const files = readdirSync(dir)
    expect(files).toHaveLength(1)
    expect(files[0]).toMatch(/^ai-trace-\d{8}-\d+\.jsonl$/)
    const lines = readFileSync(join(dir, files[0]!), 'utf8')
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l))
    expect(lines).toHaveLength(2)
    expect(lines[0]).toMatchObject({
      requestId: 'r1',
      model: 'claude-opus-5-5',
      system: 'sys',
      response: {
        text: 'Hello',
        reasoning: 'think ',
        toolCalls: [{ name: 'read_blocks' }],
        stopReason: 'tool_use',
        usage: { outputTokens: 5 },
      },
    })
    expect(lines[0].response.error).toBeUndefined()
    expect(lines[1]).toMatchObject({ requestId: 'r2', response: { error: 'boom' } })
  })
})
