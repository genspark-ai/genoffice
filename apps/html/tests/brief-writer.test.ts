import { describe, expect, it } from 'vitest'
import type {
  AgentMessage,
  AgentStreamCallbacks,
  AgentStreamRequest,
  AgentTransport,
} from '@genoffice/agent-core'
import {
  buildBriefWriterRequest,
  extractBriefJson,
  parseBriefJson,
  streamBrief,
  transcriptOf,
} from '../src/renderer/ai/brief-writer'

const history: AgentMessage[] = [
  { role: 'user', text: 'Landing page for Acme water\n\n## Document\nfile: (untitled)' },
  {
    role: 'assistant',
    text: 'A few questions first.',
    toolCalls: [{ id: 't1', name: 'ask_clarification', input: { questions: [] } }],
  },
  { role: 'tool', results: [{ toolCallId: 't1', output: 'User answers:\nq1: young athletes' }] },
  {
    role: 'assistant',
    text: '',
    toolCalls: [{ id: 't2', name: 'plan_page', input: { mode: 'new' } }],
  },
]

function scripted(steps: Array<(cb: AgentStreamCallbacks) => void>) {
  const requests: AgentStreamRequest[] = []
  const transport: AgentTransport = {
    stream(request, cb) {
      requests.push(request)
      const step = steps.shift()
      queueMicrotask(() => step?.(cb))
      return { cancel: () => undefined }
    },
  }
  return { transport, requests }
}

describe('transcriptOf', () => {
  it('flattens user, assistant, tool calls and tool results in order', () => {
    const t = transcriptOf(history)
    expect(t).toContain('### User\nLanding page for Acme water')
    expect(t).toContain('### Assistant\nA few questions first.\n[called ask_clarification')
    expect(t).toContain('### Tool result\nUser answers:\nq1: young athletes')
    expect(t.indexOf('### User')).toBeLessThan(t.indexOf('### Tool result'))
  })

  it('keeps the newest turns when the transcript exceeds its budget', () => {
    const long: AgentMessage[] = Array.from({ length: 12 }, (_, i) => ({
      role: 'user' as const,
      text: `turn ${i} ${'x'.repeat(5000)}`,
    }))
    const t = transcriptOf(long)
    expect(t.startsWith('… (earlier turns omitted)')).toBe(true)
    expect(t).toContain('turn 11')
    expect(t).not.toContain('turn 0 ')
  })
})

describe('buildBriefWriterRequest', () => {
  it('puts the transcript, mode task and notes into the user message', () => {
    const { system, user } = buildBriefWriterRequest(
      { mode: 'new', notes: 'forest green, misty' },
      history,
      '\nReply in the same language.',
    )
    expect(system).toContain('brief writer')
    expect(system.endsWith('Reply in the same language.')).toBe(true)
    expect(user).toContain('young athletes')
    expect(user).toContain('A new page in an empty document')
    expect(user).toContain('## Notes from the assistant\nforest green, misty')
    expect(user).not.toContain('## Current page')
  })

  it('shows the current page only for restyle and extract', () => {
    const page = '<html><head><style>:root{--brief-bg:#fff}</style></head>'
    expect(buildBriefWriterRequest({ mode: 'restyle', page }, history, '').user).toContain(
      '## Current page (head and start of body)\n<html><head><style>',
    )
    expect(buildBriefWriterRequest({ mode: 'new', page }, history, '').user).not.toContain(
      '## Current page',
    )
  })
})

describe('extractBriefJson / parseBriefJson', () => {
  it('drops fences and chatter and tolerates trailing commas', () => {
    const raw = 'Here is the brief:\n```json\n{"core_hook": "x", "sections": [1,],}\n```'
    const { text, complete } = extractBriefJson(raw)
    expect(complete).toBe(true)
    expect(parseBriefJson(text)).toEqual({ core_hook: 'x', sections: [1] })
  })

  it('reports an unfinished object as incomplete', () => {
    expect(extractBriefJson('{"core_hook": "x", "sections": [').complete).toBe(false)
    expect(extractBriefJson('no braces at all').text).toBe('')
    expect(parseBriefJson('[1,2]')).toBeNull()
  })
})

describe('streamBrief', () => {
  it('assembles the JSON from text deltas and parses it', async () => {
    const { transport, requests } = scripted([
      (cb) => {
        cb.onDelta('{"core_hook": "Costs fell')
        cb.onDelta(' 20%", "style": {"tone": "executive"}}')
        cb.onStopReason?.('end_turn')
        cb.onDone()
      },
    ])
    const r = await streamBrief({ transport, system: 's', user: 'u' })
    expect(r).toEqual({
      ok: true,
      raw: { core_hook: 'Costs fell 20%', style: { tone: 'executive' } },
    })
    expect(requests[0]?.tools).toEqual([])
  })

  it('a dropped connection mid-object or an empty reply fails with the reason', async () => {
    const dropped = scripted([
      (cb) => {
        cb.onDelta('{"core_hook": "half')
        cb.onError('connection was dropped')
      },
    ])
    const r1 = await streamBrief({ transport: dropped.transport, system: 's', user: 'u' })
    expect(r1.ok).toBe(false)
    if (!r1.ok) expect(r1.error).toContain('connection was dropped')
    const empty = scripted([(cb) => cb.onDone()])
    const r2 = await streamBrief({ transport: empty.transport, system: 's', user: 'u' })
    expect(r2.ok).toBe(false)
  })
})
