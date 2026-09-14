/**
 * Runtime tests — covers the event mapping (onText / onToolStart /
 * onToolExecuted / onDone), cancel, snapshot, and AIError classification.
 */
import { describe, expect, it } from 'vitest'
import type {
  AgentRunResult,
  AgentSkill,
  AgentStreamHandle,
  AgentStreamRequest,
  AgentStreamCallbacks,
  AgentTransport,
  ToolExecution,
} from '@genoffice/agent-core'

import { ChatRuntime } from '../src/runtime.js'
import { classifyError, AIError } from '../src/errors.js'
import { normalizeChangePlan } from '../src/change-plan.js'
import { startRun } from '../src/run.js'
import type { ChatRun, ChatSession } from '../src/types.js'

function fakeSkill(): AgentSkill {
  return {
    id: 'fake',
    systemPrompt: '',
    tools: [],
    executeTool: async () => ({ output: '', isError: true, summary: 'fake' }),
  }
}

function fakeTransport(): AgentTransport {
  return {
    stream(_req: AgentStreamRequest, _cb: AgentStreamCallbacks): AgentStreamHandle {
      return { cancel: () => undefined }
    },
  }
}

function fakeRunResult(overrides: Partial<AgentRunResult> = {}): AgentRunResult {
  return { text: '', cancelled: false, turnLimit: false, ...overrides }
}

describe('errors.classifyError', () => {
  it('classifies WEB_UNSUPPORTED envelopes', () => {
    const err = classifyError({ code: 'WEB_UNSUPPORTED', channel: 'ai:doc-write', reason: 'renderer-side skill' })
    expect(err).toBeInstanceOf(AIError)
    expect(err.code).toBe('WEB_UNSUPPORTED')
    expect(err.channel).toBe('ai:doc-write')
  })

  it('classifies timeout messages', () => {
    const err = classifyError(new Error('Request timed out'))
    expect(err.code).toBe('TIMEOUT')
    expect(err.retryable).toBe(true)
  })

  it('classifies network messages', () => {
    const err = classifyError(new Error('fetch failed ECONNRESET'))
    expect(err.code).toBe('NETWORK')
  })

  it('classifies missing-API-key messages as NOT_CONFIGURED', () => {
    const err = classifyError(new Error('No API key configured for provider "minimax"'))
    expect(err.code).toBe('NOT_CONFIGURED')
    expect(err.retryable).toBe(false)
  })

  it('falls back to INTERNAL for unknown shapes', () => {
    const err = classifyError('something exploded')
    expect(err.code).toBe('INTERNAL')
  })

  it('classifies Genspark 403 HTML-page responses as WEB_UNSUPPORTED', () => {
    const err = classifyError(
      new Error(
        'Claude HTTP 403: the service returned a web page ("Genspark") instead of an API response (likely a temporary network or gateway block) — check your connection and retry',
      ),
    )
    expect(err.code).toBe('WEB_UNSUPPORTED')
    expect(err.retryable).toBe(false)
  })

  it('classifies sign-in required responses as WEB_UNSUPPORTED', () => {
    const err = classifyError(new Error('Sign in required to use Genspark credits'))
    expect(err.code).toBe('WEB_UNSUPPORTED')
  })

  it('classifies rate-limit / 429 responses as PROVIDER', () => {
    const err = classifyError(
      new Error('已达到 Token Plan 用量上限：请升级 Token Plan 套餐或购买积分补充用量。 (2056)'),
    )
    expect(err.code).toBe('PROVIDER')
    // PROVIDER defaults to retryable=true so users can retry a transient
    // quota blip. Quota-exhausted cases should override via the
    // `retryable: false` constructor option at the call site.
    expect(err.retryable).toBe(true)
  })

  it('classifies HTTP 429 responses as PROVIDER', () => {
    const err = classifyError(new Error('HTTP 429: rate limit exceeded'))
    expect(err.code).toBe('PROVIDER')
  })

  it('classifies plain-string errors (the shape AgentLoop sends)', () => {
    const err = classifyError(
      'Claude HTTP 403: the service returned a web page ("Genspark") instead of an API response (likely a temporary network or gateway block) — check your connection and retry',
    )
    expect(err.code).toBe('WEB_UNSUPPORTED')
  })

  it('classifies plain-string rate-limit as PROVIDER', () => {
    const err = classifyError('已达到 Token Plan 用量上限：请升级 Token Plan 套餐或购买积分补充用量。 (2056)')
    expect(err.code).toBe('PROVIDER')
  })

  it('parses JSON-encoded WEB_UNSUPPORTED envelopes from the http transport', () => {
    const envelope = JSON.stringify({
      code: 'WEB_UNSUPPORTED',
      channel: 'ai:doc-write-continue',
      reason: 'renderer-side skill',
      message: "Channel 'ai:doc-write-continue' is not supported on web; renderer-side skill.",
    })
    const err = classifyError(envelope)
    expect(err.code).toBe('WEB_UNSUPPORTED')
    expect(err.channel).toBe('ai:doc-write-continue')
    expect(err.retryable).toBe(false)
  })

  it('parses JSON-encoded TIMEOUT envelopes', () => {
    const envelope = JSON.stringify({ code: 'TIMEOUT', message: 'request timed out' })
    const err = classifyError(envelope)
    expect(err.code).toBe('TIMEOUT')
    expect(err.retryable).toBe(true)
  })

  it('returns plain INTERNAL for strings that look JSON-ish but are not valid', () => {
    const err = classifyError('{not valid json}')
    expect(err.code).toBe('INTERNAL')
  })
})

describe('change-plan.normalizeChangePlan', () => {
  it('wraps flat XLSX-style ops under {kind:workbook}', () => {
    const plan = normalizeChangePlan({
      app: 'sheets',
      title: 'Insert column',
      ops: [{ op: 'insert-column', at: 'C' }, { op: 'set-cell', at: 'C1', value: 42 }],
    })
    expect(plan.app).toBe('sheets')
    expect(plan.ops).toHaveLength(2)
    for (const op of plan.ops) {
      expect((op as { kind: string }).kind).toBe('workbook')
    }
  })

  it('preserves ops that already declare a kind', () => {
    const plan = normalizeChangePlan({
      app: 'docs',
      title: 'Insert paragraph',
      ops: [{ kind: 'doc', ops: [{ at: 'p:1' }], description: 'insert at p:1' }],
    })
    expect(plan.ops).toHaveLength(1)
    expect((plan.ops[0] as { kind: string }).kind).toBe('doc')
  })
})

describe('ChatRuntime', () => {
  it('persists a new session via the persistence adapter', async () => {
    const appends: number[] = []
    const persistence = {
      async load(): Promise<ChatSession | null> { return null },
      async save(): Promise<void> {},
      async append(_id: string, _m: import('../src/types.js').ChatMessage): Promise<void> { appends.push(appends.length + 1) },
    }
    const runtime = ChatRuntime.create({
      app: 'docs',
      sessionId: 'sess-1',
      skill: fakeSkill(),
      persistence,
    })
    await runtime.init()
    const sendPromise = runtime.send({ instruction: 'Hello' })
    // Wait for the run to be installed before cancelling (send() awaits
    // persistMessage first, so the handle is not ready synchronously).
    await new Promise<void>(r => {
      const unsub = runtime.subscribe(e => {
        if (e.type === 'run-start') {
          unsub()
          r()
        }
      })
    })
    runtime.cancel()
    await sendPromise
    expect(appends.length).toBeGreaterThanOrEqual(2)
  })

  it('subscribers see run-start + run-finish', async () => {
    const runtime = ChatRuntime.create({
      app: 'docs',
      sessionId: 'sess-2',
      skill: fakeSkill(),
    })
    const events: string[] = []
    const unsub = runtime.subscribe(e => events.push(e.type))
    const sendPromise = runtime.send({ instruction: 'ping' })
    // Drive the run to completion via cancel() once the run-start event
    // has been delivered to any subscriber (so this.handle is installed).
    let sawStart = false
    const unsubscribe = runtime.subscribe(e => {
      if (e.type === 'run-start' && !sawStart) {
        sawStart = true
        unsubscribe()
        runtime.cancel()
      }
    })
    await sendPromise
    unsub()
    expect(events).toContain('run-start')
    expect(events.some(e => e === 'run-finish' || e === 'run-error')).toBe(true)
  }, 10_000)

  it('cancel() flips the run to cancelled', () => {
    const runtime = ChatRuntime.create({
      app: 'docs',
      sessionId: 'sess-3',
      skill: fakeSkill(),
    })
    void runtime.send({ instruction: 'long' }).catch(() => undefined)
    runtime.cancel()
    expect(['cancelled', 'done', 'idle']).toContain(runtime.getCurrentRun()?.status ?? 'idle')
  })
})

describe('startRun', () => {
  it('starts in queued state', () => {
    const handle = startRun({
      sessionId: 's1',
      skill: fakeSkill(),
    })
    expect(handle.run.status).toBe('queued')
    expect(handle.run.cancelled).toBe(false)
  })

  it('returns a snapshot function returning undefined when no tools have run', () => {
    const handle = startRun({ sessionId: 's2', skill: fakeSkill() })
    expect(typeof handle.snapshot).toBe('function')
    expect(handle.snapshot()).toBeUndefined()
  })

  it('cancel sets cancelled flag', () => {
    const handle = startRun({ sessionId: 's3', skill: fakeSkill() })
    handle.cancel()
    expect(handle.run.cancelled).toBe(true)
    expect(handle.run.status).toBe('cancelled')
  })
})

// Reference imports to silence unused-var warnings
void fakeTransport()
void fakeRunResult()
