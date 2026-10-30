import { describe, expect, it } from 'vitest'

import type { ChatMessage } from '@genoffice/project-store'

import {
  interruptedOf,
  parseRunRecords,
  runDiscardRecord,
  runErrorRecord,
  runFinishRecord,
  runStartRecord,
  RUNS_CHAT_ID,
} from '../src/renderer/src/ask-runs'

function line(text: string, over: Partial<ChatMessage> = {}): ChatMessage {
  return { seq: 0, ts: '2026-10-30T09:00:00.000Z', role: 'assistant', text, ...over }
}

describe('run-record lines', () => {
  it('carry the run id, the lifecycle kind and the fields the list needs', () => {
    const start = JSON.parse(runStartRecord('run-1', 'build the report', 'kb-ask'))
    expect(start).toEqual({
      v: 1,
      runId: 'run-1',
      kind: 'start',
      instruction: 'build the report',
      chatId: 'kb-ask',
    })
    expect(JSON.parse(runFinishRecord('run-1', 'done'))).toEqual({
      v: 1,
      runId: 'run-1',
      kind: 'finish',
      outcome: 'done',
    })
    expect(JSON.parse(runFinishRecord('run-1', 'cancelled'))).toMatchObject({
      outcome: 'cancelled',
    })
    expect(JSON.parse(runDiscardRecord('run-1'))).toMatchObject({ outcome: 'discarded' })
    expect(JSON.parse(runErrorRecord('run-1', 'provider is down'))).toMatchObject({
      kind: 'error',
      error: 'provider is down',
    })
  })

  it('caps the stored instruction and error text', () => {
    const start = JSON.parse(runStartRecord('r', 'x'.repeat(600), 'kb-ask'))
    expect(start.instruction).toHaveLength(200)
    const error = JSON.parse(runErrorRecord('r', 'e'.repeat(900)))
    expect(error.error).toHaveLength(500)
  })

  it('uses a reserved chat id that the thread list cannot match', () => {
    // panelThreads lists kb-ask and kb-ask-*; the runs chat must stay invisible there
    expect(RUNS_CHAT_ID).toBe('kb-runs')
    expect(RUNS_CHAT_ID.startsWith('kb-ask')).toBe(false)
  })
})

describe('folding stored run records', () => {
  it('a start with a finish folds to done with the transcript reference', () => {
    const runs = parseRunRecords([
      line(runStartRecord('run-1', 'summarize the notes', 'kb-ask'), {
        seq: 0,
        ts: '2026-10-30T09:00:00.000Z',
      }),
      line(runFinishRecord('run-1', 'done'), { seq: 1, ts: '2026-10-30T09:01:00.000Z' }),
    ])
    expect(runs).toHaveLength(1)
    expect(runs[0]).toMatchObject({
      runId: 'run-1',
      chatId: 'kb-ask',
      instruction: 'summarize the notes',
      status: 'done',
      startedAt: '2026-10-30T09:00:00.000Z',
      endedAt: '2026-10-30T09:01:00.000Z',
    })
  })

  it('a start without a terminal (the closed-window case) reads interrupted', () => {
    const runs = parseRunRecords([
      line(runStartRecord('run-7', 'build the deck', 'kb-ask-1731'), {
        seq: 0,
        ts: '2026-10-29T18:00:00.000Z',
      }),
    ])
    expect(runs[0]).toMatchObject({ runId: 'run-7', status: 'interrupted' })
    expect(runs[0]!.endedAt).toBeUndefined()
    expect(interruptedOf(runs)).toHaveLength(1)
  })

  it('cancelled, discarded and error terminals fold to their own statuses', () => {
    const runs = parseRunRecords([
      line(runStartRecord('a', 'one'), { seq: 0 }),
      line(runFinishRecord('a', 'cancelled'), { seq: 1 }),
      line(runStartRecord('b', 'two'), { seq: 2 }),
      line(runDiscardRecord('b'), { seq: 3 }),
      line(runStartRecord('c', 'three'), { seq: 4 }),
      line(runErrorRecord('c', 'gateway closed'), { seq: 5 }),
    ])
    const byId = new Map(runs.map((r) => [r.runId, r]))
    expect(byId.get('a')!.status).toBe('cancelled')
    expect(byId.get('b')!.status).toBe('discarded')
    expect(byId.get('c')).toMatchObject({ status: 'error', error: 'gateway closed' })
    expect(interruptedOf(runs)).toHaveLength(0)
  })

  it('the last terminal wins (an error retried to success reports done)', () => {
    const runs = parseRunRecords([
      line(runStartRecord('r', 'retry story'), { seq: 0 }),
      line(runErrorRecord('r', 'first attempt failed'), { seq: 1 }),
      line(runFinishRecord('r', 'done'), { seq: 2 }),
    ])
    expect(runs[0]!.status).toBe('done')
  })

  it('skips stored turns that are not run records, without failing', () => {
    const runs = parseRunRecords([
      line('a normal answer with JSON-ish text {"v":1}', { role: 'assistant', seq: 0 }),
      line('{"v":2,"runId":"x","kind":"start"}', { seq: 1 }),
      line('not json', { role: 'user', seq: 2 }),
      line(runStartRecord('real', 'kept', 'kb-ask'), { seq: 3 }),
    ])
    expect(runs.map((r) => r.runId)).toEqual(['real'])
  })

  it('lists the newest run first', () => {
    const runs = parseRunRecords([
      line(runStartRecord('run-1', 'older'), { seq: 0, ts: '2026-10-30T09:00:00.000Z' }),
      line(runStartRecord('run-2', 'newer'), { seq: 1, ts: '2026-10-30T10:00:00.000Z' }),
    ])
    expect(runs.map((r) => r.runId)).toEqual(['run-2', 'run-1'])
  })

  it('a terminal for an unknown run id is ignored (no start, no row)', () => {
    const runs = parseRunRecords([line(runFinishRecord('ghost', 'done'), { seq: 0 })])
    expect(runs).toHaveLength(0)
  })
})
