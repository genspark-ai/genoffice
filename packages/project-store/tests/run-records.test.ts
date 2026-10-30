import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { ProjectStore } from '../src/store.js'

function makeTempDir(): string {
  return mkdtempSync(join(tmpdir(), 'project-store-runs-'))
}

describe('agent run records', () => {
  let tmpDir: string
  let store: ProjectStore

  beforeEach(() => {
    tmpDir = makeTempDir()
    store = new ProjectStore(tmpDir)
    store.ensureDefaultProject()
  })

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true })
  })

  it('appends run events as one JSONL line each, newest-last in the file', () => {
    store.appendRunEvent('default', 'run-1', {
      kind: 'start',
      instruction: 'build the report',
      chatId: 'kb-ask',
    })
    store.appendRunEvent('default', 'run-1', { kind: 'finish', outcome: 'done' })

    const raw = readFileSync(join(tmpDir, 'projects', 'default', 'runs', 'run-1.jsonl'), 'utf8')
    const lines = raw.trim().split('\n')
    expect(lines).toHaveLength(2)
    expect(JSON.parse(lines[0]!)).toMatchObject({
      kind: 'start',
      instruction: 'build the report',
      chatId: 'kb-ask',
    })
    expect(JSON.parse(lines[0]!).ts).toBeTruthy()
    expect(JSON.parse(lines[1]!)).toMatchObject({ kind: 'finish', outcome: 'done' })
  })

  it('folds events into a summary: started, done, with the transcript reference', () => {
    store.appendRunEvent('default', 'run-1', {
      kind: 'start',
      instruction: 'summarize the notes',
      chatId: 'kb-ask',
    })
    store.appendRunEvent('default', 'run-1', { kind: 'finish', outcome: 'done' })

    const runs = store.listRuns('default')
    expect(runs).toHaveLength(1)
    expect(runs[0]).toMatchObject({
      runId: 'run-1',
      chatId: 'kb-ask',
      instruction: 'summarize the notes',
      status: 'done',
    })
    expect(runs[0]!.startedAt).toBeTruthy()
    expect(runs[0]!.endedAt).toBeTruthy()
  })

  it('a run with no terminal event reads back as interrupted (the restart case)', () => {
    // simulate a crash: start written, the window closed before any terminal
    store.appendRunEvent('default', 'run-7', {
      kind: 'start',
      instruction: 'build the deck',
      chatId: 'kb-ask-1731',
    })
    // a *new* store instance over the same directory = the next launch
    const reopened = new ProjectStore(tmpDir)
    const runs = reopened.listRuns('default')
    expect(runs).toHaveLength(1)
    expect(runs[0]).toMatchObject({
      runId: 'run-7',
      status: 'interrupted',
      instruction: 'build the deck',
    })
    expect(runs[0]!.endedAt).toBeUndefined()
  })

  it('cancelled and discarded finishes fold to their own statuses', () => {
    store.appendRunEvent('default', 'a', { kind: 'start', instruction: 'one' })
    store.appendRunEvent('default', 'a', { kind: 'finish', outcome: 'cancelled' })
    store.appendRunEvent('default', 'b', { kind: 'start', instruction: 'two' })
    store.appendRunEvent('default', 'b', { kind: 'finish', outcome: 'discarded' })
    const byId = new Map(store.listRuns('default').map((r) => [r.runId, r.status]))
    expect(byId.get('a')).toBe('cancelled')
    expect(byId.get('b')).toBe('discarded')
  })

  it('an error terminal folds to the error status with its message', () => {
    store.appendRunEvent('default', 'run-2', { kind: 'start', instruction: 'x' })
    store.appendRunEvent('default', 'run-2', { kind: 'error', error: 'provider is down' })
    const runs = store.listRuns('default')
    expect(runs[0]).toMatchObject({ runId: 'run-2', status: 'error', error: 'provider is down' })
  })

  it('a run that errored and then finished reports the finish (last terminal wins)', () => {
    store.appendRunEvent('default', 'r', { kind: 'start', instruction: 'retry story' })
    store.appendRunEvent('default', 'r', { kind: 'error', error: 'first attempt failed' })
    store.appendRunEvent('default', 'r', { kind: 'finish', outcome: 'done' })
    expect(store.listRuns('default')[0]!.status).toBe('done')
  })

  it('lists runs newest first and caps long instruction text', () => {
    store.appendRunEvent('default', 'run-1', {
      kind: 'start',
      instruction: 'first task',
      ts: '2026-10-29T09:00:00.000Z',
    })
    store.appendRunEvent('default', 'run-1', { kind: 'finish', outcome: 'done' })
    store.appendRunEvent('default', 'run-2', {
      kind: 'start',
      instruction: `long ${'x'.repeat(500)}`,
      ts: '2026-10-29T10:00:00.000Z',
    })

    const runs = store.listRuns('default')
    expect(runs.map((r) => r.runId)).toEqual(['run-2', 'run-1'])
    expect(runs[1]!.instruction.length).toBeLessThanOrEqual(200)
  })

  it('skips corrupt lines and unreadable run files without failing the listing', () => {
    const runsDir = join(tmpDir, 'projects', 'default', 'runs')
    store.appendRunEvent('default', 'good', { kind: 'start', instruction: 'fine' })
    // a torn line and a file with no start event
    writeFileSync(join(runsDir, 'good.jsonl'), 'not json at all\n', { flag: 'a' })
    writeFileSync(join(runsDir, 'no-start.jsonl'), '{"kind":"finish","outcome":"done"}\n')
    writeFileSync(join(runsDir, 'not-a-run.txt'), 'ignored\n')

    const runs = store.listRuns('default')
    expect(runs.map((r) => r.runId)).toEqual(['good'])
    expect(runs[0]!.status).toBe('interrupted')
  })

  it('deleteRun removes one run record; the others stay', () => {
    store.appendRunEvent('default', 'run-1', { kind: 'start', instruction: 'one' })
    store.appendRunEvent('default', 'run-2', { kind: 'start', instruction: 'two' })
    store.deleteRun('default', 'run-1')
    expect(store.listRuns('default').map((r) => r.runId)).toEqual(['run-2'])
    expect(existsSync(join(tmpDir, 'projects', 'default', 'runs', 'run-1.jsonl'))).toBe(false)
    // deleting something already gone is a no-op, not an error
    expect(() => store.deleteRun('default', 'run-1')).not.toThrow()
  })

  it('rejects ids that could escape the store directory (fail-closed)', () => {
    expect(() =>
      store.appendRunEvent('default', '../escape', { kind: 'start', instruction: 'x' }),
    ).toThrow(/runId/)
    expect(() => store.listRuns('../escape')).toThrow(/projectId/)
    expect(() => store.deleteRun('default', 'a/b')).toThrow(/runId/)
    expect(() => store.appendRunEvent('..', 'run-1', { kind: 'start' })).toThrow(/projectId/)
  })

  it('a write failure warns instead of throwing (run records are telemetry)', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {})
    try {
      // the runs "directory" is a file: ensureDir cannot make it
      writeFileSync(join(tmpDir, 'projects', 'default', 'runs'), 'in the way')
      expect(() =>
        store.appendRunEvent('default', 'run-1', { kind: 'start', instruction: 'x' }),
      ).not.toThrow()
      expect(warn).toHaveBeenCalled()
    } finally {
      warn.mockRestore()
    }
  })
})
