// run_script: the AI-facing entry into the scripting sandbox. Covers the
// runner bridge (runScriptForAi, stubbed worker) and the tool dispatch
// (input validation, output assembly, mutated marking).
import { describe, expect, it } from 'vitest'

import { runScriptForAi } from '../src/renderer/scripting/ai-script-runner'
import { setUniverAPI } from '../src/renderer/scripting/script-api-access'
import type { ScriptHost, WorkerMessage } from '../src/renderer/scripting/script-rpc'
import type { ScriptWorkerLike } from '../src/renderer/scripting/script-runner'
import { executeWorkbookTool, type SheetsSkillDeps } from '../src/renderer/ai/tools'

class FakeWorker implements ScriptWorkerLike {
  posted: { kind: 'run'; code: string }[] = []
  onmessage: ((event: { data: WorkerMessage }) => void) | null = null
  constructor(private readonly script: (worker: FakeWorker) => void = () => {}) {}
  postMessage(message: { kind: 'run'; code: string }): void {
    this.posted.push(message)
    if (message.kind === 'run') this.script(this)
  }
  terminate(): void {}
  /** the fake script side sends a message to the host */
  send(message: WorkerMessage): void {
    this.onmessage?.({ data: message })
  }
}

const stubHost: ScriptHost = { call: () => null }

describe('runScriptForAi', () => {
  it('captures Logger.log output and resolves when the script finishes', async () => {
    setUniverAPI(stubHost)
    const worker = new FakeWorker((w) => {
      w.send({ kind: 'log', text: 'line one' })
      w.send({ kind: 'log', text: 'line two' })
      w.send({ kind: 'done' })
    })
    const r = await runScriptForAi('Logger.log("x")', { createWorker: () => worker })
    expect(worker.posted[0]?.code).toBe('Logger.log("x")')
    expect(r.logs).toEqual(['line one', 'line two'])
    expect(r.error).toBeUndefined()
    expect(r.stopped).toBeUndefined()
  })

  it('reports a thrown script error and still resolves', async () => {
    setUniverAPI(stubHost)
    const worker = new FakeWorker((w) => {
      w.send({ kind: 'error', message: 'boom' })
    })
    const r = await runScriptForAi('throw new Error("boom")', { createWorker: () => worker })
    expect(r.error).toBe('boom')
    expect(r.logs).toEqual([])
  })

  it('fails cleanly when no workbook is open', async () => {
    setUniverAPI(null)
    const r = await runScriptForAi('anything', { createWorker: () => new FakeWorker() })
    expect(r.error).toMatch(/No workbook/)
    expect(r.logs).toEqual([])
  })
})

describe('run_script tool dispatch', () => {
  const deps = (runScript?: SheetsSkillDeps['runScript']): SheetsSkillDeps =>
    ({
      getActiveSheetInfo: { mode: 'none', sheetId: 's', sheetName: 'S', sheets: [] },
      readCells: () => ({}),
      readFormats: () => ({}),
      readSheetFeatures: () => '',
      findCells: () => ({ matches: [], truncated: false, incompleteSheets: [] }),
      selectRange: () => ({ ok: true, sheetName: 'S' }),
      tracePrecedents: () => ({ refs: [] }),
      traceDependents: () => ({ dependents: [], truncated: false }),
      proposeOperations: () => ({ ok: false, error: 'not configured' }),
      ...(runScript ? { runScript } : {}),
    }) as unknown as SheetsSkillDeps

  it('rejects empty code', async () => {
    const r = await executeWorkbookTool(
      { id: 'c1', name: 'run_script', input: { code: '   ' } },
      deps(),
    )
    if (r instanceof Promise) throw new Error('expected sync failure')
    expect(r.isError).toBe(true)
    expect(r.output).toMatch(/non-empty/)
  })

  it('rejects code beyond the storage cap', async () => {
    const r = await executeWorkbookTool(
      { id: 'c1', name: 'run_script', input: { code: 'x'.repeat(200_001) } },
      deps(),
    )
    if (r instanceof Promise) throw new Error('expected sync failure')
    expect(r.isError).toBe(true)
    expect(r.output).toMatch(/too long/)
  })

  it('is unavailable when the deps do not wire the runner', async () => {
    const r = await executeWorkbookTool(
      { id: 'c1', name: 'run_script', input: { code: '1' } },
      deps(),
    )
    const done = await r
    expect(done.isError).toBe(true)
    expect(done.output).toMatch(/not available/)
  })

  it('assembles logs into the output and marks the run as mutating', async () => {
    const r = await executeWorkbookTool(
      { id: 'c1', name: 'run_script', input: { code: 'ok' } },
      deps(async (code) => {
        expect(code).toBe('ok')
        return { logs: ['a', 'b'], error: 'threw late' }
      }),
    )
    const done = await r
    expect(done.output).toBe('Error: threw late\na\nb')
    expect(done.isError).toBe(true)
    expect(done.mutated).toBe(true)
  })

  it('says so when a clean run produced no output', async () => {
    const r = await executeWorkbookTool(
      { id: 'c1', name: 'run_script', input: { code: 'ok' } },
      deps(async () => ({ logs: [] })),
    )
    const done = await r
    expect(done.isError).toBe(false)
    expect(done.output).toMatch(/no output/)
  })
})
