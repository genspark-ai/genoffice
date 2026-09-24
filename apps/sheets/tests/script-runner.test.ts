// Script runner (issue #815): owns the worker, answers its RPC calls, and must be
// able to kill a script that hangs — that is the whole reason scripts run in a
// worker instead of in the renderer.
import { describe, expect, it, vi } from 'vitest'
import { createScriptRunner, type ScriptWorkerLike } from '../src/renderer/scripting/script-runner'
import type { HostMessage, WorkerMessage, ScriptHost } from '../src/renderer/scripting/script-rpc'

class FakeWorker implements ScriptWorkerLike {
  posted: HostMessage[] = []
  terminated = false
  onmessage: ((event: { data: WorkerMessage }) => void) | null = null
  /** messages the fake script sends back, replayed on demand by the tests */
  constructor(private readonly script: (worker: FakeWorker) => void = () => {}) {}
  postMessage(message: HostMessage): void {
    this.posted.push(message)
    if (message.kind === 'run') this.script(this)
  }
  terminate(): void {
    this.terminated = true
  }
  send(message: WorkerMessage): void {
    this.onmessage?.({ data: message })
  }
}

const okHost: ScriptHost = { call: () => ({ __h: 1, __k: 'workbook' }) }

describe('createScriptRunner', () => {
  it('starts the script and reports logs and completion', () => {
    let worker: FakeWorker | null = null
    const runner = createScriptRunner(okHost, {
      createWorker: () =>
        (worker = new FakeWorker((w) => {
          w.send({ kind: 'log', text: 'first' })
          w.send({ kind: 'log', text: 'second' })
          w.send({ kind: 'done' })
        })),
    })
    const logs: string[] = []
    const onDone = vi.fn()
    runner.run('Logger.log(1)', { onLog: (t) => logs.push(t), onDone, onError: vi.fn() })

    expect(worker!.posted[0]).toEqual({ kind: 'run', code: 'Logger.log(1)' })
    expect(logs).toEqual(['first', 'second'])
    expect(onDone).toHaveBeenCalledWith(false)
    expect(worker!.terminated).toBe(true)
    expect(runner.running).toBe(false)
  })

  it('answers the worker RPC with the host result', () => {
    let worker: FakeWorker | null = null
    const host: ScriptHost = { call: (_target, method) => `host:${method}` }
    const runner = createScriptRunner(host, {
      createWorker: () =>
        (worker = new FakeWorker((w) => {
          w.send({ kind: 'rpc', id: 9, target: 'app', method: 'getActiveSpreadsheet', args: [] })
        })),
    })
    runner.run('code', { onLog: vi.fn(), onDone: vi.fn(), onError: vi.fn() })

    expect(worker!.posted.at(-1)).toEqual({
      kind: 'result',
      id: 9,
      value: 'host:getActiveSpreadsheet',
    })
  })

  it('turns a host failure into a fail message the script can catch', () => {
    let worker: FakeWorker | null = null
    const host: ScriptHost = {
      call: () => {
        throw new Error('no workbook')
      },
    }
    const runner = createScriptRunner(host, {
      createWorker: () =>
        (worker = new FakeWorker((w) => {
          w.send({ kind: 'rpc', id: 3, target: 'app', method: 'getActiveSpreadsheet', args: [] })
        })),
    })
    runner.run('code', { onLog: vi.fn(), onDone: vi.fn(), onError: vi.fn() })
    expect(worker!.posted.at(-1)).toEqual({ kind: 'fail', id: 3, message: 'no workbook' })
  })

  it('surfaces a script error and still releases the worker', () => {
    let worker: FakeWorker | null = null
    const runner = createScriptRunner(okHost, {
      createWorker: () =>
        (worker = new FakeWorker((w) => {
          w.send({ kind: 'error', message: 'ReferenceError: nope' })
        })),
    })
    const onError = vi.fn()
    const onDone = vi.fn()
    runner.run('nope()', { onLog: vi.fn(), onDone, onError })
    expect(onError).toHaveBeenCalledWith('ReferenceError: nope')
    expect(onDone).toHaveBeenCalledWith(false)
    expect(worker!.terminated).toBe(true)
  })

  it('kills the worker on Stop', () => {
    let worker: FakeWorker | null = null
    const runner = createScriptRunner(okHost, {
      createWorker: () => (worker = new FakeWorker()),
    })
    const onDone = vi.fn()
    const logs: string[] = []
    runner.run('while (true) {}', { onLog: (t) => logs.push(t), onDone, onError: vi.fn() })
    expect(runner.running).toBe(true)

    runner.stop()
    expect(worker!.terminated).toBe(true)
    expect(onDone).toHaveBeenCalledWith(true)
    expect(logs.at(-1)).toContain('stopped')
    expect(runner.running).toBe(false)
  })

  it('kills a script that runs longer than the timeout', () => {
    vi.useFakeTimers()
    try {
      let worker: FakeWorker | null = null
      const runner = createScriptRunner(okHost, {
        createWorker: () => (worker = new FakeWorker()),
        timeoutMs: 1000,
      })
      const onError = vi.fn()
      const onDone = vi.fn()
      runner.run('while (true) {}', { onLog: vi.fn(), onDone, onError })
      vi.advanceTimersByTime(1001)
      expect(worker!.terminated).toBe(true)
      expect(onError.mock.calls[0]![0]).toMatch(/exceeded the 1s limit/)
      expect(onDone).toHaveBeenCalledWith(true)
    } finally {
      vi.useRealTimers()
    }
  })
})
