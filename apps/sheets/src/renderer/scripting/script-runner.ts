/**
 * Main-thread side of script execution (issue #815).
 *
 * Owns the worker, answers its RPC calls against the host, and guarantees a run
 * cannot hang the app: a script that never finishes (or loops forever) is killed
 * when it exceeds the timeout, and Stop kills it immediately. Killing a worker is
 * also the only way to reclaim its memory, so each run gets a fresh one.
 */
import {
  SCRIPT_RPC_TIMEOUT_MS,
  dispatchRpc,
  type HostMessage,
  type ScriptHost,
  type WorkerMessage,
} from './script-rpc'

export interface ScriptRunHandlers {
  onLog: (text: string) => void
  onDone: (stopped: boolean) => void
  onError: (message: string) => void
}

/** The slice of `Worker` this runner uses — tests pass a stub. */
export interface ScriptWorkerLike {
  postMessage(message: HostMessage): void
  terminate(): void
  onmessage: ((event: { data: WorkerMessage }) => void) | null
}

export type WorkerFactory = () => ScriptWorkerLike

const defaultWorkerFactory: WorkerFactory = () =>
  new Worker(new URL('./script-worker.ts', import.meta.url), {
    type: 'module',
    name: 'genoffice-script',
  }) as unknown as ScriptWorkerLike

export interface ScriptRunner {
  run(code: string, handlers: ScriptRunHandlers): void
  stop(): void
  dispose(): void
  readonly running: boolean
}

export interface ScriptRunnerOptions {
  createWorker?: WorkerFactory
  timeoutMs?: number
}

export function createScriptRunner(
  host: ScriptHost,
  options: ScriptRunnerOptions = {},
): ScriptRunner {
  const createWorker = options.createWorker ?? defaultWorkerFactory
  const timeoutMs = options.timeoutMs ?? SCRIPT_RPC_TIMEOUT_MS

  let worker: ScriptWorkerLike | null = null
  let handlers: ScriptRunHandlers | null = null
  let timer: ReturnType<typeof setTimeout> | null = null

  const clearTimer = () => {
    if (timer) clearTimeout(timer)
    timer = null
  }

  const finish = (stopped: boolean) => {
    clearTimer()
    const done = handlers
    handlers = null
    if (worker) {
      worker.onmessage = null
      worker.terminate()
      worker = null
    }
    done?.onDone(stopped)
  }

  return {
    get running() {
      return handlers !== null
    },
    run(code, runHandlers) {
      if (handlers) this.stop()
      handlers = runHandlers
      worker = createWorker()
      worker.onmessage = (event) => {
        const message = event.data
        if (message.kind === 'log') {
          handlers?.onLog(message.text)
          return
        }
        if (message.kind === 'done') {
          finish(false)
          return
        }
        if (message.kind === 'error') {
          handlers?.onError(message.message)
          finish(false)
          return
        }
        // rpc: answer it; the worker awaits the matching id
        const result = dispatchRpc(host, message)
        worker?.postMessage(
          result.ok
            ? { kind: 'result', id: message.id, value: result.value }
            : { kind: 'fail', id: message.id, message: result.message },
        )
      }
      worker.postMessage({ kind: 'run', code })
      timer = setTimeout(() => {
        handlers?.onError(`Script stopped: it exceeded the ${Math.round(timeoutMs / 1000)}s limit`)
        finish(true)
      }, timeoutMs)
    },
    stop() {
      if (!handlers) return
      handlers.onLog('— stopped —')
      finish(true)
    },
    dispose() {
      finish(true)
    },
  }
}
