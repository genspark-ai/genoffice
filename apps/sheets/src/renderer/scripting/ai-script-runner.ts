/**
 * AI-facing entry into the script sandbox (issue #815): the model writes the
 * same code the script editor runs, through the same pipeline — dedicated
 * worker, CSP'd sandbox, host-side method allowlist, run timeout. A fresh
 * worker and host per run, so handles and globals never leak between runs,
 * and an AI run cannot see the editor dialog's in-flight state.
 */
import { SCRIPT_RPC_TIMEOUT_MS } from './script-rpc'
import { FacadeScriptHost } from './script-host'
import { createScriptRunner, type ScriptRunner, type WorkerFactory } from './script-runner'
import { getUniverAPI } from './script-api-access'

export interface AiScriptRunResult {
  /** Logger.log output, oldest first, capped to the last 400 lines */
  logs: string[]
  /** first script-visible failure (thrown error or timeout) */
  error?: string
  /** true when the run hit the timeout (or was killed) rather than finishing */
  stopped?: boolean
}

/** Log ring size: enough context for the model to iterate, small enough to keep
 *  a runaway `for(;;) Logger.log()` from flooding the tool result. */
const MAX_LOG_LINES = 400

/** Run one script to completion. Resolves only when the run settles (done,
 *  thrown error, or timeout) — the runner's own guarantees, not a second timer. */
export async function runScriptForAi(
  code: string,
  options: { timeoutMs?: number; createWorker?: WorkerFactory } = {},
): Promise<AiScriptRunResult> {
  const api = getUniverAPI()
  if (!api) {
    return { logs: [], error: 'No workbook is open — open or create one first.' }
  }
  const logs: string[] = []
  let error: string | undefined
  const runnerOptions: { timeoutMs: number; createWorker?: WorkerFactory } = {
    timeoutMs: options.timeoutMs ?? SCRIPT_RPC_TIMEOUT_MS,
  }
  if (options.createWorker) runnerOptions.createWorker = options.createWorker
  const runner: ScriptRunner = createScriptRunner(new FacadeScriptHost(api), runnerOptions)
  const result = await new Promise<AiScriptRunResult>((resolve) => {
    runner.run(code, {
      onLog: (text) => {
        logs.push(text)
        if (logs.length > MAX_LOG_LINES) logs.splice(0, logs.length - MAX_LOG_LINES)
      },
      onError: (message) => {
        error = message
      },
      onDone: (wasStopped) => {
        const out: AiScriptRunResult = { logs: [...logs] }
        if (error !== undefined) out.error = error
        if (wasStopped) out.stopped = true
        resolve(out)
      },
    })
  })
  runner.dispose()
  return result
}
