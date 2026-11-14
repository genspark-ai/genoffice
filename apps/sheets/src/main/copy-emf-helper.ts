/**
 * The clipboard helper that appends the EMF picture of a copied range.
 *
 * The helper is a separate process: it renders the range with GDI and hands
 * the metafile to the Windows clipboard. It talks over stdin/stdout with one
 * JSON line each way.
 *
 * Two things the spawn has to get right:
 *
 * - The helper can die before it reads the payload (a crash, an antivirus
 *   kill, the timeout). The write then fails with EPIPE, and a stream error
 *   nobody listens for takes the whole main process down with it — the copy
 *   is an enhancement, so it must never break the app.
 * - The append is fire-and-forget, so two quick copies can finish out of
 *   order and the older one can land beside the newer clipboard content. The
 *   helper checks the clipboard sequence number before it appends (see
 *   win-clipboard-helper/src/main.rs), and only one helper runs at a time: a
 *   newer request kills the running one.
 */
import { spawn, type ChildProcess } from 'node:child_process'

export interface CopyEmfRunOptions {
  /** how long the helper may take before it is killed */
  timeoutMs?: number
  /** the spawn, injectable so a test can hand in a process that dies at once */
  spawnFn?: typeof spawn
}

const DEFAULT_TIMEOUT_MS = 5_000

/** The helper currently running; a newer request supersedes it. */
let running: ChildProcess | null = null

/**
 * Runs the helper once and resolves whether it appended the picture.
 *
 * Never rejects: the EMF is an enhancement on top of the copy's text
 * formats, so every failure resolves to `{ ok: false }`.
 */
export async function runCopyEmfHelper(
  helper: string,
  payload: string,
  options: CopyEmfRunOptions = {},
): Promise<{ ok: boolean }> {
  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS
  const spawnFn = options.spawnFn ?? spawn
  // A newer request supersedes the running one: two helpers appending at the
  // same time could put the older picture beside the newer clipboard content.
  if (running) {
    running.kill()
    running = null
  }
  return await new Promise((resolve) => {
    const child = spawnFn(helper, [], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true })
    running = child
    let settled = false
    const settle = (result: { ok: boolean }): void => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      if (running === child) running = null
      resolve(result)
    }
    const timer = setTimeout(() => {
      child.kill()
      settle({ ok: false })
    }, timeoutMs)
    child.on('error', () => settle({ ok: false }))
    // The helper can die before it reads the payload: the write then fails with
    // EPIPE, and an unhandled stream error would take the main process down.
    child.stdin.on('error', () => settle({ ok: false }))
    let output = ''
    child.stdout.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      output += chunk
    })
    child.on('close', () => {
      try {
        const line = output.trim().split('\n').pop() ?? ''
        const result = JSON.parse(line) as { ok?: boolean }
        settle({ ok: result.ok === true })
      } catch {
        settle({ ok: false })
      }
    })
    child.stdin.write(payload)
    child.stdin.end()
  })
}
