/**
 * Font directory scan on a worker thread with a deadline. The scan is a handful of
 * readdir/stat calls (~ms), but Office's FontCache under Group Containers has stalled
 * readdir for minutes on some machines, and a blocking scan there freezes the main
 * process and every window with it. Directories that miss the deadline are skipped.
 */
import { MessageChannel, Worker, receiveMessageOnPort } from 'node:worker_threads'

export type ScanTask =
  /** Font files directly inside the dir */
  | { kind: 'flat'; dir: string }
  /** Office cloud fonts: <base>/<cache-id>/<sub>/<Family Name>/<numeric-id>.ttf */
  | { kind: 'cloud'; base: string; sub: string }

export type ScanResult =
  /** File names (not paths) of the font files found; [] when the dir is unreadable */
  | { kind: 'flat'; files: string[] }
  /** Readable roots only; families carry absolute font paths and skip empty dirs */
  | { kind: 'cloud'; roots: Array<{ root: string; families: Array<[string, string[]]> }> }

const WORKER_SRC = `
const { workerData } = require('node:worker_threads')
const { readdirSync, statSync } = require('node:fs')
const { join } = require('node:path')
const { port, signal, tasks } = workerData
const FONT_FILE = /\\.(ttf|otf|ttc|otc)$/i
const list = (dir) => { try { return readdirSync(dir) } catch { return null } }
const isFile = (p) => { try { return statSync(p).isFile() } catch { return false } }
const fontFiles = (dir) => (list(dir) ?? []).filter((n) => FONT_FILE.test(n))
tasks.forEach((task, i) => {
  let result
  if (task.kind === 'flat') {
    result = { kind: 'flat', files: fontFiles(task.dir).filter((n) => isFile(join(task.dir, n))) }
  } else {
    const roots = []
    for (const id of list(task.base) ?? []) {
      const root = join(task.base, id, task.sub)
      const fams = list(root)
      if (fams === null) continue
      const families = fams
        .map((f) => [f, fontFiles(join(root, f)).map((n) => join(root, f, n))])
        .filter(([, files]) => files.length)
      roots.push({ root, families })
    }
    result = { kind: 'cloud', roots }
  }
  port.postMessage({ i, result })
  Atomics.add(signal, 0, 1)
  Atomics.notify(signal, 0)
})
`

/**
 * Runs the tasks on a worker and blocks until all finish or the deadline passes.
 * Result slots of tasks that did not finish stay undefined.
 */
export function scanFontDirs(tasks: ScanTask[], deadlineMs: number): Array<ScanResult | undefined> {
  const out: Array<ScanResult | undefined> = new Array(tasks.length)
  if (!tasks.length) return out
  const { port1, port2 } = new MessageChannel()
  const signal = new Int32Array(new SharedArrayBuffer(4))
  let worker: Worker
  try {
    worker = new Worker(WORKER_SRC, {
      eval: true,
      workerData: { port: port2, signal, tasks },
      transferList: [port2],
    })
  } catch (e) {
    console.warn('[fonts] font directory scan worker unavailable:', e)
    return out
  }
  worker.on('error', (e) => console.warn('[fonts] font directory scan failed:', e))
  const deadline = Date.now() + deadlineMs
  let seen = 0
  let done = 0
  for (;;) {
    for (let m = receiveMessageOnPort(port1); m; m = receiveMessageOnPort(port1)) {
      const { i, result } = m.message as { i: number; result: ScanResult }
      out[i] = result
      done++
    }
    if (done >= tasks.length) break
    const left = deadline - Date.now()
    if (left <= 0) break
    Atomics.wait(signal, 0, seen, left)
    seen = Atomics.load(signal, 0)
  }
  port1.close()
  // A worker stuck inside readdir cannot be interrupted; it must not keep the process alive
  worker.unref()
  void worker.terminate()
  if (done < tasks.length) {
    const skipped = tasks
      .filter((_, i) => out[i] === undefined)
      .map((t) => (t.kind === 'flat' ? t.dir : t.base))
    console.warn(`[fonts] font directory scan exceeded ${deadlineMs}ms, skipped:`, skipped)
  }
  return out
}
