/// The CUPS print route for silent jobs on macOS: Chromium's silent
/// `webContents.print` can drop a job without a queue entry and without a
/// callback error (network drivers especially — the print system never
/// reports back), while `lp` spools a real job whose state `lpstat -o` shows
/// and whose failures come back with a non-zero exit and a message. The
/// caller renders the pages to PDF (`webContents.printToPDF`) and hands the
/// bytes here with the same resolved options the in-app dialog would have
/// given the direct print.

import { spawn } from 'node:child_process'

export interface CupsPrintJob {
  /** CUPS queue name; undefined = the system's default destination */
  deviceName?: string
  copies?: number
  collate?: boolean
  duplexMode?: 'simplex' | 'longEdge' | 'shortEdge'
}

const SIDES: Record<NonNullable<CupsPrintJob['duplexMode']>, string> = {
  simplex: 'one-sided',
  longEdge: 'two-sided-long-edge',
  shortEdge: 'two-sided-short-edge',
}

/** The lp argument list for one resolved job. Flags only: lp reads the
 * document from stdin, so no file operand is appended here. */
export function lpArgumentsFor(job: CupsPrintJob): string[] {
  const args: string[] = []
  if (job.deviceName) args.push('-d', job.deviceName)
  if (job.copies !== undefined) args.push('-n', String(job.copies))
  if (job.collate !== undefined) args.push('-o', job.collate ? 'Collate=True' : 'Collate=False')
  if (job.duplexMode !== undefined) args.push('-o', `sides=${SIDES[job.duplexMode]}`)
  return args
}

/** printToPDF options assembled from the page geometry the caller resolved.
 * Units are the printToPDF ones (inches, scale as a float); the apps convert
 * from their own units where their constants already live. Omitted pieces
 * fall to Chromium's defaults (Letter, 1cm margins) rather than being sent
 * as zeros the caller did not mean. */
export interface PrintToPdfRequest {
  widthIn?: number
  heightIn?: number
  landscape?: boolean
  marginsIn?: { top: number; bottom: number; left: number; right: number }
  scale?: number
  /** printBackground: the office documents all need it; default true */
  printBackground?: boolean
}

export function printToPdfOptionsFor(request: PrintToPdfRequest): {
  printBackground: boolean
  margins?: { top: number; bottom: number; left: number; right: number }
  pageSize?: { width: number; height: number }
  landscape?: boolean
  scale?: number
} {
  return {
    printBackground: request.printBackground ?? true,
    ...(request.marginsIn ? { margins: request.marginsIn } : {}),
    ...(request.widthIn !== undefined && request.heightIn !== undefined
      ? { pageSize: { width: request.widthIn, height: request.heightIn } }
      : {}),
    ...(request.landscape ? { landscape: true } : {}),
    ...(request.scale !== undefined && request.scale !== 1 ? { scale: request.scale } : {}),
  }
}

/** Runs `lp` with the PDF on stdin and settles with the queue's verdict: a
 * zero exit reports the accepted job; anything else surfaces lp's own
 * message (a mangled or missing queue name fails here instead of
 * vanishing). `spawnFn` is injectable so tests can fake the process. */
export async function printPdfViaCups(
  pdf: Uint8Array,
  job: CupsPrintJob,
  spawnFn: typeof spawn = spawn,
): Promise<{ ok: boolean; error?: string }> {
  return new Promise((resolve) => {
    const child = spawnFn('/usr/bin/lp', lpArgumentsFor(job), {
      stdio: ['pipe', 'pipe', 'pipe'],
    })
    let stderr = ''
    let stdout = ''
    let settled = false
    const settle = (result: { ok: boolean; error?: string }) => {
      if (!settled) {
        settled = true
        resolve(result)
      }
    }
    child.on('error', (err) => settle({ ok: false, error: err.message }))
    child.stdout.on('data', (chunk) => {
      stdout += chunk
    })
    child.stderr.on('data', (chunk) => {
      stderr += chunk
    })
    child.on('close', (code) => {
      if (code === 0) settle({ ok: true })
      else
        settle({
          ok: false,
          error: stderr.trim() || stdout.trim() || `lp exited with code ${code}`,
        })
    })
    // lp exits early on a bad queue and closes stdin: an EPIPE here must not
    // crash the main process — the close handler carries the verdict
    child.stdin.on('error', () => undefined)
    child.stdin.end(Buffer.from(pdf))
  })
}
