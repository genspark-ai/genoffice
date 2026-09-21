import { parseDocx } from '@genoffice/docx-engine'

type Parsed = Awaited<ReturnType<typeof parseDocx>>

export interface ParseWorkerRequest {
  id: number
  bytes: ArrayBuffer
}

export type ParseWorkerResponse =
  | { id: number; ok: true; parsed: Parsed }
  | { id: number; ok: false; error: string; bytes: ArrayBuffer }

/**
 * Parsing a large .docx (unzip, XML, block model) took the UI thread for
 * seconds to tens of seconds with the "Opening…" screen frozen. The parse
 * runs in a module Worker; the bytes are transferred there and the model
 * (structured-cloneable: plain objects, Maps, typed arrays) comes back with
 * its source bytes and font buffers transferred, not copied. Without Worker
 * support, or when the worker fails, the same parse runs inline.
 */

let worker: Worker | null = null
let nextId = 0
const pending = new Map<number, (r: ParseWorkerResponse) => void>()

function getWorker(): Worker | null {
  if (worker) return worker
  if (typeof Worker === 'undefined') return null
  try {
    worker = new Worker(new URL('./parse-worker.ts', import.meta.url), { type: 'module' })
  } catch {
    return null
  }
  worker.onmessage = (e: MessageEvent<ParseWorkerResponse>) => {
    const resolve = pending.get(e.data.id)
    if (!resolve) return
    pending.delete(e.data.id)
    resolve(e.data)
  }
  // a crashed worker (out of memory on a huge file) fails every waiting parse;
  // they fall back inline and the next parse starts a fresh worker
  worker.onerror = () => {
    const waiting = [...pending.entries()]
    pending.clear()
    worker?.terminate()
    worker = null
    for (const [id, resolve] of waiting) {
      resolve({ id, ok: false, error: 'parse worker failed', bytes: new ArrayBuffer(0) })
    }
  }
  return worker
}

/** how parses ran, for probes and troubleshooting */
export const parseStats = { worker: 0, inline: 0 }

export interface ParseOffThreadOptions {
  /** the caller hands the bytes over: their buffer is transferred to the worker
   *  and the view detaches. Default copies, since a save reparse may pass the
   *  live document's own source bytes (an unchanged save returns them as is) */
  owned?: boolean
}

export async function parseDocxOffThread(
  bytes: Uint8Array,
  options: ParseOffThreadOptions = {},
): Promise<Parsed> {
  const w = getWorker()
  if (!w) {
    parseStats.inline++
    return parseDocx(bytes)
  }
  const whole = bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength
  const owned =
    options.owned && whole ? (bytes.buffer as ArrayBuffer) : (bytes.slice().buffer as ArrayBuffer)
  const id = ++nextId
  const response = await new Promise<ParseWorkerResponse>((resolve) => {
    pending.set(id, resolve)
    w.postMessage({ id, bytes: owned } satisfies ParseWorkerRequest, [owned])
  })
  if (response.ok) {
    parseStats.worker++
    return response.parsed
  }
  // a parse error surfaces the same way inline (the caller reports it);
  // a lost buffer (worker crash) is the one case with nothing to retry on
  if (response.bytes.byteLength === 0) throw new Error(response.error)
  parseStats.inline++
  return parseDocx(new Uint8Array(response.bytes))
}
if (typeof window !== 'undefined') {
  ;(window as unknown as Record<string, unknown>).__docsParseStats = parseStats
}
