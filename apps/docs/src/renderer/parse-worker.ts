import { parseDocx } from '@genoffice/docx-engine'
import type { ParseWorkerRequest, ParseWorkerResponse } from './parse-off-thread'

/** Worker entry: parse a .docx off the UI thread and hand the model back. */

const post = (msg: ParseWorkerResponse, transfer: Transferable[] = []) =>
  (self as unknown as Worker).postMessage(msg, transfer)

self.onmessage = async (e: MessageEvent<ParseWorkerRequest>) => {
  const { id, bytes } = e.data
  try {
    const parsed = await parseDocx(new Uint8Array(bytes))
    // HTML altChunks convert through the host's hidden window, reachable only
    // from the UI thread: hand the bytes back for the inline parse
    if (parsed.extras.altChunksNeedConverter) {
      post({ id, ok: false, error: 'altChunk needs the UI thread', bytes }, [bytes])
      return
    }
    // the model keeps the source bytes for save patching and the embedded font
    // faces: move those buffers back instead of copying them
    const buffers = new Set<ArrayBuffer>([parsed.internal.originalBytes.buffer as ArrayBuffer])
    for (const f of parsed.embeddedFonts ?? []) buffers.add(f.data.buffer as ArrayBuffer)
    post({ id, ok: true, parsed }, [...buffers])
  } catch (err) {
    // the caller retries on the UI thread with the same bytes
    post({ id, ok: false, error: String(err), bytes }, [bytes])
  }
}
