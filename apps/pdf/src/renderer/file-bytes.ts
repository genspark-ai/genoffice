/** Reads a granted file through the chunked IPC route: one invoke per slice
 * instead of a single multi-megabyte ArrayBuffer over IPC — a 42 MB
 * structured-clone deserialization has killed renderers before first paint
 * on Windows (genoffice#2366), while small slices arrive fine. The main-side
 * handler reports the file size with every slice, so the loop needs no
 * separate stat round trip. */
export interface PdfFileChunkApi {
  readFileChunk(
    path: string,
    start: number,
    end: number,
  ): Promise<{ data: Uint8Array; fileSize: number }>
}

export const PDF_READ_CHUNK_BYTES = 8 * 1024 * 1024

export async function readFileBytes(
  api: PdfFileChunkApi,
  path: string,
  chunkBytes = PDF_READ_CHUNK_BYTES,
): Promise<Uint8Array> {
  const parts: Uint8Array[] = []
  let offset = 0
  let fileSize = Infinity
  do {
    const end = Math.min(offset + chunkBytes, fileSize)
    const { data, fileSize: size } = await api.readFileChunk(path, offset, end)
    if (data.length === 0) break
    parts.push(data)
    offset += data.length
    fileSize = size
  } while (offset < fileSize)
  const out = new Uint8Array(offset)
  let at = 0
  for (const part of parts) {
    out.set(part, at)
    at += part.length
  }
  return out
}
