/**
 * A `ReadableStream` over bytes already in memory, fed one slice at a time.
 *
 * `new Blob([bytes]).stream()` is one line shorter, and both decompression sites
 * in this package used it, but `Blob.prototype.stream` is not everywhere: jsdom
 * — the environment the Docs, Slides, PDF, Markdown and HTML app tests run in —
 * ships a `Blob` without it (its `ReadableStream` is Node's, so streams
 * themselves are fine). #798 is the metafile side of that, #781 the zip-gate
 * side, where it failed 389 `apps/docs` tests.
 *
 * Deliberately free of Node builtins: both callers reach the renderer bundle,
 * where `node:*` imports fail the build on Vite's browser-external stubs.
 */
export function streamBytes(bytes: Uint8Array): ReadableStream<Uint8Array> {
  const SLICE = 64 * 1024
  let offset = 0
  return new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset >= bytes.length) {
        controller.close()
        return
      }
      const end = Math.min(offset + SLICE, bytes.length)
      // slice() copies, so each chunk is a standalone Uint8Array whatever the
      // caller's view is backed by, and the decompression stays metered: input
      // arrives in bounded pieces and can be cancelled before it runs away.
      controller.enqueue(bytes.slice(offset, end))
      offset = end
    },
  })
}

/**
 * `DecompressionStream` as a `Uint8Array` → `Uint8Array` transform.
 *
 * The cast is for builds that add the DOM lib (the cli does): there
 * `DecompressionStream.writable` is `WritableStream<BufferSource>`, which is not
 * assignable to the `WritableStream<Uint8Array>` `pipeThrough` asks for. Same
 * object, same runtime contract; only the two libs' generics disagree.
 */
export function decompressionStream(format: 'deflate' | 'deflate-raw' | 'gzip'): {
  readable: ReadableStream<Uint8Array>
  writable: WritableStream<Uint8Array>
} {
  return new DecompressionStream(format) as unknown as {
    readable: ReadableStream<Uint8Array>
    writable: WritableStream<Uint8Array>
  }
}
