/** Test-only wasm bootstrap. The package itself never reads files; tests may. */
import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import type { PdfiumModule } from '../../src'

let cached: Promise<PdfiumModule> | null = null

/** Load + initialize the pdfium wasm from node_modules (singleton per worker). */
export function loadPdfium(): Promise<PdfiumModule> {
  cached ??= (async () => {
    const require = createRequire(import.meta.url)
    const wasmPath = require.resolve('@embedpdf/pdfium/pdfium.wasm')
    const raw = readFileSync(wasmPath)
    // exact slice: Buffer.buffer may be a shared pool larger than the file
    const wasmBinary = raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength)
    const { init } = (await import('@embedpdf/pdfium')) as unknown as {
      init(overrides: object): Promise<object>
    }
    const wrapped = (await init({
      wasmBinary,
      // A checkout under a non-ASCII path seeds the wasm environment with
      // `_: argv[1]` — the vitest worker's full path — and the wasm's musl
      // getenv later copies every string into the heap through stringToAscii,
      // which asserts on any character above 0xff. Pin the program name and
      // drop any environment value carrying such a character (PWD and friends
      // under a localized checkout), so the loader never trips that assert.
      thisProgram: '/usr/bin/node',
      ENV: Object.fromEntries(
        Object.entries(process.env).filter(
          (entry): entry is [string, string] =>
            entry[1] !== undefined && !/[\u{100}-\u{10FFFF}]/u.test(entry[1]),
        ),
      ),
    })) as { pdfium?: unknown }
    const m = (wrapped.pdfium ?? wrapped) as PdfiumModule & { _PDFiumExt_Init(): void }
    m._PDFiumExt_Init()
    return m
  })()
  return cached
}
