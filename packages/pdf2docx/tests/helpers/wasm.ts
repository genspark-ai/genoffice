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
    // pin thisProgram like the app and CLI bootstraps do: the emscripten
    // default is the running program's path, and a checkout under a
    // non-ASCII directory crashes _environ_get at init (stringToAscii assert)
    const wrapped = (await init({ wasmBinary, thisProgram: 'genoffice' })) as { pdfium?: unknown }
    const m = (wrapped.pdfium ?? wrapped) as PdfiumModule & { _PDFiumExt_Init(): void }
    m._PDFiumExt_Init()
    return m
  })()
  return cached
}
