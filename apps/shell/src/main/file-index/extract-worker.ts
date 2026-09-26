/**
 * Worker-thread entry: text extraction and directory walks run here so a slow
 * PDF or a stalled network folder never blocks the main process.
 */
import { parentPort } from 'node:worker_threads'
import { extractText, type ExtractOcr } from './extract'
import { scanFiles } from './scan'

export type WorkerRequest =
  | { id: number; type: 'extract'; path: string; ocr?: boolean }
  | { id: number; type: 'scan'; root: string }

export type WorkerResponse =
  | { id: number; type: 'extract'; result: Awaited<ReturnType<typeof extractText>> }
  | { id: number; type: 'scan'; files: ReturnType<typeof scanFiles> }

parentPort?.on('message', async (req: WorkerRequest) => {
  if (req.type === 'extract') {
    // The OCR module (and with it the pdfium wasm) loads only when a scanned
    // PDF actually needs it; plain extractions never pay for the import.
    let ocr: ExtractOcr | undefined
    if (req.ocr) {
      const { ocrPdfToText } = await import('./ocr')
      ocr = { ocrPdf: ocrPdfToText }
    }
    const result = await extractText(req.path, ocr)
    parentPort?.postMessage({ id: req.id, type: 'extract', result } satisfies WorkerResponse)
  } else {
    parentPort?.postMessage({
      id: req.id,
      type: 'scan',
      files: scanFiles(req.root),
    } satisfies WorkerResponse)
  }
})
