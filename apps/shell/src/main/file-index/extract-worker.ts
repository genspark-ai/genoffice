/**
 * Worker-thread entry: text extraction and directory walks run here so a slow
 * PDF or a stalled network folder never blocks the main process.
 */
import { parentPort } from 'node:worker_threads'
import { extractText } from './extract'
import { scanFiles, type ScannedFile } from './scan'

export type WorkerRequest =
  { id: number; type: 'extract'; path: string } | { id: number; type: 'scan'; root: string }

export type WorkerResponse =
  | { id: number; type: 'extract'; result: Awaited<ReturnType<typeof extractText>> }
  | { id: number; type: 'scan'; files: ScannedFile[]; truncated?: boolean }

parentPort?.on('message', async (req: WorkerRequest) => {
  if (req.type === 'extract') {
    const result = await extractText(req.path)
    parentPort?.postMessage({ id: req.id, type: 'extract', result } satisfies WorkerResponse)
  } else {
    const { files, truncated } = scanFiles(req.root)
    parentPort?.postMessage({ id: req.id, type: 'scan', files, truncated } satisfies WorkerResponse)
  }
})
