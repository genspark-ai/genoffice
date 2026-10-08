// test double: every scan answers with no files and a truncation flag, as a walk
// that hit its budget before reaching anything would
import { parentPort } from 'node:worker_threads'

parentPort?.on('message', (req) => {
  if (req.type === 'scan') {
    parentPort.postMessage({ id: req.id, type: 'scan', files: [], truncated: true })
    return
  }
  parentPort.postMessage({ id: req.id, type: 'extract', result: { kind: 'name-only' } })
})
