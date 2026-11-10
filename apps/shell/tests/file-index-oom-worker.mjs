// test double: the scan reports the root's files; extracting any file named
// hog allocates until the worker's heap cap, other files extract normally
import { parentPort } from 'node:worker_threads'
import { readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'

const hoard = []
parentPort?.on('message', (req) => {
  if (req.type === 'scan') {
    const files = readdirSync(req.root, { withFileTypes: true })
      .filter((ent) => ent.isFile())
      .map((ent) => {
        const p = join(req.root, ent.name)
        const st = statSync(p)
        return { path: p, mtimeMs: st.mtimeMs, sizeBytes: st.size }
      })
    parentPort.postMessage({ id: req.id, type: 'scan', files })
    return
  }
  if (req.path.includes('hog')) {
    for (;;) hoard.push(new Array(1 << 18).fill(req.id))
  }
  parentPort.postMessage({ id: req.id, type: 'extract', result: { kind: 'text', text: 'body' } })
})
