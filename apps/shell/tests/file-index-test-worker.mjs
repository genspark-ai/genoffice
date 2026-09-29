// test double for the file-index worker: walks the root like the real scan, and
// answers every extraction except the one for the poison file, which never
// answers — a wedged parse
import { parentPort } from 'node:worker_threads'
import { readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'

parentPort?.on('message', (req) => {
  if (req.type === 'scan') {
    const files = []
    const walk = (dir) => {
      let ents
      try {
        ents = readdirSync(dir, { withFileTypes: true })
      } catch {
        return
      }
      for (const ent of ents) {
        const p = join(dir, ent.name)
        if (ent.isDirectory()) walk(p)
        else if (ent.isFile()) {
          const st = statSync(p)
          files.push({ path: p, mtimeMs: st.mtimeMs, sizeBytes: st.size })
        }
      }
    }
    walk(req.root)
    parentPort.postMessage({ id: req.id, type: 'scan', files })
    return
  }
  if (req.path.includes('poison')) return
  parentPort.postMessage({
    id: req.id,
    type: 'extract',
    result: { kind: 'text', text: `body of ${req.path}` },
  })
})
