// test double for the settle tests: a scan answers the root's files like the
// real worker; an extraction answers normally unless the marker file
// `<dir>/hold` exists next to the target — those requests wait (polling)
// until the marker is gone, which is how a test holds the extraction queue
// open and then releases it
import { existsSync, readdirSync, statSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { parentPort } from 'node:worker_threads'

const marker = (path) => join(dirname(path), 'hold')

parentPort?.on('message', (req) => {
  if (req.type === 'scan') {
    const files = readdirSync(req.root, { withFileTypes: true })
      .filter((ent) => ent.isFile())
      .map((ent) => {
        const p = join(req.root, ent.name)
        const st = statSync(p)
        return { path: p, mtimeMs: st.mtimeMs, sizeBytes: st.size }
      })
    parentPort?.postMessage({ id: req.id, type: 'scan', files })
    return
  }
  const answer = () =>
    parentPort?.postMessage({ id: req.id, type: 'extract', result: { kind: 'text', text: 'body' } })
  if (!existsSync(marker(req.path))) {
    answer()
    return
  }
  const poll = setInterval(() => {
    if (existsSync(marker(req.path))) return
    clearInterval(poll)
    answer()
  }, 10)
})
