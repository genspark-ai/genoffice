// Store-worker test double: answers the store protocol's 'count' with a
// literal, mirrors every other type straight back, and hangs (never answers)
// while the file `<dbPath>.hang` exists — the client's timeout and respawn
// paths need both behaviors from one worker file.
import { existsSync } from 'node:fs'
import { parentPort, workerData } from 'node:worker_threads'

parentPort?.on('message', (req) => {
  if (existsSync(`${workerData.dbPath}.hang`)) return
  if (req.type === 'count') {
    parentPort?.postMessage({ id: req.id, ok: true, result: 7 })
  } else if (req.type === 'upsert') {
    parentPort?.postMessage({ id: req.id, ok: false, error: 'boom from the stub' })
  } else {
    parentPort?.postMessage({ id: req.id, ok: true, result: { type: req.type } })
  }
})
