import { contextBridge, ipcRenderer } from 'electron'
import { QUARANTINE_CHANNELS } from '../shared/quarantine-channels'
import type { QuarantineApi } from '../shared/quarantine-channels'

// Quarantine bridge (H5 owns). Generic passthrough for the quarantine-only
// channels: every method forwards to its canonical QUARANTINE_CHANNELS
// entry with no logic beyond a hand-rolled shape guard. This file stays
// zod-free (compose-bridge precedent); main-process handlers re-validate
// with the zod schemas in `../shared/quarantine-schemas`.

function assertText(value: unknown, field: string, max = 4096): asserts value is string {
  if (typeof value !== 'string' || value.length === 0 || value.length > max) {
    throw new Error(`mail: bad quarantine ${field}`)
  }
}

function assertBytes(value: unknown): asserts value is Uint8Array {
  if (!(value instanceof Uint8Array) || value.byteLength === 0) {
    throw new Error('mail: bad quarantine attachment bytes')
  }
}

const quarantineApi: QuarantineApi = {
  scan: (filename, mime, data) => {
    assertText(filename, 'filename', 256)
    assertText(mime, 'mime', 128)
    assertBytes(data)
    return ipcRenderer.invoke(QUARANTINE_CHANNELS.scan, { filename, mime, data })
  },
  ingest: (filename, mime, data, opts) => {
    assertText(filename, 'filename', 256)
    assertText(mime, 'mime', 128)
    assertBytes(data)
    return ipcRenderer.invoke(QUARANTINE_CHANNELS.ingest, {
      filename,
      mime,
      data,
      ...(opts?.accountId !== undefined ? { accountId: opts.accountId } : {}),
      ...(opts?.msgUid !== undefined ? { msgUid: opts.msgUid } : {}),
    })
  },
  preview: (path) => {
    assertText(path, 'path')
    return ipcRenderer.invoke(QUARANTINE_CHANNELS.preview, { path })
  },
  openAttachment: (path, confirmed) => {
    assertText(path, 'path')
    if (typeof confirmed !== 'boolean') throw new Error('mail: bad quarantine confirm')
    return ipcRenderer.invoke(QUARANTINE_CHANNELS.open, { path, confirmed })
  },
}

contextBridge.exposeInMainWorld('mailQuarantine', quarantineApi)

export type { QuarantineApi }
