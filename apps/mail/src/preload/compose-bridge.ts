import { contextBridge, ipcRenderer } from 'electron'
import { COMPOSE_CHANNELS } from '../shared/compose-channels'
import type { ComposeApi, OutboxEntry } from '../shared/compose-channels'
import type { StagedAttachment } from '@genoffice/mail-protocols'

// Compose-session bridge (Lane C owns). Generic passthrough for the
// compose-only channels: every method forwards to its canonical
// COMPOSE_CHANNELS entry with no logic beyond a hand-rolled shape guard.
// This file stays zod-free (sheets precedent); main-process handlers
// re-validate with the zod schemas in `../shared/compose-schemas`.
// Shell-channel traffic stays in `mail-bridge.ts` (Lane S owns).

function assertText(value: unknown, field: string): asserts value is string {
  if (typeof value !== 'string' || value.length === 0 || value.length > 4096) {
    throw new Error(`mail: bad compose ${field}`)
  }
}

function assertBytes(value: unknown): asserts value is Uint8Array {
  if (!(value instanceof Uint8Array) || value.byteLength === 0) {
    throw new Error('mail: bad compose attachment bytes')
  }
}

const composeApi: ComposeApi = {
  stageAttachment: (filename, mime, data) => {
    assertText(filename, 'filename')
    assertText(mime, 'mime')
    assertBytes(data)
    return ipcRenderer.invoke(COMPOSE_CHANNELS.stageAttachment, { filename, mime, data })
  },
  outboxList: (accountId) => {
    assertText(accountId, 'accountId')
    return ipcRenderer.invoke(COMPOSE_CHANNELS.outboxList, { accountId })
  },
  openExternal: (url) => {
    assertText(url, 'url')
    return ipcRenderer.invoke(COMPOSE_CHANNELS.openExternal, { url })
  },
}

contextBridge.exposeInMainWorld('mailCompose', composeApi)

export type { ComposeApi }
export type { OutboxEntry, StagedAttachment }
