import { randomUUID } from 'node:crypto'
import { ipcMain, type WebContents } from 'electron'
import type { ZoteroCommandResult, ZoteroRendererResponse } from '../shared/ipc'
import { ZoteroWireClient, type ZoteroWireRequest } from './zotero-wire'

const clients = new Map<number, ZoteroWireClient>()
const pending = new Map<
  string,
  { webContentsId: number; resolve: (value: unknown) => void; reject: (error: Error) => void }
>()
let registered = false

function requestRenderer(contents: WebContents, request: ZoteroWireRequest): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const requestId = randomUUID()
    const timer = setTimeout(() => {
      pending.delete(requestId)
      reject(new Error(`Timed out handling ${request.command}`))
    }, 120_000)
    pending.set(requestId, {
      webContentsId: contents.id,
      resolve: (value) => {
        clearTimeout(timer)
        resolve(value)
      },
      reject: (error) => {
        clearTimeout(timer)
        reject(error)
      },
    })
    contents.send('zotero:request', { requestId, ...request })
  })
}

export function registerZoteroIpc(): void {
  if (registered) return
  registered = true

  ipcMain.handle('zotero:command', async (event, command: string): Promise<ZoteroCommandResult> => {
    if (!/^(addEditCitation|addEditBibliography|refresh|setDocPrefs|removeCodes)$/.test(command)) {
      return { ok: false, errorCode: 'unsupported-command', error: 'Unsupported Zotero command' }
    }
    let client = clients.get(event.sender.id)
    if (!client) {
      client = new ZoteroWireClient()
      clients.set(event.sender.id, client)
    }
    try {
      await client.run(command, (request) => requestRenderer(event.sender, request))
      return { ok: true }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      return {
        ok: false,
        errorCode: /ECONNREFUSED/.test(message) ? 'connection-refused' : 'operation-failed',
        error: message,
      }
    }
  })

  ipcMain.on('zotero:response', (event, response: ZoteroRendererResponse) => {
    if (!response || typeof response.requestId !== 'string') return
    const waiter = pending.get(response.requestId)
    if (!waiter || waiter.webContentsId !== event.sender.id) return
    pending.delete(response.requestId)
    if (response.ok) waiter.resolve(response.result)
    else waiter.reject(new Error(response.error || 'Zotero document operation failed'))
  })
}

export function teardownZoteroIpc(contents: WebContents): void {
  clients.get(contents.id)?.close()
  clients.delete(contents.id)
  for (const [requestId, waiter] of pending) {
    if (waiter.webContentsId !== contents.id) continue
    pending.delete(requestId)
    waiter.reject(new Error('The document was closed'))
  }
}
