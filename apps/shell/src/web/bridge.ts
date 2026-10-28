import { WebSocketServer, type WebSocket } from 'ws'
import type { IncomingMessage } from 'node:http'
import { markDocsNewBlank, queueDocsPendingOpen } from '../../../docs/src/main/docs-main'
import { createSheetsWebTab, setSheetsNewBlank } from '../../../sheets/src/main/sheets-main'
import { trackWebSlidesTab } from '../../../slides/src/main/slides-main'
import { trackWebTab as trackMarkdownTab } from '../../../markdown/src/main/markdown-main'
import { trackWebTab as trackPdfTab } from '../../../pdf/src/main/pdf-main'
import { trackWebTab as trackHtmlTab } from '../../../html/src/main/html-main'
import { decodeFrame, encodeFrame } from './serialize'
import { rewriteIn, rewriteOut } from './schemes'
import { WebTabContents, registerWebTab, unregisterWebTab } from './web-contents'
import {
  cancelWebDialogsFor,
  listWebDialogDir,
  resolveWebDialog,
  runWithWebDispatch,
  validateWebDialogPath,
  webDialogStartDirs,
} from './web-dialogs'
import { getInvokeHandler, getSendListeners } from './capture'

/**
 * The browser-side IPC transport. One WebSocket per browser tab; the tab is a
 * WebTabContents whose id keys the same per-tab state maps the desktop views
 * use. Frames are JSON text (encodeFrame/decodeFrame carry binaries as base64):
 *
 *   client → server: invoke (request/reply), send (fire-and-forget),
 *                    dialog-response / dialog-list (the web dialog overlay)
 *   server → client: event (main→renderer pushes), dialog (a native dialog
 *                    routed into the tab), dialog-list-reply, open-url
 */

export interface WebBridge {
  /** the renderer URL of one connected tab (for open-url composition) */
  tabUrl(tab: WebTabContents): string
  broadcast(channel: string, ...args: unknown[]): void
}

interface DispatchContext {
  tab: WebTabContents
  ws: WebSocket
}

type InvokeOverride = (ctx: DispatchContext, args: unknown[]) => Promise<unknown> | unknown

/** overrides for channels whose desktop semantics (shell tab strip, native
 *  windows) have no browser equivalent */
const invokeOverrides: Record<string, InvokeOverride> = {
  'win:new': (ctx, args) => {
    const app = appOfTab(ctx.tab)
    const openPath = typeof args[0] === 'string' && args[0] !== '' ? args[0] : null
    const suffix = openPath ? `?path=${encodeURIComponent(openPath)}` : '?new=1'
    ctx.ws.send(encodeFrame({ t: 'open-url', url: `/${app}/${suffix}` }))
    return undefined
  },
  'win:list': () => [],
  'win:focus': () => undefined,
}

function appOfTab(tab: WebTabContents): string {
  const m = /^\/([a-z]+)\//.exec(new URL(tab.url, 'http://x').pathname)
  return m?.[1] ?? 'docs'
}

/** parse /<app>/?path=…|?new=1 and run the app's web-tab registration. Every
 *  app registers its IPC here (the desktop equivalent happens inside the first
 *  createXxxView, which web mode never calls); the ?path=/?new=1 seeding then
 *  matches what view creation would have queued. */
function seedTabFromUrl(tab: WebTabContents): void {
  let url: URL
  try {
    url = new URL(tab.url, 'http://localhost')
  } catch {
    return
  }
  const app = url.pathname.split('/')[1]
  const path = url.searchParams.get('path') || undefined
  const isNew = url.searchParams.get('new') === '1'
  if (app === 'docs') {
    if (path) queueDocsPendingOpen(tab.id, path)
    if (isNew) markDocsNewBlank(tab.id)
  } else if (app === 'sheets') {
    createSheetsWebTab(tab.asWebContents(), path)
    if (isNew) setSheetsNewBlank()
  } else if (app === 'slides') {
    trackWebSlidesTab(tab.asWebContents(), path)
  } else if (app === 'markdown') {
    trackMarkdownTab(tab.asWebContents(), path)
  } else if (app === 'pdf') {
    trackPdfTab(tab.asWebContents(), path)
  } else if (app === 'html') {
    trackHtmlTab(tab.asWebContents(), path)
  }
}

function makeEvent(tab: WebTabContents): { sender: WebTabContents; senderId: number } {
  return { sender: tab, senderId: tab.id }
}

async function dispatchInvoke(
  ctx: DispatchContext,
  id: number,
  channel: string,
  rawArgs: unknown[],
): Promise<void> {
  let value: unknown
  const startedAt = Date.now()
  const args = rewriteIn(appOfTab(ctx.tab), rawArgs) as unknown[]
  try {
    // AsyncLocalStorage scopes the dialog patch to this dispatch's own await
    // chain: a concurrent desktop IPC dispatch keeps its native dialogs
    value = await runWithWebDispatch(ctx.tab, async () => {
      const override = invokeOverrides[channel]
      if (override) return override(ctx, args)
      const handler = getInvokeHandler(channel)
      if (!handler) throw new Error(`No handler for "${channel}"`)
      return handler(makeEvent(ctx.tab), ...args)
    })
  } catch (err) {
    ctx.ws.send(encodeFrame({ t: 'reply', id, ok: false, error: String(err) }))
    return
  } finally {
    const took = Date.now() - startedAt
    if (took > 300) console.log(`[web] invoke ${channel} took ${took}ms`)
  }
  ctx.ws.send(
    encodeFrame({ t: 'reply', id, ok: true, value: rewriteOut(appOfTab(ctx.tab), value) }),
  )
}

function dispatchSend(ctx: DispatchContext, channel: string, rawArgs: unknown[]): void {
  const listeners = getSendListeners(channel)
  const args = rewriteIn(appOfTab(ctx.tab), rawArgs) as unknown[]
  runWithWebDispatch(ctx.tab, () => {
    for (const listener of listeners) {
      try {
        listener(makeEvent(ctx.tab), ...args)
      } catch (err) {
        console.error(`[web] send handler ${channel}:`, err)
      }
    }
  })
}

async function dispatchDialogList(
  ctx: DispatchContext,
  id: number,
  dir: unknown,
  exts: unknown,
): Promise<void> {
  const entries =
    typeof dir === 'string'
      ? await listWebDialogDir(dir, Array.isArray(exts) ? (exts as string[]) : [])
      : []
  ctx.ws.send(encodeFrame({ t: 'dialog-list-reply', id, entries }))
}

export function attachWebBridge(
  wss: WebSocketServer,
  options: { tabUrl: (tab: WebTabContents) => string },
): void {
  wss.on('connection', (ws: WebSocket, req: IncomingMessage) => {
    const tab = new WebTabContents(req.url ?? '/')
    const app = appOfTab(tab)
    tab.deliver = (channel, args) => {
      if (ws.readyState === ws.OPEN) {
        ws.send(encodeFrame({ t: 'event', channel, args: rewriteOut(app, args) }))
      }
    }
    registerWebTab(tab)
    seedTabFromUrl(tab)

    ws.on('message', (data: Buffer) => {
      let frame: Record<string, unknown>
      try {
        frame = decodeFrame(data.toString('utf8'))
      } catch {
        return
      }
      const t = frame.t
      if (t === 'invoke' && typeof frame.channel === 'string') {
        void dispatchInvoke(
          { tab, ws },
          Number(frame.id),
          frame.channel,
          Array.isArray(frame.args) ? frame.args : [],
        )
      } else if (t === 'send' && typeof frame.channel === 'string') {
        dispatchSend({ tab, ws }, frame.channel, Array.isArray(frame.args) ? frame.args : [])
      } else if (t === 'dialog-response' && Number.isFinite(Number(frame.id))) {
        resolveWebDialog(Number(frame.id), frame.result as never)
      } else if (t === 'dialog-list' && Number.isFinite(Number(frame.id))) {
        void dispatchDialogList({ tab, ws }, Number(frame.id), frame.dir, frame.exts)
      } else if (t === 'dialog-start-dirs' && Number.isFinite(Number(frame.id))) {
        ws.send(
          encodeFrame({
            t: 'dialog-start-dirs-reply',
            id: frame.id,
            dirs: webDialogStartDirs(
              typeof frame.defaultPath === 'string' ? frame.defaultPath : undefined,
            ),
          }),
        )
      } else if (t === 'dialog-validate' && Number.isFinite(Number(frame.id))) {
        const kind = frame.kind === 'save' ? 'save' : 'open'
        const p = validateWebDialogPath(String(frame.path ?? ''), kind)
        ws.send(encodeFrame({ t: 'dialog-validate-reply', id: frame.id, path: p }))
      }
    })

    const close = () => {
      cancelWebDialogsFor(tab)
      unregisterWebTab(tab)
      tab.destroy()
    }
    ws.on('close', close)
    ws.on('error', close)
  })
  void options
}

/** push to every connected browser tab (settings/language style broadcasts) */
export function broadcastWebEvent(wss: WebSocketServer, channel: string, args: unknown[]): void {
  for (const client of wss.clients) {
    if (client.readyState === client.OPEN) client.send(encodeFrame({ t: 'event', channel, args }))
  }
}
