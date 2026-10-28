import { createServer, type Server } from 'node:http'
import { randomBytes } from 'node:crypto'
import { WebSocketServer } from 'ws'
import { attachWebBridge, broadcastWebEvent } from './bridge'
import { installWebDialogPatch } from './web-dialogs'
import { serveWebStatic, tokenValid, webRendererRoots, type RendererRoots } from './static'
import { setWebPrintToPDFDelegate } from './web-contents'
import { webPrintToPDF } from './web-print'

/**
 * Web mode: the running shell app doubles as the local service. The browser
 * front end talks to the same main-process IPC handlers the desktop windows
 * use — one WebSocket per browser tab stands in for the WebContents an Electron
 * view would own. Start it with `--web-server` (or GENOFFICE_WEB=1).
 */

export interface WebServerConfig {
  port: number
  host: string
  token: string | null
}

export interface WebServerHandle {
  server: Server
  config: WebServerConfig
  close(): void
}

export function parseWebServerArgv(argv: string[]): {
  enabled: boolean
  overrides: Partial<WebServerConfig>
} {
  let enabled = process.env.GENOFFICE_WEB === '1'
  const overrides: Partial<WebServerConfig> = {}
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if (arg === '--web-server') enabled = true
    else if (arg === '--web-port') overrides.port = Number(argv[++i])
    else if (arg === '--web-host') overrides.host = argv[++i]
    else if (arg === '--web-token') overrides.token = argv[++i] || randomBytes(24).toString('hex')
  }
  if (process.env.GENOFFICE_WEB_PORT) overrides.port = Number(process.env.GENOFFICE_WEB_PORT)
  if (process.env.GENOFFICE_WEB_HOST) overrides.host = process.env.GENOFFICE_WEB_HOST
  if (process.env.GENOFFICE_WEB_TOKEN !== undefined) {
    overrides.token = process.env.GENOFFICE_WEB_TOKEN || randomBytes(24).toString('hex')
  }
  return { enabled, overrides }
}

export function defaultWebConfig(): WebServerConfig {
  return { port: 8931, host: '127.0.0.1', token: null }
}

export async function startWebServer(
  overrides: Partial<WebServerConfig> = {},
): Promise<WebServerHandle> {
  const config: WebServerConfig = { ...defaultWebConfig(), ...overrides }
  const roots: RendererRoots = webRendererRoots()

  installWebDialogPatch()
  setWebPrintToPDFDelegate(webPrintToPDF)

  // an uncaught exception in web mode must not become Electron's native error
  // dialog: its modal loop stops the Node event loop, and with it every
  // HTTP/WebSocket connection the browser holds. Log loudly instead.
  process.on('uncaughtException', (err) => {
    console.error('[web] uncaught exception:', err)
  })
  process.on('unhandledRejection', (reason) => {
    console.error('[web] unhandled rejection:', reason)
  })

  const server = createServer((req, res) => {
    void serveWebStatic(req, res, { roots, token: config.token }).catch((err) => {
      console.error('[web] request failed:', err)
      if (!res.headersSent) {
        res.writeHead(500)
        res.end('internal error')
      }
    })
  })

  const wss = new WebSocketServer({ noServer: true })
  server.on('upgrade', (req, socket, head) => {
    if (config.token && !tokenValid(req, config.token)) {
      socket.write('HTTP/1.1 401 Unauthorized\r\n\r\n')
      socket.destroy()
      return
    }
    wss.handleUpgrade(req, socket, head, (ws) => {
      wss.emit('connection', ws, req)
    })
  })
  attachWebBridge(wss, {
    tabUrl: (tab) => tab.url,
  })
  bridgeSocketServer = wss

  await new Promise<void>((resolvePromise, reject) => {
    server.once('error', reject)
    server.listen(config.port, config.host, () => {
      server.removeListener('error', reject)
      resolvePromise()
    })
  })

  const where = `http://${config.host === '0.0.0.0' || config.host === '::' ? '127.0.0.1' : config.host}:${config.port}/`
  const suffix = config.token ? `?token=${config.token}` : ''
  console.log(`[web] GenOffice web UI: ${where}${suffix}`)
  if (config.token) console.log('[web] token auth is on — share the printed URL, not the bare port')

  return {
    server,
    config,
    close: () => {
      wss.close()
      server.close()
    },
  }
}

/** settings/language style broadcasts reach browser tabs too */
export function webBroadcast(channel: string, args: unknown[]): void {
  if (!bridgeSocketServer) return
  broadcastWebEvent(bridgeSocketServer, channel, args)
}

let bridgeSocketServer: WebSocketServer | null = null
