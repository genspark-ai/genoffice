import { createServer, type Server } from 'node:http'
import { randomBytes } from 'node:crypto'
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { WebSocketServer } from 'ws'
import { app } from 'electron'
import { attachWebBridge, broadcastWebEvent } from './bridge'
import { installWebDialogPatch } from './web-dialogs'
import {
  originHostOk,
  serveWebStatic,
  tokenValid,
  webRendererRoots,
  type RendererRoots,
} from './static'
import { serveWebApi } from './api'
import { createCliRunner, type CliRunner } from '../main/mcp/cli-runner'
import { showErrorDialog } from '../main/error-dialog'
import { setWebPrintToPDFDelegate } from './web-contents'
import { webPrintToPDF } from './web-print'

/**
 * Web mode: the running shell app doubles as the local service. The browser
 * front end talks to the same main-process IPC handlers the desktop windows
 * use — one WebSocket per browser tab stands in for the WebContents an Electron
 * view would own. Start it with `--web-server` (or GENOFFICE_WEB=1).
 *
 * The service sits on every captured IPC handler, so it is always
 * token-authenticated: a random token is generated unless one is given, every
 * route and the WebSocket upgrade require it, and Origin headers must match
 * the server's own Host (DNS-rebinding / cross-site guard). Share the printed
 * `?token=…` URL, never the bare port.
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

/** only `--flag value` consumes the next argv; a bare `--web-token` generates */
function flagValue(argv: string[], i: number): [value: string | undefined, next: number] {
  const value = argv[i + 1]
  if (value === undefined || value.startsWith('--')) return [undefined, i]
  return [value, i + 1]
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
    else if (arg === '--web-token') {
      // value optional: `--web-token secret` pins one, bare `--web-token`
      // generates (the token is never optional)
      const [value, consumed] = flagValue(argv, i)
      i = consumed
      overrides.token = value ?? randomBytes(24).toString('hex')
    }
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

/** the same bundled CLI the MCP tools spawn (app's Node runtime + cli entry),
 *  shared by the desktop runner in shell main so both paths stay one binary */
export function bundledWebCliRunner(): CliRunner {
  const exe = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  const sidecar = app.isPackaged
    ? join(process.resourcesPath, 'native', exe)
    : join(app.getAppPath(), '..', 'sheets', 'native', 'xlsx-engine', 'target', 'release', exe)
  return createCliRunner({
    executable: process.execPath,
    // app path is apps/shell: packages live one level above apps, the repo
    // root two levels above (same layout the shell's own runner resolves)
    entry: app.isPackaged
      ? join(process.resourcesPath, 'cli', 'genoffice.cjs')
      : join(app.getAppPath(), '..', '..', 'packages', 'cli', 'dist', 'genoffice.cjs'),
    env: existsSync(sidecar) ? { XLSX_SIDECAR_PATH: sidecar } : {},
  })
}

export async function startWebServer(
  overrides: Partial<WebServerConfig> = {},
  deps: { apiRunner?: CliRunner } = {},
): Promise<WebServerHandle> {
  const config: WebServerConfig = { ...defaultWebConfig(), ...overrides }
  // the surface reaches every file-read/write handler: a token is not optional
  config.token ??= randomBytes(24).toString('hex')
  const token = config.token
  const roots: RendererRoots = webRendererRoots()
  const apiRunner = deps.apiRunner ?? bundledWebCliRunner()

  installWebDialogPatch()
  setWebPrintToPDFDelegate(webPrintToPDF)

  // an uncaught exception in web mode must not become Electron's native error
  // dialog: its modal loop stops the Node event loop, and with it every
  // HTTP/WebSocket connection the browser holds. Log it and surface it through
  // the async (non-modal) error dialog the desktop app already uses, so the
  // failure is visible without freezing live browser sessions.
  process.on('uncaughtException', (err) => {
    console.error('[web] uncaught exception:', err)
    showErrorDialog(null, 'GenOffice web service hit an unexpected error.', err)
  })
  process.on('unhandledRejection', (reason) => {
    console.error('[web] unhandled rejection:', reason)
  })

  const bound = { host: config.host, port: config.port }
  const server = createServer((req, res) => {
    if (!originHostOk(req, bound)) {
      res.writeHead(403)
      res.end('cross-origin request refused')
      return
    }
    void serveWebApi(req, res, { token, runner: apiRunner, port: config.port })
      .then((served) => {
        if (served) return
        return serveWebStatic(req, res, { roots, token, port: config.port })
      })
      .catch((err) => {
        console.error('[web] request failed:', err)
        if (!res.headersSent) {
          res.writeHead(500)
          res.end('internal error')
        }
      })
  })

  const wss = new WebSocketServer({ noServer: true })
  server.on('upgrade', (req, socket, head) => {
    // the WebSocket is not governed by the browser's same-origin policy, so
    // the Origin check is what keeps a malicious page from driving the user's
    // session with its cookies; token check does the rest
    if (!originHostOk(req, bound) || !tokenValid(req, token, config.port)) {
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
  console.log(`[web] GenOffice web UI: ${where}?token=${token}`)
  console.log('[web] token auth is always on — share the printed URL, not the bare port')

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
