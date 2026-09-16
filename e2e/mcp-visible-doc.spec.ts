import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { existsSync } from 'node:fs'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import http, { createServer } from 'node:http'

/**
 * MCP visible-editing path, end to end against the real built app.
 *
 * Pre-seeds the MCP settings, launches the shell, then drives the local MCP
 * server over Streamable HTTP exactly as an external agent would:
 * create_session family=docx (opens a visible tab) → insert_content → apply_ops (format)
 * → save_session (writes the file). Asserts both the on-disk docx and that a
 * docs tab actually appeared in the shell UI.
 */

async function freePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const server = createServer()
    server.on('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const address = server.address()
      const port = typeof address === 'object' && address ? address.port : 0
      server.close(() => resolve(port))
    })
  })
}

interface RpcResponse {
  status: number
  headers: http.IncomingHttpHeaders
  body: unknown
}

function parseBody(text: string): unknown {
  const trimmed = text.trim()
  if (!trimmed) return null
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
    try {
      return JSON.parse(trimmed)
    } catch {
      return trimmed
    }
  }
  const out: unknown[] = []
  for (const line of trimmed.split('\n')) {
    if (!line.startsWith('data: ')) continue
    try {
      out.push(JSON.parse(line.slice(6)))
    } catch {
      /* endpoint/heartbeat line */
    }
  }
  return out.length === 1 ? out[0] : out
}

function postMcp(port: number, sessionId: string | null, payload: unknown): Promise<RpcResponse> {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(payload)
    const headers: Record<string, string | number> = {
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(data),
      Accept: 'application/json, text/event-stream',
    }
    if (sessionId) headers['Mcp-Session-Id'] = sessionId
    const req = http.request(
      { hostname: '127.0.0.1', port, path: '/mcp', method: 'POST', headers },
      (res) => {
        let text = ''
        res.on('data', (c) => (text += c))
        res.on('end', () =>
          resolve({ status: res.statusCode ?? 0, headers: res.headers, body: parseBody(text) }),
        )
      },
    )
    req.on('error', reject)
    req.write(data)
    req.end()
  })
}

async function waitForHealth(port: number, timeoutMs = 20_000): Promise<void> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    try {
      const ok = await new Promise<boolean>((resolve) => {
        const req = http.get({ hostname: '127.0.0.1', port, path: '/health' }, (res) => {
          res.resume()
          resolve(res.statusCode === 200)
        })
        req.on('error', () => resolve(false))
      })
      if (ok) return
    } catch {
      /* retry */
    }
    if (Date.now() > deadline) throw new Error('MCP server never became healthy')
    await new Promise((r) => setTimeout(r, 250))
  }
}

test.describe('MCP visible document session', () => {
  test('creates, formats and saves a docx through a visible tab', async () => {
    const port = await freePort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-e2e-'))
    const outFile = join(outDir, 'mcp-visible.docx')
    // seed an app-settings.json that both skips onboarding and turns MCP on;
    // launchShell only writes its own file when onboardingSeen is requested,
    // so seed it here and leave that option off.
    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-userdata-'))
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true, mcpEnabled: true, mcpPort: port }),
    )
    const launched = await launchShell({
      videoDir: 'mcp-visible',
      userDataDir,
    })
    const { app, page } = launched
    try {
      await waitForHealth(port)

      // handshake
      const init = await postMcp(port, null, {
        jsonrpc: '2.0',
        id: 1,
        method: 'initialize',
        params: {
          protocolVersion: '2024-11-05',
          capabilities: {},
          clientInfo: { name: 'e2e', version: '1.0' },
        },
      })
      expect(init.status).toBe(200)
      const sessionId = String(init.headers['mcp-session-id'] ?? '')
      expect(sessionId).not.toBe('')
      await postMcp(port, sessionId, { jsonrpc: '2.0', method: 'notifications/initialized' })

      const call = async (name: string, args: Record<string, unknown>) => {
        const res = await postMcp(port, sessionId, {
          jsonrpc: '2.0',
          id: Math.floor(Math.random() * 1e6),
          method: 'tools/call',
          params: { name, arguments: args },
        })
        const body = res.body as {
          result?: { isError?: boolean; content?: Array<{ text?: string }> }
        }
        const text = body?.result?.content?.map((c) => c.text ?? '').join('') ?? ''
        return { isError: body?.result?.isError === true, text }
      }

      // 1. open a visible blank document
      const created = await call('create_session', { family: 'docx' })
      expect(created.isError).toBeFalsy()

      // the shell UI shows a real docs tab (the visible half of the feature)
      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      const editorPage = await waitForPageWithUrl(app, 'genoffice-app://docs')

      // 2. write content (visible edit), 3. format it
      const inserted = await call('insert_content', {
        html: '<h1>Quarterly Report</h1><p>Revenue grew steadily.</p><ul><li>North</li><li>South</li></ul>',
      })
      expect(inserted.isError, inserted.text).toBeFalsy()

      const formatted = await call('apply_ops', {
        ops: [
          { op: 'setParagraphFormat', target: { blockIndexes: [0] }, align: 'center' },
          { op: 'setFont', target: { blockIndexes: [1] }, bold: true },
        ],
      })
      expect(formatted.isError, formatted.text).toBeFalsy()

      // the live editor really shows the content
      await expect(editorPage.locator('.ProseMirror')).toContainText('Quarterly Report')

      // 4. output to an explicit path
      const saved = await call('save_session', { path: outFile })
      expect(saved.isError, saved.text).toBeFalsy()
      expect(existsSync(outFile)).toBe(true)

      const bytes = await readFile(outFile)
      expect(bytes.byteLength).toBeGreaterThan(0)
      // docx is a zip whose entries start with PK
      expect(bytes[0]).toBe(0x50)
      expect(bytes[1]).toBe(0x4b)

      // the renderer's direct save IPC only accepts targets the MCP layer resolved
      const stray = join(outDir, 'stray.docx')
      const refused = await editorPage.evaluate(
        (target) =>
          (
            window as unknown as {
              desktop: {
                saveDocxTo: (
                  p: string,
                  d: ArrayBuffer,
                  o: boolean,
                ) => Promise<{ ok: boolean; error?: string }>
              }
            }
          ).desktop.saveDocxTo(target, new ArrayBuffer(4), true),
        stray,
      )
      expect(refused.ok).toBe(false)
      expect(refused.error).toContain('not authorized')
      expect(existsSync(stray)).toBe(false)

      // read it back through the file tool to confirm the content persisted
      const readBack = await call('read_docx', { path: outFile })
      expect(readBack.isError).toBeFalsy()
      expect(readBack.text).toContain('Quarterly Report')
      expect(readBack.text).toContain('Revenue grew steadily')
    } finally {
      await closeAndSaveVideo(launched, 'mcp-visible')
    }
  })
})
