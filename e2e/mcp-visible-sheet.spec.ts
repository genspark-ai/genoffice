import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { existsSync } from 'node:fs'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import http, { createServer } from 'node:http'
import JSZip from 'jszip'

/**
 * MCP visible grid session, end to end against the real built app.
 *
 * create_session family=xlsx (opens a visible sheets tab) → apply_sheet_ops (values +
 * formula, seen live by the user) → read_sheet → save_session (writes the file
 * through the regular save pipeline with an explicit path). Asserts the
 * on-disk xlsx (values AND the formula) and that a sheets tab actually
 * appeared in the shell UI.
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

test.describe('MCP visible grid session', () => {
  test('creates, fills and saves an xlsx through a visible sheets tab', async () => {
    const port = await freePort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-sheet-'))
    const outFile = join(outDir, 'mcp-sheet.xlsx')
    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-sheet-userdata-'))
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true, mcpEnabled: true, mcpPort: port }),
    )
    const launched = await launchShell({
      videoDir: 'mcp-visible-sheet',
      userDataDir,
    })
    const { app, page } = launched
    try {
      await waitForHealth(port)

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

      // 1. open a visible blank spreadsheet (session tools need no background flag)
      const created = await call('create_session', { family: 'xlsx' })
      expect(created.isError, created.text).toBeFalsy()

      // the shell UI shows a real sheets tab (the visible half of the feature)
      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      const editorPage = await waitForPageWithUrl(app, 'genoffice-app://sheets')

      // 2. workbook overview first: ops address sheets by id
      const overview = await call('read_sheet', {})
      expect(overview.isError, overview.text).toBeFalsy()
      const sheetId = /"sheetId":\s*"([^"]+)"/.exec(overview.text)?.[1] ?? ''
      expect(sheetId, overview.text).not.toBe('')

      // 3. fill values + one formula in one batch
      const applied = await call('apply_sheet_ops', {
        ops: [
          { op: 'set_cell', sheetId, address: 'A1', value: 'Item' },
          { op: 'set_cell', sheetId, address: 'A2', value: 'Widgets' },
          { op: 'set_cell', sheetId, address: 'B1', value: 'Qty' },
          { op: 'set_cell', sheetId, address: 'B2', value: 12 },
          { op: 'set_formula', sheetId, address: 'B3', formula: '=SUM(B2:B2)' },
        ],
      })
      expect(applied.isError, applied.text).toBeFalsy()
      expect(applied.text).toContain('"ok": true')

      // 4. read the grid back: values and the formula come through
      const read = await call('read_sheet', { addresses: ['A1', 'A2', 'B2', 'B3'] })
      expect(read.isError, read.text).toBeFalsy()
      expect(read.text).toContain('Widgets')
      expect(read.text).toContain('SUM(B2:B2)')

      // the visible tab really renders the grid (canvas mounted)
      await expect(editorPage.locator('canvas').first()).toBeVisible({ timeout: 15_000 })

      // 4. output to an explicit path (ends the session)
      const saved = await call('save_session', { path: outFile })
      expect(saved.isError, saved.text).toBeFalsy()
      expect(existsSync(outFile)).toBe(true)

      // a second edit after the save reports the closed session
      const afterSave = await call('read_sheet', { addresses: ['A1'] })
      expect(afterSave.isError).toBe(true)

      // unpack the saved xlsx: values AND the formula must be on disk
      const zip = await JSZip.loadAsync(await readFile(outFile))
      const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(sheetXml).toContain('Widgets')
      expect(sheetXml).toContain('<f>SUM(B2:B2)</f>')
    } finally {
      await closeAndSaveVideo(launched, 'mcp-visible-sheet')
    }
  })
})
