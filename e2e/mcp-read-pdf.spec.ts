import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo } from './helpers'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import http, { createServer } from 'node:http'

/**
 * MCP read_pdf against the real built app: proves the bundled main process
 * loads the pdfium wasm and extracts text over the actual HTTP transport.
 * Session-free by design — no tabs open, nothing to watch, just a read.
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

function postMcp(
  port: number,
  sessionId: string | null,
  payload: unknown,
): Promise<{ headers: http.IncomingHttpHeaders; body: string }> {
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
        res.on('end', () => resolve({ headers: res.headers, body: text }))
      },
    )
    req.on('error', reject)
    req.write(data)
    req.end()
  })
}

test.describe('MCP read_pdf', () => {
  test('extracts pdf text through the real app server', async () => {
    test.setTimeout(60_000)
    const port = await freePort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-pdf-'))
    const pdfPath = join(outDir, 'handout.pdf')
    const doc = await PDFDocument.create()
    doc.setTitle('E2E Handout')
    const font = await doc.embedFont(StandardFonts.Helvetica)
    const p1 = doc.addPage([400, 300])
    p1.drawText('Pdfium reads this', { x: 40, y: 200, size: 14, font })
    doc.addPage([400, 300]) // no text layer
    await writeFile(pdfPath, await doc.save())

    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-pdf-userdata-'))
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true, mcpEnabled: true, mcpPort: port }),
    )
    const launched = await launchShell({ videoDir: 'mcp-read-pdf', userDataDir })
    try {
      // health → initialize → tools/call, the minimal real-client handshake
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
      expect(init.body).toContain('"result"')
      const sessionId = String(init.headers['mcp-session-id'] ?? '')
      expect(sessionId).not.toBe('')
      await postMcp(port, sessionId, { jsonrpc: '2.0', method: 'notifications/initialized' })

      const call = async (args: Record<string, unknown>) =>
        postMcp(port, sessionId, {
          jsonrpc: '2.0',
          id: Math.floor(Math.random() * 1e6),
          method: 'tools/call',
          params: { name: 'read_pdf', arguments: args },
        })

      /** pull the tool's JSON payload out of the SSE-framed tools/call reply */
      const payloadOf = (body: string): Record<string, unknown> => {
        const dataLine = body.split('\n').find((l) => l.startsWith('data: ')) ?? ''
        const reply = JSON.parse(dataLine.slice(6)) as {
          result?: { content?: Array<{ text?: string }> }
        }
        return JSON.parse(reply.result?.content?.[0]?.text ?? '{}') as Record<string, unknown>
      }

      const whole = payloadOf((await call({ path: pdfPath })).body) as {
        pageCount: number
        info: { title?: string }
        pages: Array<{ page: number; text: string; hasTextLayer: boolean }>
      }
      expect(whole.pageCount).toBe(2)
      expect(whole.info.title).toBe('E2E Handout')
      expect(whole.pages[0]!.text).toContain('Pdfium reads this')
      expect(whole.pages[0]!.text).not.toContain('\u0000')
      expect(whole.pages[1]!.hasTextLayer).toBe(false)

      const paged = payloadOf((await call({ path: pdfPath, pages: '2' })).body) as {
        pages: Array<{ page: number }>
      }
      expect(paged.pages.map((p) => p.page)).toEqual([2])
    } finally {
      await closeAndSaveVideo(launched, 'mcp-read-pdf')
    }
  })
})
