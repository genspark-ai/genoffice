import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo } from './helpers'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { parseDocx } from '@genoffice/docx-engine'
import { openPptx } from '@genoffice/pptx-engine'
import JSZip from 'jszip'
import http, { createServer } from 'node:http'

/**
 * MCP headless generation against the real built app.
 *
 * The headless tools delegate to the bundled `genoffice` CLI, which the app
 * spawns on its own Node runtime (ELECTRON_RUN_AS_NODE) — this is the only test
 * that exercises that spawn path end to end: the app has to locate and run the
 * CLI, and the produced files are reparsed with the real engines.
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

test.describe('MCP headless generation', () => {
  test('spawns the bundled CLI to create docx/pptx/xlsx', async () => {
    test.setTimeout(120_000)
    const port = await freePort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-headless-'))
    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-headless-userdata-'))
    // background generation must be ON to expose the create_* tools
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({
        onboardingSeen: true,
        mcpEnabled: true,
        mcpPort: port,
        mcpBackground: true,
      }),
    )
    const launched = await launchShell({ videoDir: 'mcp-headless', userDataDir })
    try {
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
      await postMcp(port, sessionId, { jsonrpc: '2.0', method: 'notifications/initialized' })

      const payloadOf = (body: string): Record<string, unknown> => {
        const dataLine = body.split('\n').find((l) => l.startsWith('data: ')) ?? ''
        const reply = JSON.parse(dataLine.slice(6)) as {
          result?: { isError?: boolean; content?: Array<{ text?: string }> }
        }
        const text = reply.result?.content?.[0]?.text ?? ''
        if (reply.result?.isError) throw new Error(text)
        return JSON.parse(text) as Record<string, unknown>
      }
      const call = async (name: string, args: Record<string, unknown>) =>
        payloadOf(
          (
            await postMcp(port, sessionId, {
              jsonrpc: '2.0',
              id: Math.floor(Math.random() * 1e6),
              method: 'tools/call',
              params: { name, arguments: args },
            })
          ).body,
        )

      // create_docx → a real, reparsable .docx (reparse proves the CLI ran)
      const docxPath = join(outDir, 'headless.docx')
      const madeDoc = await call('create_docx', {
        title: 'Headless',
        content: '# Title\n\nBody paragraph.',
        path: docxPath,
      })
      expect(madeDoc.path).toBe(docxPath)
      const parsed = await parseDocx(new Uint8Array(await readFile(docxPath)))
      expect(parsed.blocks.filter((b) => !b.hidden)[0].type).toBe('heading')

      // read_docx reads it back through the CLI's docs reader
      const read = await call('read_docx', { path: docxPath })
      expect(String(read.text)).toContain('Title')
      expect(String(read.text)).toContain('Body paragraph')

      // create_pptx → a real 2-slide deck
      const pptxPath = join(outDir, 'headless.pptx')
      await call('create_pptx', {
        title: 'Deck',
        outline: '# One\n- a\n\n# Two\n- b',
        path: pptxPath,
      })
      const deck = await openPptx(new Uint8Array(await readFile(pptxPath)))
      expect(deck.deck.slides).toHaveLength(2)

      // create_xlsx → numeric cells survive the CLI's gateway write
      const xlsxPath = join(outDir, 'headless.xlsx')
      await call('create_xlsx', {
        title: 'Book',
        data: [
          ['Item', 'Qty'],
          ['Widget', 12],
        ],
        path: xlsxPath,
      })
      const zip = await JSZip.loadAsync(await readFile(xlsxPath))
      const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(sheetXml).toContain('<v>12</v>')
    } finally {
      await closeAndSaveVideo(launched, 'mcp-headless')
    }
  })
})
