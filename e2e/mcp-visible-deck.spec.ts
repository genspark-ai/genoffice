import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { existsSync } from 'node:fs'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import http, { createServer } from 'node:http'
import JSZip from 'jszip'

/**
 * MCP visible deck session, end to end against the real built app.
 *
 * create_session family=pptx (opens a visible slides tab) → apply_slide_ops (title + bullets,
 * seen live by the user) → read_deck (element ids) → save_session (writes the
 * file). Asserts the on-disk pptx content and that a slides tab actually
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

test.describe('MCP visible deck session', () => {
  test('creates, edits and saves a pptx through a visible slides tab', async () => {
    const port = await freePort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-deck-'))
    const outFile = join(outDir, 'mcp-deck.pptx')
    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-mcp-deck-userdata-'))
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true, mcpEnabled: true, mcpPort: port }),
    )
    const launched = await launchShell({
      videoDir: 'mcp-visible-deck',
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

      // 1. open a visible blank deck — the session tools do not need the
      // background switch, so they are registered with the default settings
      const created = await call('create_session', { family: 'pptx' })
      expect(created.isError, created.text).toBeFalsy()

      // the shell UI shows a real slides tab (the visible half of the feature)
      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)
      const editorPage = await waitForPageWithUrl(app, 'genoffice-app://slides')

      // the visible tab really renders the deck (slide canvas mounted, blank
      // slide painted). Snapshot every painted canvas: MCP edits live in the
      // main-process session with no originating renderer, so the tab only
      // hears about them via the bridge's deck-changed push — the canvases
      // must actually change once the op below lands.
      await expect(editorPage.locator('canvas').first()).toBeVisible({ timeout: 15_000 })
      const canvasFingerprints = () =>
        editorPage.evaluate(() =>
          Array.from(document.querySelectorAll('canvas'), (el) => {
            try {
              return (el as HTMLCanvasElement).toDataURL()
            } catch {
              return ''
            }
          }),
        )
      await editorPage.waitForTimeout(300) // let the initial blank paint settle
      const beforePaint = await canvasFingerprints()
      expect(beforePaint.length).toBeGreaterThan(0)

      // 2. build slide 1 with one atomic transaction
      const EMU = 9525
      const applied = await call('apply_slide_ops', {
        ops: [
          {
            op: 'addElement',
            target: { slide: 0 },
            kind: 'textbox',
            offset: { x: EMU, y: 2 * EMU, cx: 30 * EMU, cy: 3 * EMU },
            paragraphs: [{ runs: [{ text: 'Deck From MCP', bold: true, fontSize: 32 }] }],
          },
        ],
      })
      expect(applied.isError, applied.text).toBeFalsy()
      expect(applied.text).toContain('"applied": true')
      // the render tree (base64 image data) must not come back in the tool
      // result: only the canvas push needs it
      expect(applied.text).not.toContain('"slides"')

      // …and the canvas actually repaints with it (this is the "watch the deck
      // build live" half of the feature; a swallowed broadcast leaves it blank)
      await expect
        .poll(
          async () => {
            const after = await canvasFingerprints()
            return after.some((fp, i) => fp !== '' && fp !== beforePaint[i])
          },
          { timeout: 15_000 },
        )
        .toBe(true)

      // 3. read the deck back: the element exists with a durable id
      const read = await call('read_deck', {})
      expect(read.isError, read.text).toBeFalsy()
      expect(read.text).toContain('Deck From MCP')

      // 4. output to an explicit path (ends the session)
      const saved = await call('save_session', { path: outFile })
      expect(saved.isError, saved.text).toBeFalsy()
      expect(existsSync(outFile)).toBe(true)

      // a second edit after the save reports the closed session
      const afterSave = await call('apply_slide_ops', {
        ops: [{ op: 'deleteSlide', target: { slide: 0 } }],
      })
      expect(afterSave.isError).toBe(true)

      // unpack the saved pptx and assert the content landed
      const zip = await JSZip.loadAsync(await readFile(outFile))
      const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string')
      expect(slideXml).toContain('Deck From MCP')
    } finally {
      await closeAndSaveVideo(launched, 'mcp-visible-deck')
    }
  })
})
