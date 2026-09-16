import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, SHELL_DIR } from './helpers'
import { existsSync } from 'node:fs'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import http, { createServer } from 'node:http'

/** repository root, for invoking the bundled CLI from a test */
const APP_ROOT = resolve(SHELL_DIR, '..', '..')

/**
 * MCP `open_documents`, end to end against the real built app.
 *
 * The point of the tool is reaching documents the user has open rather than the
 * blank tab a session owns, so this does NOT call create_session: it opens a
 * file through the app's own routing (as a user open would), then lists, reads
 * and closes it over MCP.
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

test.describe('MCP open documents', () => {
  test('lists, reads and closes a document the app opened', async () => {
    const port = await freePort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-open-docs-e2e-'))
    const docPath = join(outDir, 'opened.md')
    await writeFile(docPath, '# Handwritten\n\nA paragraph the agent must see.\n')

    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-open-docs-userdata-'))
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true, mcpEnabled: true, mcpPort: port }),
    )
    const launched = await launchShell({ videoDir: 'mcp-open-docs', userDataDir })
    const { page } = launched
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

      // an empty workspace lists as empty rather than erroring
      const empty = await call('open_documents', { action: 'list' })
      expect(empty.isError).toBeFalsy()
      expect(JSON.parse(empty.text).documents).toEqual([])

      // open the file through the app's own routing: a tab the session does not own
      const opened = await call('open_in_genoffice', { path: docPath })
      expect(opened.isError, opened.text).toBeFalsy()
      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)

      // 1. list — the open document is discoverable, with its type and path
      const listed = await call('open_documents', { action: 'list' })
      expect(listed.isError, listed.text).toBeFalsy()
      const documents = JSON.parse(listed.text).documents as Array<Record<string, unknown>>
      expect(documents).toHaveLength(1)
      expect(documents[0]!.type).toBe('Markdown document')
      expect(documents[0]!.path).toBe(docPath)
      expect(documents[0]!.saved).toBe(true)
      expect(documents[0]!.dirty).toBe(false)

      // 2. read — live content, not a disk read
      const read = await call('open_documents', { action: 'read', target: docPath })
      expect(read.isError, read.text).toBeFalsy()
      const payload = JSON.parse(read.text) as { content: string }
      expect(payload.content).toContain('Handwritten')
      expect(payload.content).toContain('A paragraph the agent must see')

      // an unknown target names what is actually open
      const missing = await call('open_documents', {
        action: 'read',
        target: join(outDir, 'nope.md'),
      })
      expect(missing.isError).toBe(true)
      expect(missing.text).toContain('no open document matches')

      // 3. close — the tab goes away and the list is empty again
      const closed = await call('open_documents', { action: 'close', target: docPath })
      expect(closed.isError, closed.text).toBeFalsy()
      expect(JSON.parse(closed.text).closed).toBe(true)
      await expect(editorTab).toHaveCount(0)

      const after = await call('open_documents', { action: 'list' })
      expect(JSON.parse(after.text).documents).toEqual([])

      // closing a clean document left the file untouched
      expect(await readFile(docPath, 'utf8')).toContain('Handwritten')
      expect(existsSync(docPath)).toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'mcp-open-docs')
    }
  })

  test('reads a Word document the app opened, through the docs bridge', async () => {
    const port = await freePort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-open-docs-docx-'))
    const docxPath = join(outDir, 'report.docx')
    // build the .docx with the repo's own CLI, the way a user's file would exist
    const { execFileSync } = await import('node:child_process')
    const sourceMd = join(outDir, 'source.md')
    await writeFile(sourceMd, '# Quarterly Report\n\nRevenue grew steadily.\n')
    execFileSync(
      process.execPath,
      [
        join(APP_ROOT, 'packages', 'cli', 'dist', 'genoffice.cjs'),
        'create',
        '--type',
        'docx',
        '--from',
        sourceMd,
        '--out',
        docxPath,
      ],
      { stdio: 'pipe' },
    )
    expect(existsSync(docxPath)).toBe(true)

    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-open-docs-docx-userdata-'))
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true, mcpEnabled: true, mcpPort: port }),
    )
    const launched = await launchShell({ videoDir: 'mcp-open-docs-docx', userDataDir })
    const { page } = launched
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
      const sessionId = String(init.headers['mcp-session-id'] ?? '')
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

      const opened = await call('open_in_genoffice', { path: docxPath })
      expect(opened.isError, opened.text).toBeFalsy()
      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')
      await expect(editorTab).toHaveCount(1)

      const listed = await call('open_documents', { action: 'list' })
      const documents = JSON.parse(listed.text).documents as Array<{
        type: string
        path: string
      }>
      expect(documents).toHaveLength(1)
      expect(documents[0]!.type).toBe('Word document')
      expect(documents[0]!.path).toBe(docxPath)

      // read goes through the docs bridge, which waits for the editor to mount
      const read = await call('open_documents', { action: 'read', target: docxPath })
      expect(read.isError, read.text).toBeFalsy()
      expect(JSON.parse(read.text).content).toBeTruthy()
      expect(read.text).toContain('Quarterly Report')

      const closed = await call('open_documents', { action: 'close', target: docxPath })
      expect(closed.isError, closed.text).toBeFalsy()
      await expect(editorTab).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, 'mcp-open-docs-docx')
    }
  })

  test('saves unsaved edits before closing when asked to close a dirty document', async () => {
    const port = await freePort()
    const outDir = await mkdtemp(join(tmpdir(), 'genoffice-open-docs-dirty-'))
    const docPath = join(outDir, 'edited.md')
    await writeFile(docPath, '# Original\n\nBefore the edit.\n')

    const userDataDir = await mkdtemp(join(tmpdir(), 'genoffice-open-docs-dirty-userdata-'))
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({ onboardingSeen: true, mcpEnabled: true, mcpPort: port }),
    )
    const launched = await launchShell({ videoDir: 'mcp-open-docs-dirty', userDataDir })
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
      const sessionId = String(init.headers['mcp-session-id'] ?? '')
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

      const opened = await call('open_in_genoffice', { path: docPath })
      expect(opened.isError, opened.text).toBeFalsy()
      const editorPage = await waitForPageWithUrl(app, 'genoffice-app://markdown')
      await expect(editorPage.locator('.doc-editor')).toBeVisible()

      // type into the live editor: the tab is now dirty, and the change exists
      // only in memory
      await editorPage.locator('.doc-editor').click()
      await editorPage.keyboard.press('ControlOrMeta+End')
      await editorPage.keyboard.press('Enter')
      await editorPage.keyboard.type('Added through the editor.')
      const editorTab = page.locator('.tab-bar .tab-item:not(.tab-home)')

      // the list reports the unsaved change
      await expect
        .poll(async () => {
          const listed = await call('open_documents', { action: 'list' })
          const documents = JSON.parse(listed.text).documents as Array<{ dirty: boolean }>
          return documents[0]?.dirty
        })
        .toBe(true)

      // read sees the in-memory text too, not the version on disk
      const read = await call('open_documents', { action: 'read', target: docPath })
      expect((JSON.parse(read.text) as { content: string }).content).toContain(
        'Added through the editor',
      )

      // ask for the disk file to be untouched up to the moment of the close
      expect(await readFile(docPath, 'utf8')).not.toContain('Added through the editor')

      // closing with the default action saves first, then closes
      const closed = await call('open_documents', { action: 'close', target: docPath })
      expect(closed.isError, closed.text).toBeFalsy()
      const closeResult = JSON.parse(closed.text) as { closed: boolean; savedTo?: string }
      expect(closeResult.closed).toBe(true)
      expect(closeResult.savedTo).toBe(docPath)
      await expect(editorTab).toHaveCount(0)

      // the edit reached the file
      expect(await readFile(docPath, 'utf8')).toContain('Added through the editor')
    } finally {
      await closeAndSaveVideo(launched, 'mcp-open-docs-dirty')
    }
  })
})
