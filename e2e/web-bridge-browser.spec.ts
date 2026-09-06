/**
 * Web version end-to-end — a REAL browser driving the SAME renderer build the
 * desktop app uses, against the standalone dev app's HTTP bridge.
 *
 * `npm run dev -w @genoffice/<app>` starts vite (5173/5177, /api proxied) plus
 * the Electron main process whose @genoffice/ipc-bridge serves /api/ipc/* on
 * 5273/5277. Google Chrome (playwright channel 'chrome') then opens the web
 * version: window.* bridges are installed by the web-bridge bootstrap over the
 * HTTP/SSE transport, and the flows below run entirely through HTTP.
 *
 * Covers acceptance items A3 (web end-to-end edit/save/reopen) and the
 * browser-side leg of A4/A5.
 */
import { spawn, type ChildProcess } from 'node:child_process'
import { copyFile, mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { chromium, test, expect, type Browser } from '@playwright/test'
import JSZip from 'jszip'

const ROOT = resolve(__dirname, '..')

const DEV_PORT: Record<string, number> = { docs: 5173, markdown: 5177 }
const BRIDGE_PORT: Record<string, number> = { docs: 5273, markdown: 5277 }

interface DevApp {
  child: ChildProcess
  stop: () => Promise<void>
}

function spawnDevApp(app: string): DevApp {
  // detached: the npm→electron-vite→electron tree forms one process group we
  // can tear down as a whole in stop()
  const child = spawn('npm', ['run', 'dev', '-w', `@genoffice/${app}`], {
    cwd: ROOT,
    detached: true,
    stdio: 'ignore',
    env: { ...process.env },
  })
  return {
    child,
    stop: async () => {
      if (child.pid) {
        try {
          process.kill(-child.pid, 'SIGTERM')
        } catch {}
        await new Promise((resolveDone) => setTimeout(resolveDone, 2_000))
        try {
          process.kill(-child.pid, 'SIGKILL')
        } catch {}
      }
    },
  }
}

async function waitFor(
  url: string,
  predicate: (body: unknown, status: number) => boolean,
  what: string,
): Promise<void> {
  const deadline = Date.now() + 120_000
  let lastError = 'never reached'
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url, { signal: AbortSignal.timeout(3_000) })
      const text = await response.text()
      let body: unknown = text
      try {
        body = JSON.parse(text)
      } catch {}
      if (predicate(body, response.status)) return
      lastError = `status ${response.status}`
    } catch (cause) {
      lastError = String(cause)
    }
    await new Promise((resolveWait) => setTimeout(resolveWait, 1_000))
  }
  throw new Error(`${what} never became ready (${lastError})`)
}

async function waitForBridge(app: string): Promise<void> {
  await waitFor(
    `http://127.0.0.1:${BRIDGE_PORT[app]}/api/ipc/health`,
    (body) => (body as { ok?: boolean }).ok === true,
    `${app} HTTP IPC bridge`,
  )
}

async function waitForDevServer(app: string): Promise<void> {
  await waitFor(
    `http://localhost:${DEV_PORT[app]}/`,
    (_body, status) => status === 200,
    `${app} vite dev server`,
  )
}

async function launchBrowser(): Promise<Browser> {
  // playwright-managed chromium when installed; otherwise the system Chrome
  try {
    return await chromium.launch({ channel: 'chrome' })
  } catch {
    return await chromium.launch()
  }
}

async function invokeHttp(
  app: string,
  channel: string,
  args: unknown[] = [],
): Promise<InvokeResult> {
  const response = await fetch(
    `http://127.0.0.1:${BRIDGE_PORT[app]}/api/ipc/${encodeURIComponent(channel)}`,
    {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ args }),
      signal: AbortSignal.timeout(10_000),
    },
  )
  return { status: response.status, body: await response.json() }
}

interface InvokeResult {
  status: number
  body: { ok?: boolean; result?: unknown; error?: { code?: string; message?: string } }
}

test.describe('web version in a real browser', () => {
  test.setTimeout(300_000)

  test('markdown: edit → save → reopen, entirely over HTTP', async () => {
    test.skip(process.platform === 'linux', 'dev-spawn + GUI app verified on desktop hosts')
    const app = spawnDevApp('markdown')
    let browser: Browser | null = null
    try {
      await waitForBridge('markdown')
      await waitForDevServer('markdown')

      browser = await launchBrowser()
      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.markdown}/`)
      // the renderer booted on the HTTP-backed window.markdownApi bridge
      await expect(page.locator('.doc-editor')).toBeVisible({ timeout: 60_000 })

      // real UI edit
      await page.locator('.doc-editor').click()
      await page.keyboard.type('Web bridge verify paragraph')

      // real save channel from the page — silent first save (no dialog in web)
      const saveResult = await page.evaluate(async () => {
        const text = (document.querySelector('.doc-editor')?.textContent ?? '').trim()
        return await (
          window as unknown as {
            markdownApi: {
              save: (request: unknown) => Promise<{ ok: boolean; path?: string; error?: string }>
            }
          }
        ).markdownApi.save({
          text,
          imageSources: [],
          mode: 'save',
          suggestedName: 'web-bridge-verify',
        })
      })
      expect(saveResult.ok).toBe(true)
      expect(saveResult.path).toBeTruthy()

      // server-side truth: the file exists with the typed content
      const readBack = await invokeHttp('markdown', 'markdown:read-file', [saveResult.path])
      expect(readBack.status).toBe(200)
      expect(String(readBack.body.result)).toContain('Web bridge verify paragraph')

      // reopen: a fresh browser page reads the saved document over HTTP
      const page2 = await browser.newPage()
      await page2.goto(`http://localhost:${DEV_PORT.markdown}/`)
      await expect(page2.locator('.doc-editor')).toBeVisible({ timeout: 60_000 })
      const reopened = await page2.evaluate(async (path: string) => {
        return await (
          window as unknown as {
            markdownApi: { readFile: (path: string) => Promise<string> }
          }
        ).markdownApi.readFile(path)
      }, saveResult.path!)
      expect(reopened).toContain('Web bridge verify paragraph')
    } finally {
      await browser?.close()
      await app.stop()
    }
  })

  test('docs: create → edit → save → reopen, entirely over HTTP', async () => {
    test.skip(process.platform === 'linux', 'dev-spawn + GUI app verified on desktop hosts')
    const app = spawnDevApp('docs')
    let browser: Browser | null = null
    try {
      await waitForBridge('docs')
      await waitForDevServer('docs')

      // A3 begins with a real document creation request, not an on-disk fixture.
      const created = await invokeHttp('docs', 'docs:create-document', [
        { type: 'md', title: 'web-bridge-created', content: '# Web bridge created' },
      ])
      expect(created.status).toBe(200)
      expect(created.body.result).toMatchObject({ ok: true })
      const createdPath = (created.body.result as { path?: string }).path
      expect(createdPath).toBeTruthy()
      expect(await readFile(createdPath!, 'utf8')).toContain('# Web bridge created')

      // a real docx fixture generated by packages/docx-engine scripts
      const workDir = await mkdtemp(join(tmpdir(), 'web-bridge-docs-'))
      const docPath = join(workDir, 'web-verify.docx')
      await copyFile(join(ROOT, 'fixtures/generated/simple.docx'), docPath)
      browser = await launchBrowser()
      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.docs}/`)
      await expect(page.locator('.doc-editor, #root *').first()).toBeVisible({ timeout: 60_000 })
      await page.waitForFunction(() => Boolean((window as { __aidocs?: unknown }).__aidocs))
      // load a real document through the renderer pipeline: its bytes cross the
      // bridge as binary and the editor state is initialized from them
      await page.evaluate((path: string) => {
        return (
          window as unknown as {
            __aidocs?: { openPath: (path: string) => Promise<unknown> }
          }
        ).__aidocs?.openPath(path)
      }, docPath)
      await expect(page.locator('.editor-scroll .ProseMirror')).toBeVisible({ timeout: 30_000 })

      // real UI edit
      await expect(page.locator('.editor-scroll .ProseMirror')).toBeVisible({ timeout: 30_000 })
      await page.locator('.editor-scroll .ProseMirror').click()
      await page.keyboard.press('ControlOrMeta+End')
      await page.keyboard.type(' HTTP bridge edit')

      // renderer's own Cmd+S pipeline serializes and calls saveDocx(path, bytes) over HTTP
      await page.keyboard.press('ControlOrMeta+s')
      await expect
        .poll(
          async () => {
            const readBack = await invokeHttp('docs', 'docs:read-path', [docPath])
            if (readBack.status !== 200) return false
            const encoded = readBack.body.result as
              { __ipcBytes?: string; b64?: string } | undefined
            if (encoded?.__ipcBytes !== 'ab' || typeof encoded.b64 !== 'string') return false
            const bytes = Buffer.from(encoded.b64, 'base64')
            // the docx is a zip — inspect word/document.xml, not raw bytes
            const zip = await JSZip.loadAsync(Buffer.from(bytes))
            const xml = await zip.file('word/document.xml')?.async('string')
            return (xml ?? '').includes('HTTP bridge edit')
          },
          { timeout: 30_000, intervals: [1_000] },
        )
        .toBe(true)

      // reopen: fresh page, same document, the edit is on disk
      const page2 = await browser.newPage()
      await page2.goto(`http://localhost:${DEV_PORT.docs}/`)
      await expect(page2.locator('.doc-editor, #root *').first()).toBeVisible({ timeout: 60_000 })
      await page2.waitForTimeout(1_000)
      await page2.evaluate((path: string) => {
        return (
          window as unknown as {
            __aidocs?: { openPath: (path: string) => Promise<unknown> }
          }
        ).__aidocs?.openPath(path)
      }, docPath)
      await expect(page2.locator('.editor-scroll .ProseMirror')).toContainText('HTTP bridge edit', {
        timeout: 30_000,
      })
    } finally {
      await browser?.close()
      await app.stop()
    }
  })
})
