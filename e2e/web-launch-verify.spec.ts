/**
 * Real-launch verification: the web version is started with the documented
 * `npm run dev -w @genoffice/<app>` commands (Electron main + Vite + HTTP
 * bridge), and a real Chromium browser drives the core flows over HTTP.
 *
 * Unlike web-bridge-browser.spec.ts, this spec does NOT spawn the dev apps —
 * it requires them to already be running (docs: 5173/5273, markdown: 5177/5277)
 * and verifies the documented startup path end to end.
 */
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { chromium, test, expect, type Browser } from '@playwright/test'
import JSZip from 'jszip'

const ROOT = resolve(__dirname, '..')
const DEV_PORT: Record<string, number> = { docs: 5173, markdown: 5177 }
const BRIDGE_PORT: Record<string, number> = { docs: 5273, markdown: 5277 }

async function launchBrowser(): Promise<Browser> {
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
): Promise<{
  status: number
  body: { ok?: boolean; result?: unknown; error?: { code?: string } }
}> {
  const response = await fetch(
    `http://127.0.0.1:${BRIDGE_PORT[app]}/api/ipc/${encodeURIComponent(channel)}`,
    {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ args }),
      signal: AbortSignal.timeout(15_000),
    },
  )
  return { status: response.status, body: await response.json() }
}

test.describe('real-launch web version (dev servers already running)', () => {
  test.setTimeout(240_000)

  test('docs: create → edit → save → reopen over HTTP in a real browser', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    const browser = await launchBrowser()
    try {
      // the documented startup command is already running: bridge + vite
      const health = await invokeHttp('docs', 'app:get-language')
      expect(health.status).toBe(200)
      expect(health.body.ok).toBe(true)

      // create a real document through the bridge
      const created = await invokeHttp('docs', 'docs:create-document', [
        { type: 'md', title: 'launch-verify', content: '# Launch verify' },
      ])
      expect(created.status).toBe(200)
      const createdPath = (created.body.result as { path?: string }).path
      expect(createdPath).toBeTruthy()

      // a real docx fixture for the edit flow
      const workDir = await mkdtemp(join(tmpdir(), 'web-launch-docs-'))
      const docPath = join(workDir, 'launch-verify.docx')
      await copyFile(join(ROOT, 'fixtures/generated/simple.docx'), docPath)

      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.docs}/`)
      await expect(page.locator('.doc-editor, #root *').first()).toBeVisible({ timeout: 60_000 })
      await page.waitForFunction(() => Boolean((window as { __aidocs?: unknown }).__aidocs))

      // load the fixture through the renderer pipeline (HTTP bridge)
      await page.evaluate((path: string) => {
        return (
          window as unknown as {
            __aidocs?: { openPath: (path: string) => Promise<unknown> }
          }
        ).__aidocs?.openPath(path)
      }, docPath)
      await expect(page.locator('.editor-scroll .ProseMirror')).toBeVisible({ timeout: 30_000 })

      // real UI edit
      await page.locator('.editor-scroll .ProseMirror').click()
      await page.keyboard.press('ControlOrMeta+End')
      await page.keyboard.type(' real launch edit')

      // renderer's own Cmd+S pipeline saves over HTTP
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
            const zip = await JSZip.loadAsync(Buffer.from(bytes))
            const xml = await zip.file('word/document.xml')?.async('string')
            return (xml ?? '').includes('real launch edit')
          },
          { timeout: 30_000, intervals: [1_000] },
        )
        .toBe(true)

      // reopen in a fresh page: the edit is on disk
      const page2 = await browser.newPage()
      await page2.goto(`http://localhost:${DEV_PORT.docs}/`)
      await expect(page2.locator('.doc-editor, #root *').first()).toBeVisible({ timeout: 60_000 })
      await page2.waitForFunction(() => Boolean((window as { __aidocs?: unknown }).__aidocs))
      await page2.evaluate((path: string) => {
        return (
          window as unknown as {
            __aidocs?: { openPath: (path: string) => Promise<unknown> }
          }
        ).__aidocs?.openPath(path)
      }, docPath)
      await expect(page2.locator('.editor-scroll .ProseMirror')).toContainText('real launch edit', {
        timeout: 30_000,
      })
    } finally {
      await browser.close()
    }
  })

  test('markdown: edit → save → reopen over HTTP in a real browser', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    const browser = await launchBrowser()
    try {
      const health = await invokeHttp('markdown', 'app:get-language')
      expect(health.status).toBe(200)
      expect(health.body.ok).toBe(true)

      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.markdown}/`)
      await expect(page.locator('.doc-editor')).toBeVisible({ timeout: 60_000 })

      // real UI edit
      await page.locator('.doc-editor').click()
      await page.keyboard.type('real launch md edit')

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
          suggestedName: 'launch-verify',
        })
      })
      expect(saveResult.ok).toBe(true)
      expect(saveResult.path).toBeTruthy()

      // server-side truth: the file exists with the typed content
      const readBack = await invokeHttp('markdown', 'markdown:read-file', [saveResult.path])
      expect(readBack.status).toBe(200)
      expect(String(readBack.body.result)).toContain('real launch md edit')

      // reopen in a fresh page
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
      expect(reopened).toContain('real launch md edit')
    } finally {
      await browser.close()
    }
  })

  test('web equivalents: native-only channels are wired to browser APIs, not WEB_UNSUPPORTED', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    const browser = await launchBrowser()
    try {
      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.docs}/`)
      await expect(page.locator('.doc-editor, #root *').first()).toBeVisible({ timeout: 60_000 })
      await page.waitForFunction(() => Boolean((window as { desktop?: unknown }).desktop))

      // font metrics (canvas measureText) resolves instead of rejecting
      const metrics = await page.evaluate(async () => {
        return await (
          window as unknown as {
            desktop: { fontMetrics: (family: string) => Promise<unknown> }
          }
        ).desktop.fontMetrics('Arial')
      })
      expect(metrics).toBeTruthy()

      // the open-dialog channel is overridden by the browser file picker
      const openFn = await page.evaluate(() => {
        const desktop = (window as unknown as { desktop: Record<string, unknown> }).desktop
        return typeof desktop.openDocx === 'function'
      })
      expect(openFn).toBe(true)
    } finally {
      await browser.close()
    }
  })
})
