/**
 * Real-launch verification for ALL six apps: the web version is started with
 * the documented `npm run dev -w @genoffice/<app>` commands (Electron main +
 * Vite + HTTP bridge), and a real Chromium browser drives the core flows over
 * HTTP. Unlike web-bridge-browser.spec.ts, this spec does NOT spawn the dev
 * apps — it requires them to already be running (docs 5173/5273, sheets
 * 5174/5274, slides 5175/5275, pdf 5176/5276, markdown 5177/5277, shell
 * 5199/5299) and verifies the documented startup path end to end.
 */
import { copyFile, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { chromium, test, expect, type Browser } from '@playwright/test'
import JSZip from 'jszip'

const ROOT = resolve(__dirname, '..')

const DEV_PORT: Record<string, number> = {
  docs: 5173,
  sheets: 5174,
  slides: 5175,
  pdf: 5176,
  markdown: 5177,
  shell: 5199,
}
const BRIDGE_PORT: Record<string, number> = {
  docs: 5273,
  sheets: 5274,
  slides: 5275,
  pdf: 5276,
  markdown: 5277,
  shell: 5299,
}

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
  body: { ok?: boolean; result?: unknown; error?: { code?: string; message?: string } }
}> {
  const response = await fetch(
    `http://127.0.0.1:${BRIDGE_PORT[app]}/api/ipc/${encodeURIComponent(channel)}`,
    {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ args }),
      signal: AbortSignal.timeout(20_000),
    },
  )
  return { status: response.status, body: await response.json() }
}

async function waitForBridge(app: string): Promise<void> {
  const deadline = Date.now() + 60_000
  let last = 'never reached'
  while (Date.now() < deadline) {
    try {
      const r = await invokeHttp(app, 'app:get-language')
      if (r.status === 200 && r.body.ok === true) return
      last = `status ${r.status}`
    } catch (cause) {
      last = String(cause)
    }
    await new Promise((done) => setTimeout(done, 1_000))
  }
  throw new Error(`${app} HTTP IPC bridge never became ready (${last})`)
}

/** Minimal one-page PDF (pdfjs tolerates the hand-written xref). */
function minimalPdf(text: string): Buffer {
  const content = `BT /F1 24 Tf 72 720 Td (${text}) Tj ET`
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  let pdf = '%PDF-1.4\n'
  const offsets: number[] = []
  objects.forEach((body, i) => {
    offsets.push(pdf.length)
    pdf += `${i + 1} 0 obj\n${body}\nendobj\n`
  })
  const xrefStart = pdf.length
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const off of offsets) pdf += `${String(off).padStart(10, '0')} 00000 n \n`
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF\n`
  return Buffer.from(pdf, 'latin1')
}

/** Minimal xlsx with one sheet containing a single cell value. */
async function minimalXlsx(value: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${value}</t></is></c></row></sheetData>
</worksheet>`,
  )
  return await zip.generateAsync({ type: 'nodebuffer' })
}

test.describe('real-launch web version — all six apps (dev servers already running)', () => {
  test.setTimeout(300_000)

  test('docs: create → edit → save → reopen over HTTP in a real browser', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    await waitForBridge('docs')
    const browser = await launchBrowser()
    try {
      const created = await invokeHttp('docs', 'docs:create-document', [
        { type: 'md', title: 'launch-all-docs', content: '# Launch all docs' },
      ])
      expect(created.status).toBe(200)
      const createdPath = (created.body.result as { path?: string }).path
      expect(createdPath).toBeTruthy()

      const workDir = await mkdtemp(join(tmpdir(), 'web-launch-all-docs-'))
      const docPath = join(workDir, 'launch-all.docx')
      await copyFile(join(ROOT, 'fixtures/generated/simple.docx'), docPath)

      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.docs}/`)
      await expect(page.locator('.doc-editor, #root *').first()).toBeVisible({ timeout: 60_000 })
      await page.waitForFunction(() => Boolean((window as { __aidocs?: unknown }).__aidocs))

      await page.evaluate((path: string) => {
        return (
          window as unknown as {
            __aidocs?: { openPath: (path: string) => Promise<unknown> }
          }
        ).__aidocs?.openPath(path)
      }, docPath)
      await expect(page.locator('.editor-scroll .ProseMirror')).toBeVisible({ timeout: 30_000 })

      await page.locator('.editor-scroll .ProseMirror').click()
      await page.keyboard.press('ControlOrMeta+End')
      await page.keyboard.type(' launch all edit')
      await page.keyboard.press('ControlOrMeta+s')
      await expect
        .poll(
          async () => {
            const readBack = await invokeHttp('docs', 'docs:read-path', [docPath])
            if (readBack.status !== 200) return false
            const encoded = readBack.body.result as
              { __ipcBytes?: string; b64?: string } | undefined
            if (encoded?.__ipcBytes !== 'ab' || typeof encoded.b64 !== 'string') return false
            const zip = await JSZip.loadAsync(Buffer.from(encoded.b64, 'base64'))
            const xml = await zip.file('word/document.xml')?.async('string')
            return (xml ?? '').includes('launch all edit')
          },
          { timeout: 30_000, intervals: [1_000] },
        )
        .toBe(true)

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
      await expect(page2.locator('.editor-scroll .ProseMirror')).toContainText('launch all edit', {
        timeout: 30_000,
      })
    } finally {
      await browser.close()
    }
  })

  test('markdown: edit → save → reopen over HTTP in a real browser', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    await waitForBridge('markdown')
    const browser = await launchBrowser()
    try {
      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.markdown}/`)
      await expect(page.locator('.doc-editor')).toBeVisible({ timeout: 60_000 })

      await page.locator('.doc-editor').click()
      await page.keyboard.type('launch all md edit')

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
          suggestedName: 'launch-all-md',
        })
      })
      expect(saveResult.ok).toBe(true)
      expect(saveResult.path).toBeTruthy()

      const readBack = await invokeHttp('markdown', 'markdown:read-file', [saveResult.path])
      expect(readBack.status).toBe(200)
      expect(String(readBack.body.result)).toContain('launch all md edit')

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
      expect(reopened).toContain('launch all md edit')
    } finally {
      await browser.close()
    }
  })

  test('sheets: web bridge + open workbook over HTTP in a real browser', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    await waitForBridge('sheets')
    const browser = await launchBrowser()
    try {
      const workDir = await mkdtemp(join(tmpdir(), 'web-launch-all-sheets-'))
      const xlsxPath = join(workDir, 'launch-all.xlsx')
      await writeFile(xlsxPath, await minimalXlsx('launch all sheets'))

      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.sheets}/`)
      await expect(page.locator('#root *').first()).toBeVisible({ timeout: 60_000 })
      await page.waitForFunction(() => Boolean((window as { desktopApi?: unknown }).desktopApi))

      // the HTTP-constructed bridge exposes the same surface as the preload
      const surface = await page.evaluate(() => {
        const api = (window as unknown as { desktopApi: Record<string, unknown> }).desktopApi
        return {
          hasSelectWorkbook: typeof api.selectWorkbook === 'function',
          hasSave: typeof api.saveWorkbookEdits === 'function',
          hasProject: Boolean((window as unknown as { projectApi?: unknown }).projectApi),
        }
      })
      expect(surface.hasSelectWorkbook).toBe(true)
      expect(surface.hasSave).toBe(true)
      expect(surface.hasProject).toBe(true)

      // open a real workbook through the bridge (the web equivalent of the
      // native open dialog: browser picker → temp file → open-path)
      const opened = await invokeHttp('sheets', 'workbook:open-path', [xlsxPath])
      expect(opened.status).toBe(200)
      expect(opened.body.ok).toBe(true)
      const session = opened.body.result as { sessionId?: string } | null
      expect(session?.sessionId).toBeTruthy()

      // the workbook session is live: read the opened cell back over HTTP
      const file = opened.body.result as {
        sessionId?: string
        sheets?: { id?: string; name?: string }[]
      }
      const sheetId = file.sheets?.[0]?.id
      expect(sheetId).toBeTruthy()
      const range = await invokeHttp('sheets', 'workbook:read-range', [
        {
          sessionId: file.sessionId,
          sheetId,
          range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        },
      ])
      expect(range.status).toBe(200)
      expect(JSON.stringify(range.body.result)).toContain('launch all sheets')
    } finally {
      await browser.close()
    }
  })

  test('slides: web bridge + new blank → render → save over HTTP in a real browser', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    await waitForBridge('slides')
    const browser = await launchBrowser()
    try {
      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.slides}/`)
      await expect(page.locator('#root *').first()).toBeVisible({ timeout: 60_000 })
      await page.waitForFunction(() => Boolean((window as { slidesApi?: unknown }).slidesApi))

      const surface = await page.evaluate(() => {
        const api = (window as unknown as { slidesApi: Record<string, unknown> }).slidesApi
        return {
          hasNewBlank: typeof api.newBlank === 'function',
          hasSave: typeof api.save === 'function',
          hasExportImages: typeof api.exportImages === 'function',
          hasProject: Boolean((window as unknown as { projectApi?: unknown }).projectApi),
        }
      })
      expect(surface.hasNewBlank).toBe(true)
      expect(surface.hasSave).toBe(true)
      expect(surface.hasExportImages).toBe(true)
      expect(surface.hasProject).toBe(true)

      // new blank deck through the bridge
      const blank = await invokeHttp('slides', 'slides:new-blank', [1200])
      expect(blank.status).toBe(200)
      const opened = blank.body.result as { slides?: unknown[]; size?: { cx: number; cy: number } }
      expect(Array.isArray(opened.slides)).toBe(true)
      expect(opened.slides!.length).toBeGreaterThan(0)

      // render slides over HTTP
      const rendered = await invokeHttp('slides', 'slides:get-render-slides')
      expect(rendered.status).toBe(200)
      expect(Array.isArray(rendered.body.result)).toBe(true)

      // save the deck (untitled first save lands in the drafts folder)
      const saved = await invokeHttp('slides', 'slides:save')
      expect(saved.status).toBe(200)
      expect((saved.body.result as { ok?: boolean }).ok).toBe(true)

      // web-native export overrides: pickExportDir returns a temp dir, and
      // exportImages writes PNGs the browser downloads
      const exportDir = await page.evaluate(async () => {
        return await (
          window as unknown as { slidesApi: { pickExportDir: () => Promise<string | null> } }
        ).slidesApi.pickExportDir()
      })
      expect(exportDir).toBeTruthy()
      const exported = await page.evaluate(async (dir: string) => {
        return await (
          window as unknown as {
            slidesApi: {
              exportImages: (op: unknown) => Promise<{ ok: boolean; paths?: string[] }>
            }
          }
        ).slidesApi.exportImages({
          dir,
          baseName: 'launch-all',
          pngsBase64: [
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
          ],
        })
      }, exportDir!)
      expect(exported.ok).toBe(true)
      expect(exported.paths?.length).toBe(1)
    } finally {
      await browser.close()
    }
  })

  test('pdf: web bridge + open via #open= → read → save over HTTP in a real browser', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    await waitForBridge('pdf')
    const browser = await launchBrowser()
    try {
      const workDir = await mkdtemp(join(tmpdir(), 'web-launch-all-pdf-'))
      const pdfPath = join(workDir, 'launch-all.pdf')
      await writeFile(pdfPath, minimalPdf('launch all pdf'))

      const page = await browser.newPage()
      await page.goto(
        `http://localhost:${DEV_PORT.pdf}/#open=${encodeURIComponent(pdfPath)}`,
      )
      await expect(page.locator('#root *').first()).toBeVisible({ timeout: 60_000 })
      await page.waitForFunction(() => Boolean((window as { pdfApi?: unknown }).pdfApi))

      const surface = await page.evaluate(() => {
        const api = (window as unknown as { pdfApi: Record<string, unknown> }).pdfApi
        return {
          hasReadFile: typeof api.readFile === 'function',
          hasSave: typeof api.save === 'function',
          hasProject: Boolean((window as unknown as { projectApi?: unknown }).projectApi),
        }
      })
      expect(surface.hasReadFile).toBe(true)
      expect(surface.hasSave).toBe(true)
      expect(surface.hasProject).toBe(true)

      // the renderer consumed the #open= path and loaded the document
      await expect(page.locator('.pdf-page, .pdf-viewer, canvas').first()).toBeVisible({
        timeout: 60_000,
      })

      // read the file back over HTTP (path granted by pdf:open-path)
      const readBack = await invokeHttp('pdf', 'pdf:read-file', [pdfPath])
      expect(readBack.status).toBe(200)
      expect(readBack.body.ok).toBe(true)

      // save a copy over HTTP
      const saved = await invokeHttp('pdf', 'pdf:save', [
        { path: pdfPath, markups: [], drawings: [], formValues: [], stamps: [] },
      ])
      expect(saved.status).toBe(200)
      expect((saved.body.result as { ok?: boolean }).ok).toBe(true)
    } finally {
      await browser.close()
    }
  })

  test('shell: web bridge + home/tabs over HTTP in a real browser', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    await waitForBridge('shell')
    const browser = await launchBrowser()
    try {
      const page = await browser.newPage()
      await page.goto(`http://localhost:${DEV_PORT.shell}/`)
      await expect(page.locator('#root *').first()).toBeVisible({ timeout: 60_000 })
      await page.waitForFunction(() => Boolean((window as { aiOffice?: unknown }).aiOffice))

      const surface = await page.evaluate(() => {
        const w = window as unknown as Record<string, unknown>
        const home = w.aiOffice as Record<string, unknown>
        const tabs = w.aiOfficeTabs as Record<string, unknown>
        return {
          hasBrowse: typeof home.browse === 'function',
          hasRecents: typeof home.recents === 'function',
          hasOpenPath: typeof home.openPath === 'function',
          hasTabsList: typeof tabs.list === 'function',
          hasTabsActivate: typeof tabs.activate === 'function',
          hasProject: Boolean(w.aiOfficeProject),
        }
      })
      expect(surface.hasBrowse).toBe(true)
      expect(surface.hasRecents).toBe(true)
      expect(surface.hasOpenPath).toBe(true)
      expect(surface.hasTabsList).toBe(true)
      expect(surface.hasTabsActivate).toBe(true)
      expect(surface.hasProject).toBe(true)

      // home data flows over HTTP (RecentPage: { entries, total, totalAll })
      const recents = await invokeHttp('shell', 'home:recents', [''])
      expect(recents.status).toBe(200)
      expect(Array.isArray((recents.body.result as { entries?: unknown[] }).entries)).toBe(true)

      // web-native tab management resolves (browser tabs, no WEB_UNSUPPORTED)
      const tabsResult = await page.evaluate(async () => {
        return await (
          window as unknown as { aiOfficeTabs: { list: () => Promise<unknown> } }
        ).aiOfficeTabs.list()
      })
      expect(Array.isArray(tabsResult)).toBe(true)
    } finally {
      await browser.close()
    }
  })

  test('native-only channels: browser equivalents installed (no WEB_UNSUPPORTED)', async () => {
    test.skip(process.platform === 'linux', 'GUI browser verified on desktop hosts')
    for (const app of ['docs', 'markdown', 'sheets', 'slides', 'pdf', 'shell']) {
      await waitForBridge(app)
    }
    const browser = await launchBrowser()
    try {
      // docs: font metrics (canvas) and print (browser dialog) are wired
      const docsPage = await browser.newPage()
      await docsPage.goto(`http://localhost:${DEV_PORT.docs}/`)
      await expect(docsPage.locator('.doc-editor, #root *').first()).toBeVisible({ timeout: 60_000 })
      await docsPage.waitForFunction(() => Boolean((window as { desktop?: unknown }).desktop))
      const fontMetrics = await docsPage.evaluate(async () => {
        return await (
          window as unknown as {
            desktop: { fontMetrics: (family: string) => Promise<unknown> }
          }
        ).desktop.fontMetrics('Arial')
      })
      expect(fontMetrics).toBeTruthy()
      await docsPage.close()

      // sheets: CSV save confirm resolves without a native dialog
      const sheetsPage = await browser.newPage()
      await sheetsPage.goto(`http://localhost:${DEV_PORT.sheets}/`)
      await expect(sheetsPage.locator('#root *').first()).toBeVisible({ timeout: 60_000 })
      await sheetsPage.waitForFunction(() => Boolean((window as { desktopApi?: unknown }).desktopApi))
      const csvChoice = await sheetsPage.evaluate(async () => {
        return await (
          window as unknown as { desktopApi: { confirmCsvSave: () => Promise<unknown> } }
        ).desktopApi.confirmCsvSave()
      })
      expect(csvChoice).toBe('csv')
      await sheetsPage.close()

      // slides: fullscreen override resolves (browser requestFullscreen path)
      const slidesPage = await browser.newPage()
      await slidesPage.goto(`http://localhost:${DEV_PORT.slides}/`)
      await expect(slidesPage.locator('#root *').first()).toBeVisible({ timeout: 60_000 })
      await slidesPage.waitForFunction(() => Boolean((window as { slidesApi?: unknown }).slidesApi))
      const fullscreen = await slidesPage.evaluate(async () => {
        try {
          await (
            window as unknown as { slidesApi: { setShowFullScreen: (on: boolean) => Promise<unknown> } }
          ).slidesApi.setShowFullScreen(false)
          return 'resolved'
        } catch (error) {
          return String(error)
        }
      })
      expect(fullscreen).not.toMatch(/WEB_UNSUPPORTED|desktop/i)
      await slidesPage.close()

      // shell: reveal-path resolves without a native dialog
      const shellPage = await browser.newPage()
      await shellPage.goto(`http://localhost:${DEV_PORT.shell}/`)
      await expect(shellPage.locator('#root *').first()).toBeVisible({ timeout: 60_000 })
      await shellPage.waitForFunction(() => Boolean((window as { aiOffice?: unknown }).aiOffice))
      const reveal = await shellPage.evaluate(async () => {
        try {
          await (
            window as unknown as { aiOffice: { revealPath: (path: string) => Promise<unknown> } }
          ).aiOffice.revealPath('/tmp/nonexistent')
          return 'resolved'
        } catch (error) {
          return String(error)
        }
      })
      expect(reveal).not.toMatch(/WEB_UNSUPPORTED|desktop/i)
      await shellPage.close()
    } finally {
      await browser.close()
    }
  })
})
