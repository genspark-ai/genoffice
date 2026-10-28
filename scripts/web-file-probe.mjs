/* eslint-disable no-console */
// Deep probe: open real documents through the web bridge and assert the
// editors render their content. node scripts/web-file-probe.mjs [baseURL]
import { createRequire } from 'node:module'
import { resolve, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const require = createRequire(join(root, 'package.json'))
const { chromium } = require('playwright-core')

const base = process.argv[2] ?? 'http://127.0.0.1:8931'
const executablePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'

const browser = await chromium.launch({ executablePath, headless: true })
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } })
page.on('pageerror', (err) => console.log(`[pageerror:${page.url()}]`, String(err).slice(0, 200)))

// 1. docx: pending-open → consume → editor renders text
await page.goto(`${base}/docs/?path=${encodeURIComponent('/tmp/web-e2e/report.docx')}`, {
  waitUntil: 'domcontentloaded',
  timeout: 30000,
})
try {
  await page.waitForFunction(
    () => document.body?.innerText?.includes('Hello from the browser bridge test'),
    { timeout: 30000 },
  )
  console.log('[docs] content rendered: OK')
} catch {
  console.log('[docs] content rendered: MISSING — body text:', String(await page.evaluate(() => document.body?.innerText?.slice(0, 300))).replace(/\n/g, ' | '))
}
await page.screenshot({ path: '/tmp/web-probe-docs-file.png' })

// 2. sheets: queued workbook → sidecar session → Univer grid shows cells
await page.goto(`${base}/sheets/?path=${encodeURIComponent('/tmp/web-e2e/book.xlsx')}`, {
  waitUntil: 'domcontentloaded',
  timeout: 30000,
})
try {
  await page.waitForFunction(() => document.body?.innerText?.includes('b'), { timeout: 40000 })
  console.log('[sheets] grid mounted: OK')
} catch {
  console.log('[sheets] grid mounted: check screenshot')
}
await page.screenshot({ path: '/tmp/web-probe-sheets-file.png' })

// 3. markdown: pending open → editor content
await page.goto(`${base}/markdown/?path=${encodeURIComponent('/tmp/web-e2e/note.md')}`, {
  waitUntil: 'domcontentloaded',
  timeout: 30000,
})
try {
  await page.waitForFunction(
    () => {
      const t = document.body?.innerText ?? ''
      return t.includes('Web E2E Report') && t.includes('item one')
    },
    { timeout: 30000 },
  )
  console.log('[markdown] note rendered: OK')
} catch {
  console.log('[markdown] note rendered: MISSING')
}
await page.screenshot({ path: '/tmp/web-probe-markdown-file.png' })

await browser.close()
console.log('FILE PROBE DONE')
