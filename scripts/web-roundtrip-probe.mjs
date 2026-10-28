/* eslint-disable no-console */
// Round-trip probe: open report.docx in the browser, type into the editor,
// save (Ctrl+S), and verify the file on disk gained the typed text.
import { createRequire } from 'node:module'
import { readFileSync } from 'node:fs'
import { resolve, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const require = createRequire(join(root, 'package.json'))
const { chromium } = require('playwright-core')

const base = process.argv[2] ?? 'http://127.0.0.1:8931'
const file = '/tmp/web-e2e/report.docx'
const executablePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'

const browser = await chromium.launch({ executablePath, headless: true })
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } })
page.on('pageerror', (err) => console.log('[pageerror]', String(err).slice(0, 200)))

await page.goto(`${base}/docs/?path=${encodeURIComponent(file)}`, {
  waitUntil: 'domcontentloaded',
  timeout: 30000,
})
await page.waitForFunction(
  () => document.body?.innerText?.includes('Hello from the browser bridge test'),
  { timeout: 30000 },
)
console.log('[docs] open: OK')

// click at the end of the first paragraph and type
await page.locator('.ProseMirror, [contenteditable="true"]').first().click()
await page.keyboard.type(' SAVED-FROM-BROWSER')
await new Promise((r) => setTimeout(r, 800))
await page.keyboard.press('Meta+s')
await new Promise((r) => setTimeout(r, 2500))

const bytes = readFileSync(file)
const { execFileSync } = await import('node:child_process')
const xml = execFileSync('unzip', ['-p', file, 'word/document.xml'], { maxBuffer: 64 * 1024 * 1024 }).toString('utf8')
const ok = xml.includes('SAVED-FROM-BROWSER')
console.log(`[docs] save round trip: ${ok ? 'OK' : 'MISSING'} (file ${bytes.length} bytes)`)
await page.screenshot({ path: '/tmp/web-probe-save.png' })
await browser.close()
console.log(ok ? 'ROUND TRIP DONE' : 'ROUND TRIP FAILED')
process.exit(ok ? 0 : 1)
