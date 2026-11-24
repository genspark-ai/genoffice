/* eslint-disable no-undef */
// Probe the web UI with the system Chrome: load each editor page, wait for the
// renderer to boot, and report what came up. Run with the app's playwright-core:
//   node scripts/web-browser-probe.mjs [baseURL]
import { createRequire } from 'node:module'
import { resolve, dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const require = createRequire(join(root, 'package.json'))
const { chromium } = require('playwright-core')

const base = process.argv[2] ?? 'http://127.0.0.1:8931'
const token = process.env.WEB_TOKEN ?? ''
const withToken = (url) => `${url}${url.includes('?') ? '&' : '?'}token=${token}`

const executablePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'

async function probe(page, url, name, markers) {
  await page.goto(withToken(url), { waitUntil: 'domcontentloaded', timeout: 30000 })
  const outcomes = {}
  for (const marker of markers) {
    try {
      await page.waitForSelector(marker, { timeout: 25000 })
      outcomes[marker] = 'ok'
    } catch {
      outcomes[marker] = 'MISSING'
    }
  }
  const bridge = await page.evaluate(() => ({
    hasClient: Boolean(window.__genofficeWebClient),
    hasDesktop: Boolean(window.desktop),
    hasDesktopApi: Boolean(window.desktopApi),
    hasSlidesApi: Boolean(window.slidesApi),
    hasPdfApi: Boolean(window.pdfApi),
    hasMarkdownApi: Boolean(window.markdownApi),
    hasHtmlApi: Boolean(window.htmlApi),
  }))
  console.log(`[${name}]`, JSON.stringify({ outcomes, bridge }))
  await page.screenshot({ path: `/tmp/web-probe-${name}.png`, fullPage: false })
}

const browser = await chromium.launch({ executablePath, headless: true })
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } })
page.on('console', (msg) => {
  if (msg.type() === 'error') console.log(`[console:${page.url()}]`, msg.text().slice(0, 200))
})
page.on('pageerror', (err) => console.log(`[pageerror:${page.url()}]`, String(err).slice(0, 300)))

await probe(page, `${base}/docs/`, 'docs', ['#root'])
await probe(page, `${base}/sheets/?path=`, 'sheets', ['#root'])
await probe(page, `${base}/pdf/`, 'pdf', ['#root'])
await probe(page, `${base}/markdown/`, 'markdown', ['#root'])
await probe(page, `${base}/html/`, 'html', ['#root'])
await probe(page, `${base}/slides/`, 'slides', ['#root'])

await browser.close()
console.log('PROBE DONE')
