import { mkdirSync } from 'node:fs'
import { join } from 'node:path'
import { chromium } from 'playwright-core'

/**
 * Renders the preview pages with headless Chromium and writes PNGs to
 * /tmp/mods-preview. The pages bundle the real components and the real
 * stylesheet, so what is captured is the product's rendering.
 */
const out = '/tmp/mods-preview'
mkdirSync(out, { recursive: true })
const browser = await chromium.launch()
const base = 'file://' + join(process.cwd(), 'e2e/panel-preview')
const pages = [
  { name: 'mods-registry', url: base + '/index.html', theme: 'light' },
  { name: 'mods-registry-dark', url: base + '/index.html', theme: 'dark' },
  { name: 'mods-empty', url: base + '/index.html?state=empty', theme: 'light' },
  { name: 'slots-plain', url: base + '/slots.html?plain=1', theme: 'light' },
  { name: 'slots', url: base + '/slots.html', theme: 'light' },
  { name: 'slots-dark', url: base + '/slots.html', theme: 'dark' },
  { name: 'slots-limited', url: base + '/slots.html?plain=1&host=limited', theme: 'light' },
  { name: 'palette', url: base + '/palette.html', theme: 'light' },
]
for (const entry of pages) {
  const page = await browser.newPage({
    viewport: { width: 1360, height: 860 },
    colorScheme: entry.theme,
  })
  await page.goto(entry.url)
  await page.waitForTimeout(700)
  const path = join(out, entry.name + '.png')
  await page.screenshot({ path })
  console.log('SHOT', path)
  // the panel's script section is two steps by design: pick a document, dry-run,
  // and only then does apply come alive. Drive it for real in the panel page.
  if (entry.name === 'mods-registry') {
    const pick = page.locator('.mods-script-pick').first()
    if (await pick.count()) {
      await pick.selectOption({ index: 1 })
      await page.locator('.mods-script-btn').first().click()
      await page.waitForTimeout(500)
      const runPath = join(out, 'mods-scripts-run.png')
      await page.screenshot({ path: runPath })
      console.log('SHOT', runPath)
      await page.locator('.mods-script-btn--apply').first().click()
      await page.waitForTimeout(500)
      const appliedPath = join(out, 'mods-scripts-applied.png')
      await page.screenshot({ path: appliedPath })
      console.log('SHOT', appliedPath)

      // the import review: what a pack declares, before anything is installed
      await page.locator('.mods-ghost-btn').nth(1).click()
      await page.waitForTimeout(400)
      const packPath = join(out, 'mods-pack-review.png')
      await page.screenshot({ path: packPath })
      console.log('SHOT', packPath)
    }
  }
  await page.close()
}
await browser.close()
