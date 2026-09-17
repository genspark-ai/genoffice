import { test, expect, chromium } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

// Use each app's production styles in a minimal editor layout. Browser geometry
// catches grid/flex ordering and resize handles that jsdom cannot measure.
for (const app of ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'html']) {
  test(`${app}: switching AI panel sides preserves navigation and the collapsed rail`, async () => {
    const browser = await chromium.launch()
    try {
      const page = await browser.newPage({ viewport: { width: 1200, height: 800 } })
      const sheets = app === 'sheets'
      const panelClass = app === 'docs' || app === 'slides' ? 'ai-panel' : 'copilot'
      await page.setContent(
        `<div class="app-shell"><div class="${sheets ? 'sheet-body' : 'app-main'}" style="width:1000px;height:600px">` +
          (sheets
            ? '<aside class="copilot"><div class="ai-panel-resizer"></div><button class="ai-panel-collapse"><svg width="16" height="16"></svg></button>AI chat</aside>'
            : `<div class="ai-dock"><aside class="${panelClass}" style="width:360px"><div class="ai-panel-resizer"></div><button class="ai-panel-collapse"><svg width="16" height="16"></svg></button>AI chat</aside><button class="ai-rail" style="display:none">AI</button></div>`) +
          '<div class="fixture-document" style="flex:1;min-width:0;display:flex"><nav style="width:160px">Navigation</nav><article style="flex:1">Document</article></div></div></div>',
      )
      for (const path of [
        'packages/ui/src/tokens.css',
        'packages/ui/src/ai-panel-prefs.css',
        `apps/${app}/src/renderer/styles.css`,
      ]) {
        await page.addStyleTag({ content: await readFile(join(__dirname, '..', path), 'utf8') })
      }
      const dock = page.locator(sheets ? '.copilot' : '.ai-dock')
      const document = page.locator('.fixture-document')
      const nav = page.locator('nav')
      for (const side of ['left', 'right', 'left']) {
        await page.evaluate((side) => {
          document.documentElement.dataset.aiPanelSide = side
        }, side)
        const iconTransform = await page
          .locator('.ai-panel-collapse svg')
          .evaluate((icon) => getComputedStyle(icon).transform)
        expect(iconTransform).toBe(side === 'right' ? 'matrix(-1, 0, 0, 1, 0, 0)' : 'none')
        await expect.poll(async () => Math.round((await dock.boundingBox())!.width)).toBe(360)
        const box = (await dock.boundingBox())!
        const docBox = (await document.boundingBox())!
        expect(side === 'left' ? box.x < docBox.x : box.x > docBox.x).toBe(true)
        expect((await nav.boundingBox())!.x).toBeCloseTo(docBox.x, 0)
        const handle = (await page.locator('.ai-panel-resizer').boundingBox())!
        expect(
          Math.abs(
            side === 'right' ? handle.x - box.x : handle.x + handle.width - box.x - box.width,
          ),
        ).toBeLessThanOrEqual(4)
      }
      await page.evaluate((sheets) => {
        document.documentElement.dataset.aiPanelSide = 'right'
        if (sheets) {
          document.querySelector('.app-shell')!.classList.add('copilot-collapsed')
          document.querySelector('.copilot')!.classList.add('collapsed')
        } else {
          document.querySelector('.ai-dock')!.classList.add('collapsed')
          ;(document.querySelector('aside') as HTMLElement).style.display = 'none'
          ;(document.querySelector('.ai-rail') as HTMLElement).style.display = 'flex'
        }
      }, sheets)
      await expect.poll(async () => Math.round((await dock.boundingBox())!.width)).toBe(34)
      expect((await dock.boundingBox())!.x).toBeGreaterThan((await document.boundingBox())!.x)
    } finally {
      await browser.close()
    }
  })
}
