import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

const MAC = process.platform === 'darwin'
/** ⌥⌘K on macOS, Ctrl+Alt+K elsewhere — CI runs Linux. */
const COMPACT_KEY = MAC ? 'Alt+Meta+KeyK' : 'Control+Alt+KeyK'

/**
 * The compact ribbon must fit what it holds (genoffice#1891, docs parity).
 *
 * The sheets band is the `.ribbon[data-ribbon-body]` element — fixed-height
 * with `overflow-y: hidden`, so any control taller than the compact band is
 * sliced off rather than scrolled to. The tall buttons drop to their icon
 * plate, and every stacked group (`.ribbon-rows`, `.row-stack`,
 * `.styles-stack`, `.check-column`, `.tool-stack`, the chart-type grid) goes
 * side by side. This walks every tab in compact density and asserts nothing
 * overflows; jsdom cannot compute layout, so a unit test would pass while the
 * real band clipped.
 */
test('sheets: compact ribbon fits every tab without clipping', async () => {
  test.setTimeout(240_000)
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { lang: 'en' },
    videoDir: 'sheets-ribbon-compact-clip',
  })
  try {
    const shellPage = launched.app.windows()[0]!
    await shellPage.evaluate(() =>
      (window as never as { aiOffice: { newSheet(): Promise<void> } }).aiOffice.newSheet(),
    )
    const page = await waitForPageWithUrl(launched.app, '://sheets/')
    await page.waitForSelector('[data-ribbon-body]', { timeout: 30_000 })
    await page.waitForTimeout(1_500)

    // ⌥⌘K / Ctrl+Alt+K, sent the way the app receives it
    const compact = () => page.evaluate(() => !!document.querySelector('.ribbon-compact'))
    for (let i = 0; i < 3 && !(await compact()); i += 1) {
      await page.keyboard.press(COMPACT_KEY)
      await page.waitForTimeout(1_500)
    }
    expect(await compact(), 'compact density did not engage').toBe(true)

    const tabs = await page.evaluate(() =>
      [...document.querySelectorAll('.ribbon-tabs > button:not(.qa-btn)')]
        .map((t) => t.textContent?.trim() ?? '')
        .filter(Boolean),
    )
    expect(tabs.length).toBeGreaterThan(3)

    for (const tab of tabs) {
      await page.evaluate((name) => {
        const el = [...document.querySelectorAll('.ribbon-tabs > button:not(.qa-btn)')].find(
          (x) => x.textContent?.trim() === name,
        )
        ;(el as HTMLElement | undefined)?.click()
      }, tab)
      await page.waitForTimeout(600)
      // measure the band the reader is actually in: re-engage compact first
      if (!(await compact())) {
        await page.keyboard.press(COMPACT_KEY)
        await page.waitForTimeout(1_200)
      }
      expect(await compact(), `"${tab}" left compact density off`).toBe(true)

      const measured = await page.evaluate(() => {
        const bands = [...document.querySelectorAll('[data-ribbon-body]')] as HTMLElement[]
        const body = bands.find((b) => b.offsetParent !== null) ?? bands[0]
        if (!body) return null
        const tallest = [...body.querySelectorAll('*')]
          .map((el) => Math.round(el.getBoundingClientRect().height))
          .reduce((a, b) => Math.max(a, b), 0)
        return {
          avail: body.clientHeight,
          scrollH: body.scrollHeight,
          tallest,
          clipped: body.scrollHeight > body.clientHeight + 1,
        }
      })
      expect(measured, `no ribbon on the "${tab}" tab`).not.toBeNull()
      const m = measured!
      expect(
        m.clipped,
        `"${tab}" overflows the compact band: tallest control ${m.tallest}px, content needs ` +
          `${m.scrollH}px, band has ${m.avail}px`,
      ).toBe(false)
      expect(m.scrollH, `"${tab}" needs more than the band holds`).toBeLessThanOrEqual(m.avail)
    }
    await page.screenshot({ path: screenshotPath('sheets-ribbon-compact-all-tabs') })
  } finally {
    await closeAndSaveVideo(launched, 'sheets-ribbon-compact-clip')
  }
})
