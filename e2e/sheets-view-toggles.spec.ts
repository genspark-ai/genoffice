import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * The View tab's Show group holds five toggles — gridlines, formula bar,
 * headings, right-to-left and the cross-highlight one. Stacked in a single
 * column they need 83px, and the ribbon band is 75px tall with
 * `overflow-y: hidden`, so the last one was sliced off and the reader
 * could not find it (genoffice#2081).
 *
 * The compact ribbon already lays this group out side by side; the normal
 * density did not. jsdom cannot compute layout, so the measurement has to
 * come from the real band.
 */
test('sheets: the View tab shows every display toggle', async () => {
  test.setTimeout(240_000)
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { lang: 'en' },
    videoDir: 'sheets-view-toggles',
  })
  try {
    const shellPage = launched.app.windows()[0]!
    await shellPage.evaluate(() =>
      (window as never as { aiOffice: { newSheet(): Promise<void> } }).aiOffice.newSheet(),
    )
    const page = await waitForPageWithUrl(launched.app, '://sheets/')
    await page.waitForSelector('[data-ribbon-body]', { timeout: 30_000 })
    await page.waitForTimeout(1_500)

    // the View tab
    const opened = await page.evaluate(() => {
      const el = [...document.querySelectorAll('.ribbon-tabs > button:not(.qa-btn)')].find(
        (x) => x.textContent?.trim() === 'View',
      ) as HTMLElement | undefined
      el?.click()
      return Boolean(el)
    })
    expect(opened, 'the View tab was not found').toBe(true)
    await page.waitForTimeout(800)

    const measured = await page.evaluate(() => {
      const bands = [...document.querySelectorAll('[data-ribbon-body]')] as HTMLElement[]
      const body = bands.find((b) => b.offsetParent !== null) ?? bands[0]
      if (!body) return null
      const column = body.querySelector('.check-column') as HTMLElement | null
      const items = column ? ([...column.querySelectorAll('.check-item')] as HTMLElement[]) : []
      const band = body.getBoundingClientRect()
      return {
        avail: body.clientHeight,
        scrollH: body.scrollHeight,
        clipped: body.scrollHeight > body.clientHeight + 1,
        count: items.length,
        labels: items.map((i) => i.textContent?.trim() ?? ''),
        // a toggle is only really there if it paints inside the band
        hidden: items
          .filter((i) => i.getBoundingClientRect().bottom > band.bottom + 1)
          .map((i) => i.textContent?.trim() ?? ''),
      }
    })
    expect(measured, 'no ribbon band on the View tab').not.toBeNull()
    const m = measured!

    expect(m.count, `the Show group holds ${m.count} toggles`).toBe(5)
    expect(
      m.labels[m.labels.length - 1],
      'the cross-highlight toggle is the one that was cut off',
    ).toContain('Highlight active row')
    expect(m.hidden, `these toggles paint below the band: ${m.hidden.join(', ')}`).toEqual([])
    expect(
      m.clipped,
      `the View band overflows: content needs ${m.scrollH}px, band has ${m.avail}px`,
    ).toBe(false)

    await page.screenshot({ path: screenshotPath('sheets-view-toggles') })
  } finally {
    await closeAndSaveVideo(launched, 'sheets-view-toggles')
  }
})
