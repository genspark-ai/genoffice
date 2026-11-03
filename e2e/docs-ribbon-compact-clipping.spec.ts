import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

const MAC = process.platform === 'darwin'
/** ⌥⌘K on macOS, Ctrl+Alt+K elsewhere — CI runs Linux. */
const COMPACT_KEY = MAC ? 'Alt+Meta+KeyK' : 'Control+Alt+KeyK'

/**
 * The compact ribbon must fit what it holds (genoffice#1891).
 *
 * `.ribbon-body` is a fixed-height band with `overflow-y: hidden`, and that is
 * deliberate: switching tabs must never resize the row. It is also a lid. The
 * first compact mode only wrote rules for `.rb-big`, so every other control —
 * the font and size pickers, the style gallery cards, the split carets, the
 * page-layout columns — kept its full height and `overflow-y: hidden` sliced
 * the difference off the top and bottom. The font picker came out cut through
 * the middle: visible, clickable, and useless.
 *
 * This walks every tab with the ribbon in compact density and asserts nothing
 * overflows. It is the guard for the next control someone adds without a
 * compact rule: jsdom cannot compute layout, so a unit test here would pass
 * happily while the real grid clipped.
 */

test('compact ribbon fits every tab without clipping', async () => {
  test.setTimeout(240_000)
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { lang: 'en' },
    videoDir: 'ribbon-compact-clip',
  })
  try {
    // a blank document, from the home screen's own new-document action
    const shellPage = launched.app.windows()[0]!
    await shellPage.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await page.waitForFunction(() => document.body.textContent?.includes('Page') ?? false, null, {
      timeout: 30_000,
    })
    await page.waitForTimeout(1_500)

    // ⌥⌘K / Ctrl+Alt+K, sent the way the app receives it
    const compact = () => page.evaluate(() => !!document.querySelector('.ribbon-compact'))
    for (let i = 0; i < 3 && !(await compact()); i += 1) {
      await page.keyboard.press(COMPACT_KEY)
      await page.waitForTimeout(1_500)
    }
    expect(await compact(), 'compact density did not engage').toBe(true)

    const tabs = await page.evaluate(() =>
      [...document.querySelectorAll('.ribbon-tab')]
        .map((t) => t.textContent?.trim() ?? '')
        .filter(Boolean),
    )
    expect(tabs.length).toBeGreaterThan(3)

    for (const tab of tabs) {
      await page.evaluate((name) => {
        const el = [...document.querySelectorAll('.ribbon-tab')].find(
          (x) => x.textContent?.trim() === name,
        )
        ;(el as HTMLElement | undefined)?.click()
      }, tab)
      await page.waitForTimeout(600)
      // Switching tabs drops the density back to full — measure the band the
      // reader is actually in, so re-engage it before reading anything.
      if (!(await compact())) {
        await page.keyboard.press(COMPACT_KEY)
        await page.waitForTimeout(1_200)
      }
      expect(await compact(), `"${tab}" left compact density off`).toBe(true)

      const measured = await page.evaluate(() => {
        const body = document.querySelector('.ribbon-body')
        if (!body) return null
        const tallest = [...body.querySelectorAll('*')]
          .map((el) => Math.round(el.getBoundingClientRect().height))
          .reduce((a, b) => Math.max(a, b), 0)
        return {
          avail: body.clientHeight,
          scrollH: body.scrollHeight,
          tallest,
          clipped: body.scrollHeight > body.clientHeight + 1,
          culprits: [...body.querySelectorAll('*')]
            .map((el) => ({ el, r: el.getBoundingClientRect() }))
            .filter((x) => x.r.height > 20)
            .sort((a, b) => b.r.bottom - a.r.bottom)
            .slice(0, 5)
            .map(
              (x) =>
                `${(x.el.className || x.el.tagName).toString().slice(0, 60)} h=${Math.round(x.r.height)} top=${Math.round(x.r.top)}`,
            ),
          openPanels: [...document.querySelectorAll('[data-rb-panel]')].map(
            (p) => `${p.className} h=${Math.round(p.getBoundingClientRect().height)}`,
          ),
        }
      })
      expect(measured, `no ribbon on the "${tab}" tab`).not.toBeNull()
      const m = measured!
      expect(
        m.clipped,
        `"${tab}" overflows the compact band: tallest control ${m.tallest}px, content needs ` +
          `${m.scrollH}px, band has ${m.avail}px; deepest-bottom culprits: ` +
          `${(m.culprits ?? []).join(' | ')}; open panels: ${(m.openPanels ?? []).join(', ') || 'none'}`,
      ).toBe(false)
      expect(m.scrollH, `"${tab}" needs more than the band holds`).toBeLessThanOrEqual(m.avail)
    }
    await page.screenshot({ path: screenshotPath('ribbon-compact-all-tabs') })
  } finally {
    await closeAndSaveVideo(launched, 'ribbon-compact-clip')
  }
})

test('compact style cards show the full label, the menu the full card', async () => {
  test.setTimeout(180_000)
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { lang: 'en' },
    videoDir: 'ribbon-compact-style-cards',
  })
  try {
    const shellPage = launched.app.windows()[0]!
    await shellPage.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await page.waitForFunction(() => document.body.textContent?.includes('Page') ?? false, null, {
      timeout: 30_000,
    })

    // narrow enough that the gallery cannot hold every quick style, so the
    // "more styles" expander (and its menu of full-size cards) is guaranteed
    await launched.app.evaluate(async ({ BrowserWindow }, width) => {
      BrowserWindow.getAllWindows()[0]?.setSize(width, 700)
      await new Promise((r) => setTimeout(r, 300))
    }, 640)
    await page.waitForTimeout(1_000)

    const compact = () => page.evaluate(() => !!document.querySelector('.ribbon-compact'))
    for (let i = 0; i < 3 && !(await compact()); i += 1) {
      await page.keyboard.press(COMPACT_KEY)
      await page.waitForTimeout(1_500)
    }
    expect(await compact(), 'compact density did not engage').toBe(true)

    // The 30px compact card is label-only by design: the preview is clipped
    // out of the visual flow and the label must not ellipsize or spill — the
    // 46×30 card that stacked a 20px preview over the label clipped both.
    // The gallery renders every entry and clips the wrapped rows visually, so
    // "visible" means first-row (same offsetTop as the first card), which is
    // exactly the set the expander exists to reach.
    const { cards, visibleCount } = await page.evaluate(() => {
      const all = [...document.querySelectorAll('.style-gallery > .style-card')]
      const firstTop = all[0]?.offsetTop
      return {
        cards: all.map((card) => {
          const label = card.querySelector('.style-card-label') as HTMLElement
          const preview = card.querySelector('.style-card-preview') as HTMLElement
          const box = card.getBoundingClientRect()
          return {
            name: label.textContent ?? '',
            cardH: Math.round(box.height),
            labelFits: label.scrollWidth <= label.clientWidth + 1,
            contentFits: card.scrollHeight <= card.clientHeight + 1,
            previewHidden: preview.getBoundingClientRect().width <= 1,
          }
        }),
        visibleCount: all.filter((c) => c.offsetTop === firstTop).length,
      }
    })
    expect(cards.length, 'no style cards on the Home tab').toBeGreaterThan(3)
    for (const c of cards) {
      expect(c.cardH, `"${c.name}" is not the 30px compact card`).toBe(30)
      expect(c.labelFits, `"${c.name}" label is ellipsized in compact`).toBe(true)
      expect(c.contentFits, `"${c.name}" content spills out of its compact card`).toBe(true)
      expect(c.previewHidden, `"${c.name}" preview is still visible in compact`).toBe(true)
    }

    // The overflow menu reuses the card markup; it must keep the full 62px
    // card with the preview showing — the one place both halves still fit.
    expect(visibleCount, '640px window left every style card visible').toBeLessThan(cards.length)
    await page.click('.style-gallery-more')
    await page.waitForSelector('.style-gallery-menu')
    await page.waitForTimeout(300)
    const menuCards = await page.evaluate(() =>
      [...document.querySelectorAll('.style-gallery-menu .style-card')].map((card) => {
        const preview = card.querySelector('.style-card-preview') as HTMLElement
        const box = card.getBoundingClientRect()
        return {
          name: (card.querySelector('.style-card-label') as HTMLElement).textContent ?? '',
          cardH: Math.round(box.height),
          previewVisible: preview.getBoundingClientRect().width > 10,
          contentFits: card.scrollHeight <= card.clientHeight + 1,
        }
      }),
    )
    expect(menuCards.length, 'the overflow menu did not list the clipped styles').toBeGreaterThan(
      visibleCount,
    )
    for (const c of menuCards) {
      expect(c.cardH, `"${c.name}" menu card lost its full height in compact`).toBe(62)
      expect(c.previewVisible, `"${c.name}" menu preview is hidden in compact`).toBe(true)
      expect(c.contentFits, `"${c.name}" menu card clips its content`).toBe(true)
    }
    await page.screenshot({ path: screenshotPath('ribbon-compact-style-menu') })
  } finally {
    await closeAndSaveVideo(launched, 'ribbon-compact-style-cards')
  }
})
