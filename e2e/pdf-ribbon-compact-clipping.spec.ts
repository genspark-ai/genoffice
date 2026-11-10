import { test, expect } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument } from 'pdf-lib'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

const MAC = process.platform === 'darwin'
/** ⌥⌘K on macOS, Ctrl+Alt+K elsewhere — CI runs Linux. */
const COMPACT_KEY = MAC ? 'Alt+Meta+KeyK' : 'Control+Alt+KeyK'

/**
 * The compact ribbon must fit what it holds (genoffice#1891, docs parity).
 *
 * `.ribbon-body` is a fixed-height band with `overflow-y: hidden`, and that is
 * deliberate: switching tabs must never resize the row. It is also a lid —
 * the docs compact band clipped its font pickers until every tall control got
 * a compact rule (genoffice#1942). This walks every pdf tab with the ribbon
 * in compact density and asserts nothing overflows. jsdom cannot compute
 * layout, so a unit test would pass while the real band clipped.
 */
test('pdf: compact ribbon fits every tab without clipping', async () => {
  test.setTimeout(240_000)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-pdf-ribbon-compact-'))
  const pdfPath = join(dir, 'blank.pdf')
  const blank = await PDFDocument.create()
  blank.addPage([612, 792])
  await writeFile(pdfPath, await blank.save({ useObjectStreams: false }))

  const launched = await launchShell({
    onboardingSeen: true,
    settings: { lang: 'en' },
    videoDir: 'pdf-ribbon-compact-clip',
    openFile: pdfPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://pdf/')
    await page.waitForSelector('.ribbon-body', { timeout: 30_000 })
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
      // measure the band the reader is actually in: re-engage compact first
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
    await page.screenshot({ path: screenshotPath('pdf-ribbon-compact-all-tabs') })
  } finally {
    await closeAndSaveVideo(launched, 'pdf-ribbon-compact-clip')
  }
})
