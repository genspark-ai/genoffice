/**
 * Re-enabling Review → Spelling must restore red squiggles on EXISTING text
 * with no click or keystroke from the user (the earlier focus-cycle fix did
 * not actually do this). Squiggles are
 * native Chromium markers invisible to the DOM, so the assertion is
 * pixel-based: count red-ish pixels over the page.
 */
import { test, expect } from '@playwright/test'
import { PNG } from 'pngjs'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

async function redCount(page: Page): Promise<number> {
  const buf = await page.locator('.doc-page').screenshot()
  const png = PNG.sync.read(buf)
  let n = 0
  for (let i = 0; i < png.data.length; i += 4) {
    const r = png.data[i]!,
      g = png.data[i + 1]!,
      b = png.data[i + 2]!
    if (r > 140 && g < 110 && b < 110) n++
  }
  return n
}

const wait = (ms: number) => new Promise((r) => setTimeout(r, ms))

// squiggles worth of red pixels: the original delta the assertions were built on
const SQUIGGLE_PIXELS = 100
const POLL = { timeout: 15_000, intervals: [250, 500, 1000] }

const pageText = (page: Page) =>
  page.evaluate(() => document.querySelector('.doc-page')?.textContent ?? '')

/** Chromium paints markers word by word: wait for the count to pass `atLeast`, then to settle. */
async function settledRedCount(page: Page, atLeast = 0): Promise<number> {
  const deadline = Date.now() + POLL.timeout
  let last = await redCount(page)
  while (Date.now() < deadline) {
    await wait(500)
    const next = await redCount(page)
    if (next === last && next >= atLeast) return next
    last = next
  }
  return last
}

// belt and braces: the native spellchecker is outside the app's control
test.describe.configure({ retries: 1 })

test('re-enabling spellcheck respells existing text without user input', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'spellcheck-reenable' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, '://docs/')
    const docPage = editor.locator('.doc-page[contenteditable="true"][spellcheck="true"]')
    await docPage.waitFor()

    await docPage.click()
    await editor.keyboard.type('Je vais a la mison ce soir', { delay: 20 })
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('encore la mison demain matin', { delay: 20 })
    await expect.poll(() => pageText(editor), POLL).toContain('demain matin')

    // no squiggles at all (e.g. the dictionary could not be provisioned in
    // this environment): the pixel assertions below would be meaningless
    const baseline = await settledRedCount(editor, SQUIGGLE_PIXELS)
    test.skip(baseline < SQUIGGLE_PIXELS, 'native spellchecker inactive in this environment')
    const textBefore = await pageText(editor)

    const spelling = editor.getByRole('button', { name: 'Spelling' })
    await editor.getByRole('button', { name: 'Review' }).click()
    await spelling.waitFor()

    await spelling.click()
    await expect
      .poll(() => redCount(editor), { ...POLL, message: 'markers clear once spellcheck is off' })
      .toBeLessThanOrEqual(baseline - SQUIGGLE_PIXELS)
    const off = await settledRedCount(editor)

    // re-enable via the ribbon only — no click into the text, no typing
    await spelling.click()
    await expect
      .poll(() => redCount(editor), {
        ...POLL,
        message: 'existing text regains its spelling markers after re-enabling',
      })
      .toBeGreaterThanOrEqual(Math.max(off + SQUIGGLE_PIXELS + 1, Math.floor(baseline * 0.8)))

    // the markers return the moment the kick's trusted space lands; the kick
    // scrubs that space a beat later, so the "no trace" check must wait for it
    await expect
      .poll(() => pageText(editor), { ...POLL, message: 'the respell kick leaves no trace' })
      .toBe(textBefore)
    const textAfter = await pageText(editor)
    expect(textAfter).not.toContain('\u200b')
    expect(textAfter).not.toContain('  ')
  } finally {
    await closeAndSaveVideo(launched, 'spellcheck-reenable')
  }
})

test('toggling spellcheck never scrolls the view to the caret', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'spellcheck-noscroll' })
  const { app } = launched
  try {
    await launched.page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, '://docs/')
    await editor.locator('.doc-page').waitFor()
    await wait(1500)

    // two pages of short lines, caret ends up on the last page
    await editor.locator('.doc-page').click()
    for (let i = 0; i < 58; i++) {
      await editor.keyboard.type(`ligne ${i}`, { delay: 0 })
      await editor.keyboard.press('Enter')
    }
    await wait(1000)
    const pages = await editor.locator('.page-gap-inline, .page-gap').count()
    expect(pages).toBeGreaterThan(0) // the caret really sits pages below the top

    // look at the top of the document while the caret stays at the end
    const scrollTo = (y: number) =>
      editor.evaluate((top) => {
        const s = document.querySelector('.editor-scroll')
        if (s) s.scrollTop = top
        return s?.scrollTop ?? -1
      }, y)
    await scrollTo(0)
    await wait(300)

    const spelling = editor.getByRole('button', { name: 'Spelling' })
    await editor.getByRole('button', { name: 'Review' }).click()
    await spelling.waitFor()
    const before = await editor.evaluate(
      () => document.querySelector('.editor-scroll')?.scrollTop ?? -1,
    )
    await spelling.click() // off
    await wait(800)
    await spelling.click() // on again → respell kick types at the (off-screen) caret
    await wait(2500)
    const after = await editor.evaluate(
      () => document.querySelector('.editor-scroll')?.scrollTop ?? -1,
    )
    expect(Math.abs(after - before)).toBeLessThanOrEqual(2) // no jump to the caret's page
  } finally {
    await closeAndSaveVideo(launched, 'spellcheck-noscroll')
  }
})
