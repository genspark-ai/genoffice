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

test('re-enabling spellcheck respells existing text without user input', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'spellcheck-reenable' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, 'docs/out')
    await editor.locator('.doc-page').waitFor()
    await wait(1500)

    await editor.locator('.doc-page').click()
    await editor.keyboard.type('Je vais a la mison ce soir', { delay: 20 })
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('encore la mison demain matin', { delay: 20 })
    await wait(2500)
    const baseline = await redCount(editor)
    const textBefore = await editor.evaluate(
      () => document.querySelector('.doc-page')?.textContent ?? '',
    )

    const spelling = editor.getByRole('button', { name: 'Spelling' })
    await editor.getByRole('button', { name: 'Review' }).click()
    await spelling.waitFor()

    await spelling.click()
    await wait(1500)
    const off = await redCount(editor)
    // no squiggles at all (e.g. the dictionary could not be provisioned in
    // this environment): the pixel assertion below would be meaningless
    test.skip(baseline < off + 100, 'native spellchecker inactive in this environment')

    // re-enable via the ribbon only — no click into the text, no typing
    await spelling.click()
    await wait(3500)
    const on = await redCount(editor)

    const textAfter = await editor.evaluate(
      () => document.querySelector('.doc-page')?.textContent ?? '',
    )

    expect(on).toBeGreaterThan(off + 100) // squiggles came back…
    expect(on).toBeGreaterThanOrEqual(Math.floor(baseline * 0.8)) // …on the existing lines
    expect(textAfter).toBe(textBefore) // and the respell kick left no trace
    expect(textAfter).not.toContain('​')
    expect(textAfter).not.toContain('  ')
  } finally {
    await closeAndSaveVideo(launched, 'spellcheck-reenable')
  }
})
