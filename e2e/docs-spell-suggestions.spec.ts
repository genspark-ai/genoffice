/**
 * Right-clicking a misspelled word in the document body lists Chromium's
 * suggestions at the top of the Docs context menu, and picking one replaces
 * the word through Blink (one Undo step takes it back, and the other
 * misspellings of the paragraph keep their squiggles).
 */
import { test, expect } from '@playwright/test'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const POLL = { timeout: 15_000, intervals: [250, 500, 1000] }

/**
 * Squiggles are native markers invisible to the DOM, and under CI's xvfb
 * Chromium does not repaint them, so red pixels are no witness. The
 * context menu is: Chromium hands the menu the misspelling under the pointer
 * only when a marker is there, and the Docs menu then lists suggestions or
 * Add to Dictionary. Right-click the word, read the menu, dismiss it with a
 * click in the page's top-left margin (the menu opens below and to the right
 * of the pointer): after Add to Dictionary the respell kick shields
 * keystrokes for a moment, so an Escape could be swallowed.
 */
async function markedViaMenu(page: Page, at: { x: number; y: number }): Promise<boolean> {
  await page.mouse.click(at.x, at.y, { button: 'right' })
  const menu = page.locator('.ctx-menu')
  await menu.waitFor()
  await page.waitForTimeout(400)
  const marked =
    (await menu.locator('.ctx-item-strong').count()) > 0 ||
    (await menu.locator('.ctx-item', { hasText: 'Add to Dictionary' }).count()) > 0
  const pageBox = (await page.locator('.doc-page').boundingBox())!
  await page.mouse.click(pageBox.x + 10, pageBox.y + 10)
  await expect(menu).toHaveCount(0)
  return marked
}

const pageText = (page: Page) =>
  page.evaluate(() => document.querySelector('.doc-page')?.textContent ?? '')

/** viewport rect of the first occurrence of `word` in the document body */
const wordRect = (page: Page, word: string) =>
  page.evaluate((w) => {
    const root = document.querySelector('.doc-page')
    if (!root) return null
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const idx = node.textContent?.indexOf(w) ?? -1
      if (idx === -1) continue
      const range = document.createRange()
      range.setStart(node, idx)
      range.setEnd(node, idx + w.length)
      const r = range.getBoundingClientRect()
      return { x: r.left, y: r.top, w: r.width, h: r.height }
    }
    return null
  }, word)

const center = (r: { x: number; y: number; w: number; h: number }) => ({
  x: r.x + r.w / 2,
  y: r.y + r.h / 2,
})

// the native spellchecker is outside the app's control
test.describe.configure({ retries: 1 })

test('context menu offers spelling suggestions and applies one', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'spell-suggestions' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, '://docs/')
    const docPage = editor.locator('.doc-page[contenteditable="true"][spellcheck="true"]')
    await docPage.waitFor()

    // native menus are invisible to the DOM: count Menu.popup calls in the main process
    await app.evaluate(({ Menu }) => {
      const g = globalThis as unknown as { __popups: number }
      g.__popups = 0
      const proto = Menu.prototype as unknown as { popup: (...a: unknown[]) => unknown }
      const orig = proto.popup
      proto.popup = function (...args: unknown[]) {
        g.__popups++
        return orig.apply(this, args)
      }
    })
    const nativePopups = () =>
      app.evaluate(() => (globalThis as unknown as { __popups: number }).__popups)

    // a word no dictionary has, unique per run: Add to Dictionary writes to the
    // OS custom dictionary on some platforms, so a fixed word would stop being
    // flagged after the first run on a developer machine
    const dictWord = `qzv${Date.now() % 100000}x`
    await docPage.click()
    await editor.keyboard.type(`The quick brown fox jumsp over the ${dictWord} lazyy dog `, {
      delay: 20,
    })
    await expect.poll(() => pageText(editor), POLL).toContain('lazyy dog')

    // give Blink a moment to mark the word before the first right-click
    await editor.waitForTimeout(1500)
    const targetRect = await wordRect(editor, 'jumsp')
    expect(targetRect).not.toBeNull()
    const target = center(targetRect!)

    let suggestion: string | null = null
    for (let attempt = 0; attempt < 6 && !suggestion; attempt++) {
      await editor.mouse.click(target.x, target.y, { button: 'right' })
      const menu = editor.locator('.ctx-menu')
      await menu.waitFor()
      const strong = menu.locator('.ctx-item-strong .ctx-label')
      try {
        await strong.first().waitFor({ timeout: 2000 })
        suggestion = await strong.first().textContent()
      } catch {
        await editor.keyboard.press('Escape')
        await editor.waitForTimeout(1000)
      }
    }
    test.skip(!suggestion, 'native spellchecker produced no suggestions in this environment')
    await editor.screenshot({ path: test.info().outputPath('spell-menu.png') })

    // the custom menu claimed the click: no native menu popped alongside it
    await expect(editor.locator('.ctx-menu')).toHaveCount(1)
    await expect(editor.locator('.ctx-menu')).toContainText('Add to Dictionary')
    expect(await nativePopups()).toBe(0)

    // two right-clicks in quick succession: each click is matched to its own
    // context-menu event, so the first one must not fall through to the native menu
    const otherRect = await wordRect(editor, dictWord)
    const other = center(otherRect!)
    await editor.mouse.click(target.x, target.y, { button: 'right' })
    await editor.mouse.click(other.x, other.y, { button: 'right' })
    await editor.locator('.ctx-menu').waitFor()
    await editor.waitForTimeout(700)
    expect(await nativePopups()).toBe(0)
    await expect(editor.locator('.ctx-menu')).toHaveCount(1)
    await editor.keyboard.press('Escape')
    // positive control for the popup counter: a surface without a React menu
    // (the AI pane input) still gets the native one
    const aiInput = editor.locator('textarea').first()
    await aiInput.click({ button: 'right' })
    await expect.poll(nativePopups, POLL).toBe(1)
    await editor.keyboard.press('Escape')

    // Add to Dictionary must clear the existing squiggle without the user
    // typing. Done before the replace/undo below: an undo re-renders the
    // paragraph DOM and Blink only re-marks on typing, so the typed word is
    // still marked here
    const blank = { x: other.x, y: other.y + 200 }
    await editor.mouse.click(blank.x, blank.y)
    await editor.waitForTimeout(500)
    const addToDict = editor.locator('.ctx-menu .ctx-item', { hasText: 'Add to Dictionary' })
    // Chromium reports the misspelling only once its marker is in place: retry the right-click
    for (let attempt = 0; attempt < 6 && !(await addToDict.count()); attempt++) {
      await editor.mouse.click(other.x, other.y, { button: 'right' })
      await editor.locator('.ctx-menu').waitFor()
      await editor.waitForTimeout(400)
      if (!(await addToDict.count())) {
        await editor.keyboard.press('Escape')
        await editor.waitForTimeout(800)
      }
    }
    await addToDict.click()
    await editor.mouse.click(blank.x, blank.y)
    await expect
      .poll(() => markedViaMenu(editor, other), {
        ...POLL,
        message: 'marker gone after Add to Dictionary',
      })
      .toBe(false)
    // the other misspelling keeps its marker: the dictionary changed, not the
    // checker (which re-walks the paragraph after the dictionary change)
    await expect
      .poll(() => markedViaMenu(editor, target), {
        ...POLL,
        message: 'marker kept on jumsp',
      })
      .toBe(true)

    // "jumsp" was typed too and never re-rendered, so its marker is still there
    await editor.mouse.click(target.x, target.y, { button: 'right' })
    await editor.locator('.ctx-menu .ctx-item-strong').first().waitFor()
    await editor.locator('.ctx-menu .ctx-item-strong').first().click()
    await expect.poll(() => pageText(editor), POLL).toContain(`fox ${suggestion} over`)
    await expect.poll(() => pageText(editor), POLL).not.toContain('jumsp')
    // the fix went through Blink, so the paragraph's other misspelling keeps
    // its marker (a scripted text-node rewrite would have dropped it). Blink
    // re-checks the edited paragraph on idle, so the marker can be briefly gone
    await editor.mouse.click(blank.x, blank.y)
    await expect
      .poll(async () => markedViaMenu(editor, center((await wordRect(editor, 'lazyy'))!)), {
        ...POLL,
        message: 'marker back on lazyy after the replacement',
      })
      .toBe(true)

    await editor.keyboard.press(process.platform === 'darwin' ? 'Meta+z' : 'Control+z')
    await expect.poll(() => pageText(editor), POLL).toContain('jumsp')
  } finally {
    await closeAndSaveVideo(launched)
  }
})
