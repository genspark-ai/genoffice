/**
 * Insert > Text Box follows Word: "Draw Text Box" arms the crosshair draw mode
 * and a drag inserts a floating box of the dragged size with the caret inside
 * and Shape Format active; the Home tab has no Editing group on macOS.
 */
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const MAC = process.platform === 'darwin'
const POLL = { timeout: 15_000, intervals: [250, 500, 1000] }
const LETTER_REGIONS = new Set('US CA MX PH LR MM PR CL CO VE GT CR PA DO'.split(' '))

test('draw text box inserts a box of the dragged size ready to type', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-textbox-draw' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, '://docs/')
    const docPage = editor.locator('.doc-page[contenteditable="true"]')
    await docPage.waitFor()
    await docPage.click()

    // the blank document's paper follows the OS region like Word (Letter vs A4)
    const locale = await editor.evaluate(() =>
      (
        window as unknown as { desktop: { getSystemLocale(): Promise<string> } }
      ).desktop.getSystemLocale(),
    )
    const region = (() => {
      try {
        return new Intl.Locale(locale.replace(/_/g, '-')).maximize().region
      } catch {
        return undefined
      }
    })()
    const letter = !!region && LETTER_REGIONS.has(region)
    const pageW = await docPage.evaluate((el) =>
      Math.round(parseFloat(getComputedStyle(el).getPropertyValue('--page-w'))),
    )
    expect(pageW, `system locale ${locale}`).toBe(letter ? 816 : 794)

    await expect(editor.locator('.rb-editing')).toHaveCount(MAC ? 0 : 1)

    const activeTab = editor.locator('.ribbon-tab.active')
    // clicking the already selected tab collapses the ribbon (Word for Mac model)
    const openTextBoxMenu = async () => {
      if ((await activeTab.textContent()) !== 'Insert')
        await editor.locator('.ribbon-tab', { hasText: 'Insert' }).click()
      await editor.locator('.rb-big', { hasText: 'Text Box' }).click()
    }
    await openTextBoxMenu()
    const menu = editor.locator('.rb-textbox-menu')
    await expect(menu.locator('button')).toHaveText(['Draw Text Box', 'Simple Text Box'])
    await menu.locator('button', { hasText: 'Draw Text Box' }).click()
    await expect(docPage).toHaveClass(/doc-shape-drawing/)

    const pageBox = (await docPage.boundingBox())!
    const x0 = pageBox.x + 120
    const y0 = pageBox.y + 160
    await editor.mouse.move(x0, y0)
    await editor.mouse.down()
    await editor.mouse.move(x0 + 120, y0 + 60, { steps: 5 })
    await editor.mouse.move(x0 + 240, y0 + 120, { steps: 5 })
    await editor.mouse.up()

    const box = docPage.locator('.doc-textbox')
    await expect(box).toHaveCount(1)
    await expect(docPage).not.toHaveClass(/doc-shape-drawing/)
    await expect
      .poll(async () => {
        const b = await box.boundingBox()
        return b ? [Math.round(b.width), Math.round(b.height)] : null
      }, POLL)
      .toEqual([240, 120])

    await expect(activeTab).toHaveText('Shape Format')
    await expect
      .poll(
        () => editor.evaluate(() => !!document.activeElement?.closest('.doc-textbox-editor')),
        POLL,
      )
      .toBe(true)
    await editor.keyboard.type('Drawn', { delay: 10 })
    await expect(box.locator('.doc-textbox-editor')).toContainText('Drawn')
    await editor.screenshot({ path: test.info().outputPath('textbox-drawn.png') })

    // a plain click in draw mode inserts the predefined 5 x 3 cm box
    await editor.keyboard.press('Escape')
    await docPage.click({ position: { x: 300, y: 500 } })
    await openTextBoxMenu()
    await menu.locator('button', { hasText: 'Draw Text Box' }).click()
    await editor.mouse.click(pageBox.x + 200, pageBox.y + 520)
    await expect(box).toHaveCount(2)
    const sizeOf = (i: number) =>
      box.nth(i).evaluate((el) => [el.style.width, el.style.height].join(' '))
    await expect.poll(() => sizeOf(1), POLL).toBe('189px 113px')
    await expect
      .poll(
        () => editor.evaluate(() => !!document.activeElement?.closest('.doc-textbox-editor')),
        POLL,
      )
      .toBe(true)

    // Simple Text Box drops the same default box at the caret for keyboard users
    await editor.keyboard.press('Escape')
    await docPage.click({ position: { x: 300, y: 700 } })
    await openTextBoxMenu()
    await menu.locator('button', { hasText: 'Simple Text Box' }).click()
    await expect(box).toHaveCount(3)
    await expect.poll(() => sizeOf(2), POLL).toBe('189px 113px')
    await expect(activeTab).toHaveText('Shape Format')
  } finally {
    await closeAndSaveVideo(launched, 'docs-textbox-draw')
  }
})
