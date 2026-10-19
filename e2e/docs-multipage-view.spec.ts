/**
 * Word's Multiple Pages: zooming out flows the pages into columns, the live
 * editor sits on the active page and follows clicks and caret moves, and
 * zooming back in collapses the grid to the single column.
 */
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const MAC = process.platform === 'darwin'
const POLL = { timeout: 15_000, intervals: [250, 500, 1000] }

test('multiple pages grid follows clicks and the caret, collapses at 100%', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-multipage-view' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, '://docs/')
    const docPage = editor.locator('.doc-page[contenteditable="true"]')
    await docPage.waitFor()
    await docPage.click()
    await editor.keyboard.type('Page one.', { delay: 10 })
    await editor.keyboard.press('ControlOrMeta+Enter')
    await editor.keyboard.type('Page two.', { delay: 10 })
    await editor.keyboard.press('ControlOrMeta+Enter')
    await editor.keyboard.type('Page three.', { delay: 10 })
    await expect(editor.locator('.status-page')).toContainText('3', POLL)

    // ---- View ▸ Multiple Pages: at least two columns, the live page is the caret's ----
    await editor.locator('.ribbon-tab', { hasText: /^View$/ }).click()
    await editor.locator('.rb-big', { hasText: /^Multiple Pages$/ }).click()
    const grid = editor.locator('.doc-zoom.multipage')
    await grid.waitFor(POLL)
    const slots = editor.locator('.mp-slot')
    await expect(slots).toHaveCount(2, POLL)
    await editor.screenshot({ path: 'test-results/docs-multipage-view-grid.png' })
    await expect(editor.locator('.mp-slot[data-page="3"]')).toHaveCount(0)
    const lefts = await slots.evaluateAll((els) =>
      els.map((el) => Math.round(el.getBoundingClientRect().left)),
    )
    expect(new Set(lefts).size).toBe(2)
    const liveBox = (await editor.locator('.doc-live').boundingBox())!
    const slot2Box = (await editor.locator('.mp-slot[data-page="2"]').boundingBox())!
    expect(Math.abs(liveBox.height - slot2Box.height)).toBeLessThan(2)
    await expect(editor.locator('.mp-slot[data-page="1"] .mp-page')).toContainText('Page one.')
    await expect(docPage).toContainText('Page three.')
    await expect(editor.locator('.status-page')).toContainText('3')

    // ---- clicking a painted page moves the live editor there and lands the caret ----
    const slot1 = editor.locator('.mp-slot[data-page="1"]')
    const box1 = (await slot1.boundingBox())!
    await editor.mouse.click(box1.x + box1.width / 2, box1.y + box1.height * 0.1)
    await expect(editor.locator('.mp-slot[data-page="1"]')).toHaveCount(0, POLL)
    await expect(editor.locator('.mp-slot[data-page="3"]')).toHaveCount(1)
    await editor.keyboard.type(' Edited.', { delay: 10 })
    await expect(docPage.locator('p').first()).toContainText('Page one. Edited.')
    await expect(editor.locator('.status-page')).toContainText('1')
    await editor.screenshot({ path: 'test-results/docs-multipage-view.png' })

    // ---- a caret move onto another page turns the live window ----
    // macOS End only scrolls; ⌘↓ is the document-end caret move there
    await editor.keyboard.press(MAC ? 'Meta+ArrowDown' : 'Control+End')
    await expect(editor.locator('.mp-slot[data-page="3"]')).toHaveCount(0, POLL)
    await expect(editor.locator('.mp-slot[data-page="1"]')).toHaveCount(1)
    await expect(editor.locator('.mp-slot[data-page="1"] .mp-page')).toContainText(
      'Page one. Edited.',
      POLL,
    )

    // ---- One Page turns the arrangement off; a plain zoom-out alone keeps one column ----
    await editor.locator('.rb-big', { hasText: /^One Page$/ }).click()
    await expect(editor.locator('.doc-zoom.multipage')).toHaveCount(0, POLL)
    await editor.locator('.zoom-slider').focus()
    await editor.keyboard.press('Home')
    await expect(editor.locator('.zoom-value')).toHaveText('10%')
    await expect(editor.locator('.doc-zoom.multipage')).toHaveCount(0)
    await editor.locator('.rb-big', { hasText: /^Multiple Pages$/ }).click()
    await grid.waitFor(POLL)

    // ---- the Zoom dialog offers the mode; 100% collapses the grid ----
    await editor.locator('.zoom-value').click()
    const dialog = editor.locator('.zoom-dialog')
    await dialog.waitFor()
    await expect(dialog.locator('.zoom-dialog-radio', { hasText: 'Multiple Pages' })).toHaveCount(1)
    await dialog.locator('.zoom-dialog-radio', { hasText: '100%' }).click()
    await dialog.locator('button[type="submit"]').click()
    await expect(editor.locator('.doc-zoom.multipage')).toHaveCount(0, POLL)
    await expect(slots).toHaveCount(0)
    await expect(docPage).toContainText('Page three.')
  } finally {
    await closeAndSaveVideo(launched)
  }
})
