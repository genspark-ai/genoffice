/**
 * Word context-menu and Review-tab parity: right-clicking a hyperlink offers
 * Edit / Open / Copy / Remove Hyperlink instead of the generic "Hyperlink…",
 * Copy puts the address on the clipboard, Remove keeps the text; Review >
 * Comments has Delete ▾ / Previous / Next that walk the comment anchors and
 * light the thread up in the pane.
 */
import { test, expect } from '@playwright/test'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const POLL = { timeout: 15_000, intervals: [250, 500, 1000] }

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
      return { x: r.left + r.width / 2, y: r.top + r.height / 2 }
    }
    return null
  }, word)

const selectedText = (page: Page) => page.evaluate(() => window.getSelection()?.toString() ?? '')

/** select the first occurrence of `word` through the editor (a mouse double-click's word rules vary per platform) */
const selectWord = (page: Page, word: string) =>
  page.evaluate((w) => {
    const ed = (window as unknown as { __aidocs: { editor: any } }).__aidocs.editor
    let hit: { from: number; to: number } | null = null
    ed.state.doc.descendants((node: any, pos: number) => {
      if (hit || !node.isText) return
      const idx = (node.text as string).indexOf(w)
      if (idx !== -1) hit = { from: pos + idx, to: pos + idx + w.length }
    })
    if (hit) ed.chain().focus().setTextSelection(hit).run()
    return hit
  }, word)

const menuLabels = (page: Page) =>
  page.locator('.ctx-menu .ctx-label').evaluateAll((els) => els.map((el) => el.textContent))

test('hyperlink context menu items and Review comment navigation', async () => {
  test.setTimeout(150_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'ctx-word-items' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, '://docs/')
    const docPage = editor.locator('.doc-page[contenteditable="true"]')
    await docPage.waitFor()
    await docPage.click()
    await editor.keyboard.type('Visit example now. Then alpha and beta end.', { delay: 15 })
    await expect
      .poll(() => editor.evaluate(() => document.querySelector('.doc-page')?.textContent), POLL)
      .toContain('beta end.')

    // ---- link: select "example", ⌘K, address, Enter ----
    const example = (await wordRect(editor, 'example'))!
    expect(await selectWord(editor, 'example')).not.toBeNull()
    await expect.poll(() => selectedText(editor), POLL).toBe('example')
    await editor.keyboard.press(process.platform === 'darwin' ? 'Meta+k' : 'Control+k')
    const modal = editor.locator('.modal')
    await modal.waitFor()
    const url = 'https://example.com/e2e-link'
    await modal.locator('input').nth(1).fill(url)
    await modal.locator('input').nth(1).press('Enter')
    await expect(editor.locator('.doc-page a.doc-link')).toHaveCount(1)
    await expect(editor.locator('.doc-page a.doc-link')).toHaveText('example')

    // ---- right-click on the link: Word's four items replace "Hyperlink…" ----
    await editor.mouse.click(example.x, example.y, { button: 'right' })
    const menu = editor.locator('.ctx-menu')
    await menu.waitFor()
    const labels = await menuLabels(editor)
    expect(labels).toEqual(
      expect.arrayContaining([
        'Edit Hyperlink…',
        'Open Hyperlink',
        'Copy Hyperlink',
        'Remove Hyperlink',
      ]),
    )
    expect(labels).not.toContain('Hyperlink…')
    await menu.locator('.ctx-item', { hasText: 'Copy Hyperlink' }).click()
    await expect.poll(() => app.evaluate(({ clipboard }) => clipboard.readText()), POLL).toBe(url)

    await editor.mouse.click(example.x, example.y, { button: 'right' })
    await menu.waitFor()
    await menu.locator('.ctx-item', { hasText: 'Remove Hyperlink' }).click()
    await expect(editor.locator('.doc-page a.doc-link')).toHaveCount(0)
    await expect(docPage).toContainText('Visit example now.')

    // off a link the generic item is back
    const now = (await wordRect(editor, 'now'))!
    await editor.mouse.click(now.x, now.y, { button: 'right' })
    await menu.waitFor()
    expect(await menuLabels(editor)).toContain('Hyperlink…')
    await editor.keyboard.press('Escape')

    // ---- two comments on "alpha" and "beta" ----
    await editor.locator('.ribbon-tab', { hasText: 'Review' }).click()
    const newComment = editor.locator('.rb-big', { hasText: 'New Comment' })
    const compose = editor.locator('.comment-compose')
    for (const [word, text] of [
      ['alpha', 'first note'],
      ['beta', 'second note'],
    ] as const) {
      expect(await selectWord(editor, word)).not.toBeNull()
      await expect.poll(() => selectedText(editor), POLL).toBe(word)
      await newComment.click()
      await compose.locator('textarea').fill(text)
      await compose.locator('.comment-compose-actions .primary').click()
      await expect(editor.locator('.comment-thread')).toHaveCount(word === 'alpha' ? 1 : 2)
    }

    // ---- Previous / Next walk the anchors and focus the pane card ----
    const next = editor.locator('button[data-tip="Go to the next comment"]')
    const prev = editor.locator('button[data-tip="Go to the previous comment"]')
    await editor.evaluate(() =>
      (
        window as unknown as { __aidocs: { editor: any } }
      ).__aidocs.editor.commands.setTextSelection(1),
    )
    await next.click()
    await expect.poll(() => selectedText(editor), POLL).toBe('alpha')
    await expect(editor.locator('.comment-card.active')).toContainText('first note')
    await next.click()
    await expect.poll(() => selectedText(editor), POLL).toBe('beta')
    await expect(editor.locator('.comment-card.active')).toContainText('second note')
    await prev.click()
    await expect.poll(() => selectedText(editor), POLL).toBe('alpha')

    // ---- Delete ▾: the thread at the caret, then all ----
    const deleteSplit = editor.locator('button[data-tip="Delete the comment at the cursor"]')
    await deleteSplit.click()
    await editor.locator('.layout-menu button', { hasText: /^Delete$/ }).click()
    await expect(editor.locator('.comment-thread')).toHaveCount(1)
    await expect(editor.locator('.comment-thread')).toContainText('second note')
    await deleteSplit.click()
    await editor
      .locator('.layout-menu button', { hasText: 'Delete All Comments in Document' })
      .click()
    await expect(editor.locator('.comment-thread')).toHaveCount(0)
    await expect(editor.locator('.doc-page .doc-comment')).toHaveCount(0)
    await editor.screenshot({ path: test.info().outputPath('after.png') })
  } finally {
    await closeAndSaveVideo(launched, 'ctx-word-items')
  }
})
