/**
 * Navigation pane (Word for Mac model): heading tree with caret highlight,
 * click moves the caret, collapse, search results with document highlights,
 * per-heading hit counts, page thumbnails, heading context menu, drag reorder,
 * resizable width.
 */
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const MAC = process.platform === 'darwin'
const POLL = { timeout: 15_000, intervals: [250, 500, 1000] }

test('headings tree, search results, pages and heading edits', async () => {
  test.setTimeout(150_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-nav-pane' })
  const { app, page } = launched
  try {
    await page.locator('.quick-card').first().click()
    const editor = await waitForPageWithUrl(app, '://docs/')
    const docPage = editor.locator('.doc-page[contenteditable="true"]')
    await docPage.waitFor()
    await docPage.click()
    const headingKey = (n: number) => (MAC ? `Alt+Meta+${n}` : `Control+Alt+${n}`)
    await editor.keyboard.type('Intro', { delay: 10 })
    await editor.keyboard.press(headingKey(1))
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('intro body alpha', { delay: 10 })
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('Scope', { delay: 10 })
    await editor.keyboard.press(headingKey(2))
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('scope body alpha alpha', { delay: 10 })
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('Methods', { delay: 10 })
    await editor.keyboard.press(headingKey(1))
    await editor.keyboard.press('Enter')
    await editor.keyboard.type('methods body', { delay: 10 })
    await expect(docPage.locator('h1')).toHaveCount(2)
    await expect(docPage.locator('h2')).toHaveCount(1)

    await editor.locator('.ribbon-tab', { hasText: /^View$/ }).click()
    await editor.locator('.rb-big', { hasText: /^Navigation Pane$/ }).click()
    const pane = editor.locator('.nav-pane')
    await pane.waitFor()
    const rows = pane.locator('.nav-row')
    await expect(rows).toHaveCount(3)
    await expect(rows.locator('.nav-row-text')).toHaveText(['Intro', 'Scope', 'Methods'])
    const padding = (i: number) =>
      rows.nth(i).evaluate((el) => parseFloat(getComputedStyle(el).paddingLeft))
    expect(await padding(1)).toBeGreaterThan(await padding(0))

    // caret is at the end of the document: the last heading is current
    await expect(pane.locator('.nav-row.on .nav-row-text')).toHaveText('Methods')

    const caretBlock = () =>
      editor.evaluate(() => {
        const ed = (window as unknown as { __aidocs: { editor: import('@tiptap/core').Editor } })
          .__aidocs.editor
        const $p = ed.state.doc.resolve(ed.state.selection.from)
        return `${$p.parent.type.name}:${$p.parent.textContent}:${$p.parentOffset}`
      })

    // ---- click moves the caret to the heading start ----
    await rows.nth(0).click()
    await expect.poll(caretBlock, POLL).toBe('docHeading:Intro:0')
    await expect(pane.locator('.nav-row.on .nav-row-text')).toHaveText('Intro')
    // caret inside a section highlights its heading
    await docPage.locator('p', { hasText: 'scope body' }).click()
    await expect(pane.locator('.nav-row.on .nav-row-text')).toHaveText('Scope')

    // ---- collapse / expand ----
    await rows.nth(0).locator('.nav-twisty').click()
    await expect(rows.locator('.nav-row-text')).toHaveText(['Intro', 'Methods'])
    // the caret is still in Scope: its collapsed parent carries the highlight
    await expect(pane.locator('.nav-row.on .nav-row-text')).toHaveText('Intro')
    await rows.nth(0).locator('.nav-twisty').click()
    await expect(rows).toHaveCount(3)

    // ---- search: results tab, document highlights, next/previous ----
    const search = pane.locator('.nav-search-input')
    await search.fill('alpha')
    await expect(pane.locator('.nav-tab.on')).toHaveText('Results')
    await expect(pane.locator('.nav-result')).toHaveCount(3, POLL)
    await expect(docPage.locator('.search-hit')).toHaveCount(3)
    await expect(pane.locator('.nav-search-count')).toHaveText('1/3')
    await search.press('Enter')
    await expect(pane.locator('.nav-search-count')).toHaveText('2/3')
    await expect(pane.locator('.nav-result.on mark')).toHaveText('alpha')
    await search.press('Shift+Enter')
    await expect(pane.locator('.nav-search-count')).toHaveText('1/3')
    await pane.locator('.nav-result').nth(2).click()
    await expect(pane.locator('.nav-search-count')).toHaveText('3/3')
    await expect
      .poll(
        () =>
          editor.evaluate(() => {
            const ed = (
              window as unknown as { __aidocs: { editor: import('@tiptap/core').Editor } }
            ).__aidocs.editor
            const { from, to } = ed.state.selection
            return ed.state.doc.textBetween(from, to)
          }),
        POLL,
      )
      .toBe('alpha')

    // headings tab: hit counts roll up, headings without hits dim
    await pane.locator('.nav-tab', { hasText: 'Headings' }).click()
    await expect(rows.nth(0).locator('.nav-badge')).toHaveText('3')
    await expect(rows.nth(1).locator('.nav-badge')).toHaveText('2')
    await expect(rows.nth(2)).toHaveClass(/dim/)
    await expect(rows.nth(0)).not.toHaveClass(/dim/)

    // pages tab filters to pages with hits; clearing the search shows all pages
    await pane.locator('.nav-tab', { hasText: 'Pages' }).click()
    await expect(pane.locator('.nav-thumb')).toHaveCount(1, POLL)
    await expect(pane.locator('.nav-thumb.hit')).toHaveCount(1)
    await search.fill('')
    await expect(docPage.locator('.search-hit')).toHaveCount(0)
    // clearing returns to the tab that was active before typing
    await expect(pane.locator('.nav-tab.on')).toHaveText('Headings')
    await pane.locator('.nav-tab', { hasText: 'Pages' }).click()
    await expect(pane.locator('.nav-thumb')).toHaveCount(1, POLL)
    await expect(pane.locator('.nav-thumb.on')).toHaveCount(1)
    await expect(pane.locator('.nav-thumb-page .nav-thumb-slot')).not.toHaveCount(0, POLL)
    await docPage.locator('p', { hasText: 'methods body' }).click()
    await pane.locator('.nav-thumb').first().click()
    await expect.poll(caretBlock, POLL).toBe('docHeading:Intro:0')

    // ---- context menu: promote, then delete heading with content ----
    await pane.locator('.nav-tab', { hasText: 'Headings' }).click()
    await rows.nth(1).click({ button: 'right' })
    const menu = editor.locator('.nav-menu')
    await menu.waitFor()
    await menu.locator('.ctx-item', { hasText: /^Promote$/ }).click()
    await expect(docPage.locator('h1')).toHaveCount(3)
    expect(await padding(1)).toBe(await padding(0))
    await rows.nth(1).click({ button: 'right' })
    await menu.locator('.ctx-item', { hasText: /^Select Heading and Content$/ }).click()
    await expect
      .poll(
        () =>
          editor.evaluate(() => {
            const ed = (
              window as unknown as { __aidocs: { editor: import('@tiptap/core').Editor } }
            ).__aidocs.editor
            const { from, to } = ed.state.selection
            return ed.state.doc.textBetween(from, to, '|')
          }),
        POLL,
      )
      .toBe('Scope|scope body alpha alpha')

    // ---- drag Methods (with its body) before Intro ----
    const dt = await editor.evaluateHandle(() => new DataTransfer())
    await rows.nth(2).dispatchEvent('dragstart', { dataTransfer: dt })
    const introBox = await rows.nth(0).boundingBox()
    if (!introBox) throw new Error('no heading row box')
    await rows.nth(0).dispatchEvent('dragover', {
      dataTransfer: dt,
      clientX: introBox.x + 20,
      clientY: introBox.y + 2,
    })
    await expect(rows.nth(0)).toHaveClass(/drop-before/)
    await rows.nth(0).dispatchEvent('drop', { dataTransfer: dt })
    await rows.nth(0).dispatchEvent('dragend')
    await expect(rows.locator('.nav-row-text')).toHaveText(['Methods', 'Intro', 'Scope'])
    await expect
      .poll(
        () =>
          docPage.evaluate((el) =>
            Array.from(el.querySelectorAll('h1, h2, p')).map((b) => b.textContent),
          ),
        POLL,
      )
      .toEqual([
        'Methods',
        'methods body',
        'Intro',
        'intro body alpha',
        'Scope',
        'scope body alpha alpha',
      ])
    // one undo step brings the whole subtree back
    await docPage.click()
    await editor.keyboard.press('ControlOrMeta+z')
    await expect(rows.locator('.nav-row-text')).toHaveText(['Intro', 'Scope', 'Methods'])

    await rows.nth(2).click({ button: 'right' })
    await menu.locator('.ctx-item', { hasText: /^Delete$/ }).click()
    await expect(rows).toHaveCount(2)
    await expect(docPage).not.toContainText('methods body')

    // New Heading After: the blank heading is listed at once, caret inside it
    await rows.nth(1).click({ button: 'right' })
    await menu.locator('.ctx-item', { hasText: /^New Heading After$/ }).click()
    await expect(rows).toHaveCount(3)
    await expect.poll(caretBlock, POLL).toBe('docHeading::0')
    await editor.keyboard.type('Extra', { delay: 10 })
    await expect(rows.locator('.nav-row-text')).toHaveText(['Intro', 'Scope', 'Extra'])

    // ---- resizable, persisted width; close button ----
    const before = await pane.evaluate((el) => el.getBoundingClientRect())
    const handle = pane.locator('.nav-pane-resizer')
    const hb = await handle.boundingBox()
    if (!hb) throw new Error('no resizer')
    await editor.mouse.move(hb.x + hb.width / 2, hb.y + 100)
    await editor.mouse.down()
    await editor.mouse.move(hb.x + hb.width / 2 + 80, hb.y + 100, { steps: 4 })
    await editor.mouse.up()
    await expect
      .poll(() => pane.evaluate((el) => el.getBoundingClientRect().width))
      .toBeCloseTo(before.width + 80, 0)
    expect(await editor.evaluate(() => localStorage.getItem('aidocs.navWidth'))).toBe(
      String(Math.round(before.width + 80)),
    )
    await pane.locator('.nav-pane-close').click()
    await expect(pane).toHaveCount(0)
    expect(await editor.evaluate(() => localStorage.getItem('aidocs.showNav'))).toBe('0')
  } finally {
    await closeAndSaveVideo(launched, 'docs-nav-pane')
  }
})
