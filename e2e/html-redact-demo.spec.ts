import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

/**
 * Withholding a value from the model, driven through the real app.
 *
 * The entry point is the **source view**, deliberately: a rendered page puts no
 * caret in a `<script>` string, a `content=` value or a comment, which is where
 * an HTML page's secrets usually are. The source pane shows all of it and lets
 * every one of them be selected.
 *
 * What is asserted afterwards is the contract, not the pixels: the file still
 * holds the real value (the page keeps working) and now carries the mark, so
 * the model's view of it is the placeholder and nothing else.
 */

const LABEL = '客户电话'
const TEXT_SECRET = '13800138000'
const ATTR_SECRET = 'sk-ATTR-1111'
const SCRIPT_SECRET = 'sk-SCRIPT-2222'

const SOURCE = [
  '<!doctype html>',
  '<html>',
  '<head>',
  '  <meta name="api-key" content="' + ATTR_SECRET + '">',
  '  <script>',
  '    const KEY = "' + SCRIPT_SECRET + '";',
  '    render(KEY);',
  '  </script>',
  '</head>',
  '<body>',
  '  <p>Call ' + TEXT_SECRET + ' now to confirm the order.</p>',
  '  <p>Ordinary paragraph.</p>',
  '</body>',
  '</html>',
  '',
].join('\n')

/** double-click is the gesture a reader would use: it selects the word under the caret */
async function selectWord(page: import('@playwright/test').Page, word: string) {
  const box = await page.evaluate((w) => {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT)
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const at = n.textContent?.indexOf(w) ?? -1
      if (at === -1) continue
      const range = document.createRange()
      range.setStart(n, at)
      range.setEnd(n, at + w.length)
      const r = range.getBoundingClientRect()
      if (r.width === 0) continue
      return { x: r.left + r.width / 2, y: r.top + r.height / 2 }
    }
    return null
  }, word)
  expect(box, `word not found on screen: ${word}`).not.toBeNull()
  await page.mouse.dblclick(box!.x, box!.y)
}

/** right-click the middle of the current selection, which is what opens the menu */
async function rightClickSelection(page: import('@playwright/test').Page) {
  const [x, y] = await rightClickPoint(page)
  await page.mouse.click(x, y, { button: 'right' })
  await expect(page.locator('[role="menu"]')).toBeVisible({ timeout: 15_000 })
}

async function hideSelection(page: import('@playwright/test').Page, label: string) {
  await page.locator('[role="menuitem"]').first().click()
  const dialog = page.locator('.redact-dialog')
  await expect(dialog).toBeVisible()
  const input = dialog.locator('input').first()
  await input.fill(label)
  // the marker the model will be shown, live
  await expect(dialog).toContainText(`{{${label}}}`)
  await dialog.locator('button').last().click()
  await expect(dialog).toHaveCount(0)
}

test('a reader can withhold a value in the source view without changing the page', async () => {
  test.setTimeout(180_000)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-redact-html-'))
  const htmlPath = join(dir, 'page.html')
  await writeFile(htmlPath, SOURCE)

  const launched = await launchShell({
    onboardingSeen: true,
    settings: { lang: 'en' },
    videoDir: 'redact-html',
    openFile: htmlPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://html/')
    await expect(page.locator('.preview-frame')).toBeVisible()

    // the preview is where the page is seen; the source pane is where it is kept
    await page.locator('.rb-view', { hasText: /Source/ }).click()
    const cm = page.locator('.source-editor .cm-content')
    await expect(cm).toBeVisible()
    await page.screenshot({ path: screenshotPath('html-1-source-view') })

    // ── visible text ──
    await selectWord(page, TEXT_SECRET)
    await rightClickSelection(page)
    await page.screenshot({ path: screenshotPath('html-2-context-menu') })
    await hideSelection(page, LABEL)
    await page.screenshot({ path: screenshotPath('html-3-marked') })

    // ── a value inside a <script> ──
    await selectWord(page, SCRIPT_SECRET)
    await rightClickSelection(page)
    await hideSelection(page, 'inline key')
    await page.screenshot({ path: screenshotPath('html-4-script-marked') })

    // ── an attribute value ──
    await selectWord(page, ATTR_SECRET)
    await rightClickSelection(page)
    await hideSelection(page, 'meta key')

    // save the way the reader would: the file already has a path, so ⌘/Ctrl+S
    // writes in place through the same serializer the app uses
    await expect(page.locator('.status-save')).toHaveText(/Unsaved/)
    await cm.click() // the dialog took the focus; give the shortcut its window back
    await page.keyboard.press('ControlOrMeta+s')
    await expect(page.locator('.status-save')).toHaveText(/Saved/)

    const after = await readFile(htmlPath, 'utf8')
    // the marks are in the file…
    expect(after).toContain(`data-gx-redact="${LABEL}"`)
    expect(after).toContain('/*gx:redact:')
    expect(after).toMatch(/data-gx-redact-content="/)
    // …and so are the original values, because a mark never removes anything
    expect(after).toContain(TEXT_SECRET)
    expect(after).toContain(ATTR_SECRET)
    expect(after).toContain(SCRIPT_SECRET)
    // no placeholder leaked into the page the reader ships
    expect(after).not.toContain('{{')
    await page.screenshot({ path: screenshotPath('html-5-saved') })
  } finally {
    await closeAndSaveVideo(launched, 'redact-html')
  }
})

/** the point to right-click: the middle of the current selection */
async function rightClickPoint(page: import('@playwright/test').Page): Promise<[number, number]> {
  return page.evaluate(() => {
    const sel = window.getSelection()
    if (!sel || sel.rangeCount === 0) throw new Error('nothing is selected')
    const r = sel.getRangeAt(0).getBoundingClientRect()
    return [r.left + r.width / 2, r.top + r.height / 2] as [number, number]
  })
}
