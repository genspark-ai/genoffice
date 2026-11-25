/**
 * The red lines under misspelled words are drawn by the page, so the words a
 * file came with have one without anyone typing next to them. They are a CSS
 * highlight: the test reads which words it covers.
 */
import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-spelling-marks'
const POLL = { timeout: 20_000, intervals: [250, 500, 1000] }

/** the words under a red line, in the order the page drew them */
const marked = (page: Page) =>
  page.evaluate(() => {
    const highlight = (CSS as unknown as { highlights: Map<string, Set<Range>> }).highlights.get(
      'docs-spelling',
    )
    return [...(highlight ?? [])].map((range) => range.toString())
  })

test('misspelled words of an opened file are marked; typing, Ignore Once and the Spelling button follow', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-spelling-marks-')))
  const docPath = join(dir, 'marks.docx')
  writeFileSync(
    docPath,
    await docxWith(para('First the mispeled word and plmokn here.') + para('A clean line')),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'mispeled' })).toBeVisible({
      timeout: 30_000,
    })
    const known = await page.evaluate(() => {
      const api = (window as unknown as { desktop: { spellIsMisspelled?(w: string): boolean } })
        .desktop
      return api.spellIsMisspelled?.('mispeled') === true && !api.spellIsMisspelled?.('clean')
    })
    test.skip(!known, 'no dictionary in this environment')

    // nothing was typed: the words came with the file
    await expect.poll(() => marked(page), POLL).toEqual(['mispeled', 'plmokn'])
    await page.screenshot({ path: test.info().outputPath('marks-on-open.png') })

    // the word being typed waits for the caret to leave it
    await page.evaluate(() => {
      const { editor } = (
        window as unknown as { __aidocs: { editor: import('@tiptap/core').Editor } }
      ).__aidocs
      editor
        .chain()
        .focus()
        .setTextSelection(editor.state.doc.content.size - 1)
        .run()
    })
    await page.keyboard.type(' wrnog', { delay: 20 })
    await expect(page.locator('.doc-page p', { hasText: 'A clean line wrnog' })).toBeVisible()
    expect(await marked(page)).toEqual(['mispeled', 'plmokn'])
    await page.keyboard.type(' ', { delay: 20 })
    await expect.poll(() => marked(page), POLL).toEqual(['mispeled', 'plmokn', 'wrnog'])

    // Ignore Once takes the line off that word
    await page.locator('.doc-page p', { hasText: 'plmokn' }).evaluate((p) => {
      const node = [...p.querySelectorAll('*'), p]
        .flatMap((el) => [...el.childNodes])
        .find((n) => n.nodeType === 3 && n.textContent!.includes('plmokn'))!
      const range = document.createRange()
      const at = node.textContent!.indexOf('plmokn')
      range.setStart(node, at)
      range.setEnd(node, at + 6)
      const box = range.getBoundingClientRect()
      ;(window as unknown as { __at: { x: number; y: number } }).__at = {
        x: box.left + box.width / 2,
        y: box.top + box.height / 2,
      }
    })
    const at = await page.evaluate(
      () => (window as unknown as { __at: { x: number; y: number } }).__at,
    )
    await page.mouse.click(at.x, at.y, { button: 'right' })
    await page
      .locator('.ctx-menu .ctx-item', {
        has: page.locator('.ctx-label', { hasText: /^Ignore Once$/ }),
      })
      .click()
    await expect.poll(() => marked(page), POLL).toEqual(['mispeled', 'wrnog'])

    // Review > Spelling off: no lines; on again: they are back with no typing
    await page.getByRole('button', { name: 'Review' }).click()
    const spelling = page.getByRole('button', { name: 'Spelling' })
    await spelling.click()
    await expect.poll(() => marked(page), POLL).toEqual([])
    await spelling.click()
    await expect.poll(() => marked(page), POLL).toEqual(['mispeled', 'wrnog'])
    await page.screenshot({ path: test.info().outputPath('marks-back-on.png') })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
