/**
 * Insert > Text Box > Draw Vertical Text Box: a click drops a tall box whose
 * text runs down it, the file says so, and the box reads back vertical.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

test('a vertical text box is drawn from the menu, typed into, saved and read back', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-vertical-box-')))
  const docPath = join(dir, 'vertical.docx')
  writeFileSync(docPath, await docxWith(para('First') + para('Last')))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-vertical-text-box',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const docPage = page.locator('.doc-page[contenteditable="true"]')
    await expect(page.locator('.doc-page p', { hasText: 'First' })).toBeVisible({
      timeout: 30_000,
    })
    await app.evaluate(({ Menu }) => {
      const insert = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Insert')
      const textBox = insert?.submenu?.items.find((item) => item.label === 'Text Box')
      textBox?.submenu?.items.find((item) => item.label === 'Draw Vertical Text Box')?.click()
    })
    await expect(docPage).toHaveClass(/doc-shape-drawing/)
    const paper = (await docPage.boundingBox())!
    await page.mouse.click(paper.x + 300, paper.y + 260)

    const box = docPage.locator('.doc-textbox')
    await expect(box).toHaveCount(1)
    await expect(box).toHaveClass(/doc-textbox-vert/)
    // a click drops the default box: taller than it is wide
    const size = (await box.boundingBox())!
    expect(size.height).toBeGreaterThan(size.width)
    // the box takes the focus a moment after it is dropped: keys typed before that go to the page
    await expect
      .poll(() => page.evaluate(() => !!document.activeElement?.closest('.doc-textbox-editor')))
      .toBe(true)
    await page.keyboard.type('Down the box')
    await expect(box).toContainText('Down the box')
    // the text runs down: its line is taller than it is wide
    const line = await box.evaluate((el) => {
      const p = el.querySelector('p') ?? el
      const r = p.getBoundingClientRect()
      return { mode: getComputedStyle(p).writingMode, tall: r.height > r.width }
    })
    expect(line).toEqual({ mode: 'vertical-rl', tall: true })
    await page.screenshot({ path: test.info().outputPath('vertical-box.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
          return xml.includes('<wps:bodyPr vert="eaVert"') && xml.includes('Down the box')
        },
        { timeout: 30_000 },
      )
      .toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'docs-vertical-text-box')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-vertical-text-box',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    const box = page.locator('.doc-page .doc-textbox')
    await expect(box).toHaveCount(1, { timeout: 30_000 })
    await expect(box).toHaveClass(/doc-textbox-eavert/)
    await expect(box).toContainText('Down the box')
  } finally {
    await closeAndSaveVideo(launched, 'docs-vertical-text-box')
    rmSync(dir, { recursive: true, force: true })
  }
})
