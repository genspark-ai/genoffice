import { test, expect } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const styled = (style: string, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

test('Design > style sets: a card gives the document that look, on the page and in the saved file', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-style-sets-')))
  const docPath = join(dir, 'report.docx')
  writeFileSync(
    docPath,
    await docxWith(
      styled('Title', 'Annual report') +
        styled('Heading1', 'Overview') +
        ['apple', 'banana'].map(para).join(''),
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-design-style-sets',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const heading = page.locator('.doc-page :is(h1, p)', { hasText: 'Overview' }).first()
    await expect(heading).toBeVisible({ timeout: 30_000 })
    await page.locator('.doc-page p', { hasText: 'apple' }).first().click()

    await page.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
    const strip = page.locator('.style-set-strip')
    // the document as it stands comes first, then Word's sets
    await expect(strip.locator('.style-set-card').first()).toHaveAttribute(
      'data-testid',
      'style-set-current',
    )
    await expect(strip.getByTestId('style-set-distinctive')).toBeVisible()
    await page.screenshot({ path: test.info().outputPath('style-sets-before.png') })

    await page.getByTestId('style-set-more').click()
    const menu = page.locator('.style-set-menu')
    await expect(menu.locator('.style-set-card')).toHaveCount(13)
    await page.screenshot({ path: test.info().outputPath('style-sets-menu.png') })
    await menu.getByTestId('style-set-modern').click()
    await expect(menu).toBeHidden()

    // Modern: headings are white capitals on a bar of the first accent colour
    await expect
      .poll(() => heading.evaluate((el) => getComputedStyle(el).textTransform))
      .toBe('uppercase')
    await expect
      .poll(() => heading.evaluate((el) => getComputedStyle(el).backgroundColor))
      .not.toBe('rgba(0, 0, 0, 0)')
    await page.screenshot({ path: test.info().outputPath('style-sets-after.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => {
        const zip = await JSZip.loadAsync(readFileSync(docPath))
        const xml = (await zip.file('word/styles.xml')?.async('string')) ?? ''
        const style = /<w:style\b[^>]*w:styleId="Heading1"[\s\S]*?<\/w:style>/.exec(xml)?.[0] ?? ''
        return [/<w:caps\/>/.test(style), /<w:shd\b[^>]*w:themeFill="accent1"/.test(style)]
      })
      .toEqual([true, true])
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})
