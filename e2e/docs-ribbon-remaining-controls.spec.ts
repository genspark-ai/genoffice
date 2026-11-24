/**
 * Ribbon controls Word has at fixed places: References > Mark Entry and Insert
 * Index, Review > Display for Review, Layout > the labelled spacing boxes, and
 * Mailings > Record, Merge Range with From and To. Update Labels
 * is greyed out while the document is not a sheet of labels.
 */
import { test, expect } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-ribbon-remaining-controls'
const REV = 'w:author="Ann" w:date="2026-09-01T10:00:00Z"'

const documentXml = async (file: string) =>
  (await (await JSZip.loadAsync(readFileSync(file))).file('word/document.xml')?.async('string')) ??
  ''

test('Mark Entry, Insert Index, Display for Review and the spacing boxes change the page and the saved file', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ribbon-rest-')))
  const docPath = join(dir, 'report.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('Granite') +
        `<w:p><w:r><w:t xml:space="preserve">Beta keeps </w:t></w:r><w:ins w:id="1" ${REV}><w:r><w:t>inserted words</w:t></w:r></w:ins><w:del w:id="2" ${REV}><w:r><w:delText>removed words</w:delText></w:r></w:del></w:p>` +
        para('Closing line'),
    ),
  )
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p', { hasText: 'Granite' }).first()
    await expect(first).toBeVisible({ timeout: 30_000 })
    const tab = (name: string) =>
      page.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).click()
    const big = (name: string) =>
      page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })

    // Review: the view is named on the control and picked from its list
    await tab('Review')
    const root = page.locator('.app')
    const display = page.getByRole('button', { name: 'Display for Review', exact: true })
    const removed = page.locator('.doc-del', { hasText: 'removed words' })
    await expect(display).toHaveText('All Markup')
    await expect(removed).toBeVisible()
    await display.click()
    await expect(page.locator('.layout-menu button')).toHaveText([
      'Simple Markup',
      'All Markup ✓',
      'No Markup',
      'Original',
    ])
    await page.locator('.layout-menu button', { hasText: 'No Markup' }).click()
    await expect(root).toHaveClass(/rev-display-none/)
    await expect(display).toHaveText('No Markup')
    await expect(removed).toBeHidden()
    await display.click()
    await page.locator('.layout-menu button', { hasText: 'Original' }).click()
    await expect(root).toHaveClass(/rev-display-original/)
    await expect(page.locator('.doc-ins', { hasText: 'inserted words' })).toBeHidden()
    await page.screenshot({ path: test.info().outputPath('display-for-review.png') })
    await display.click()
    await page.locator('.layout-menu button', { hasText: 'All Markup' }).click()
    await expect(root).not.toHaveClass(/rev-display-/)

    // References: Mark Entry takes the selected word, Insert Index lists it
    await first.dblclick()
    await tab('References')
    await expect(big('Style')).toBeEnabled()
    await big('Mark Entry').click()
    const prompt = page.locator('.modal', { hasText: 'Mark Index Entry' })
    await expect(prompt.locator('input[name="main"]')).toHaveValue('Granite')
    await prompt.locator('input[name="main"]').press('Enter')
    await expect(prompt).toBeHidden()
    await expect(page.locator('.doc-page .doc-xe-mark')).toHaveAttribute('data-xe-term', 'Granite')
    await big('Insert Index').click()
    await expect
      .poll(() => page.locator('.doc-page').first().innerText())
      .toMatch(/Granite[\s\S]+Granite, 1/)
    await page.screenshot({ path: test.info().outputPath('references.png') })

    // Layout: the spacing boxes answer to their names
    await page.locator('.doc-page p', { hasText: 'Beta keeps' }).first().click()
    await tab('Layout')
    /** the gap between the first paragraph and the one below it, in px */
    const gap = () =>
      page.evaluate(() => {
        const [a, b] = ['Granite', 'Beta keeps'].map((text) =>
          [...document.querySelectorAll('.doc-page p')].find((p) => p.textContent?.includes(text))!,
        )
        return b.getBoundingClientRect().top - a.getBoundingClientRect().bottom
      })
    const gapBefore = await gap()
    const before = page.getByRole('spinbutton', { name: 'Space Before', exact: true })
    const after = page.getByRole('spinbutton', { name: 'Space After', exact: true })
    await before.fill('18')
    await before.press('Enter')
    // Enter hands the focus back to the page a frame later, and would take it from the next box mid-entry
    await expect(page.locator('.ProseMirror.doc-page').first()).toBeFocused()
    await after.fill('6')
    await after.press('Enter')
    await expect.poll(async () => (await gap()) - gapBefore).toBeGreaterThan(10)
    await page.screenshot({ path: test.info().outputPath('layout.png') })

    // the index's section was saved when it went in: wait for this save by what was changed since
    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 20_000 })
      .toMatch(/<w:spacing[^>]*w:before="360"[^>]*w:after="120"/)
    const saved = await documentXml(docPath)
    expect(saved).toMatch(/XE\s+"Granite"/)
    expect(saved).toMatch(/<w:instrText[^>]*>\s*INDEX\b/)
    // the view is not part of the document: both changes are still tracked
    expect(saved).toContain('<w:ins ')
    expect(saved).toContain('<w:del ')

    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(reopened.locator('.doc-page .doc-xe-mark')).toHaveAttribute(
      'data-xe-term',
      'Granite',
      { timeout: 30_000 },
    )
    await reopened.locator('.doc-page p', { hasText: 'Beta keeps' }).first().click()
    await reopened.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await expect(
      reopened.getByRole('spinbutton', { name: 'Space Before', exact: true }),
    ).toHaveValue('18')
    await expect(
      reopened.getByRole('spinbutton', { name: 'Space After', exact: true }),
    ).toHaveValue('6')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Record moves the preview; Merge Range with From and To decides what Finish & Merge writes', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-merge-range-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(para('Dear') + para('Regards')))
  const listPath = join(dir, 'people.csv')
  writeFileSync(listPath, 'Name\nAda\nGrace\nAlan\nLinus\nBo\n')
  const outPath = join(dir, 'merged.docx')
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const greeting = page.locator('.doc-page p', { hasText: 'Dear' }).first()
    await expect(greeting).toBeVisible({ timeout: 30_000 })
    await app.evaluate(
      ({ dialog }, [list, out]) => {
        dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [list] })
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
      },
      [listPath, outPath],
    )
    await page.locator('.ribbon-tab', { hasText: /^Mailings$/ }).click()
    const big = (name: string) =>
      page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })
    const box = (name: string) => page.getByRole('spinbutton', { name, exact: true })
    const recordStatus = page.locator('[data-ribbon-body] [data-testid="mailings-record-status"]')
    await expect(recordStatus).toHaveText('No data source selected')
    await expect(big('Update Labels')).toBeDisabled()
    await expect(box('Record')).toBeDisabled()
    await big('Select Recipients').click()
    await expect(recordStatus).toHaveText('Record 1 of 5')
    await expect(recordStatus).toBeVisible()
    await expect(box('Record')).toHaveValue('1')

    await greeting.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ')
    await big('Insert Merge Field').click()
    await page.locator('.rb-field-item', { hasText: /^Name$/ }).click()
    await big('Preview Results').click()
    const preview = page.locator('.doc-page .doc-merge-preview')
    await expect(preview).toHaveText('Ada')

    // Record: a number typed goes to that record, and the buttons keep the box in step
    await box('Record').fill('3')
    await box('Record').press('Enter')
    await expect(preview).toHaveText('Alan')
    await expect(recordStatus).toHaveText('Record 3 of 5')
    await page.locator('[data-ribbon-body] button[aria-label="Next"]').click()
    await expect(box('Record')).toHaveValue('4')
    await expect(preview).toHaveText('Linus')
    await box('Record').fill('99')
    await box('Record').press('Enter')
    await expect(box('Record')).toHaveValue('5')
    await expect(preview).toHaveText('Bo')

    const range = page.getByRole('button', { name: 'Merge Range', exact: true })
    const panel = page.locator('.modal-mailmerge')
    const names = async () => (await documentXml(outPath)).match(/Ada|Grace|Alan|Linus|Bo\b/g) ?? []
    const finish = async () => {
      await big('Finish & Merge').click()
      await panel.locator('input[name="mailmerge-output"]').nth(1).check()
      await panel.locator('button.btn-primary').click()
      await expect(panel.locator('.mailmerge-done')).toBeVisible({ timeout: 30_000 })
      await panel.getByRole('button', { name: 'Close' }).click()
      await expect(panel).toBeHidden()
    }

    // All: From and To stay greyed out
    await expect(range).toHaveText('All')
    await expect(box('From')).toBeDisabled()
    await finish()
    expect(await names()).toEqual(['Ada', 'Grace', 'Alan', 'Linus', 'Bo'])

    // Current Record: the one the preview shows
    await range.click()
    await page.locator('.layout-menu button', { hasText: 'Current Record' }).click()
    await expect(range).toHaveText('Current Record')
    await finish()
    expect(await names()).toEqual(['Bo'])

    // Custom: From through To
    await range.click()
    await page.locator('.layout-menu button', { hasText: 'Custom' }).click()
    await expect(box('From')).toBeEnabled()
    await expect(box('To')).toHaveValue('5')
    await box('From').fill('2')
    await box('From').press('Enter')
    await box('To').fill('4')
    await box('To').press('Enter')
    await page.screenshot({ path: test.info().outputPath('mailings.png') })
    await finish()
    expect(await names()).toEqual(['Grace', 'Alan', 'Linus'])

    // a range that holds no record has nothing to merge
    await box('From').fill('5')
    await box('From').press('Enter')
    await big('Finish & Merge').click()
    await expect(panel.locator('button.btn-primary')).toBeDisabled()
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
