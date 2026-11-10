/**
 * Word's Reviewing Pane: Review → Reviewing Pane docks right (independent
 * of Navigation on the left), sums up Insertions / Deletions / Moves /
 * Formatting / Comments and lists every tracked change with author and
 * excerpt; Accept on an entry removes it from the list and from the
 * document. Reviewers after the first get their own ink.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const ANN = 'w:author="Ann" w:date="2026-09-01T10:00:00Z"'
const BEN = 'w:author="Ben" w:date="2026-09-02T10:00:00Z"'

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`

async function trackedDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>` +
      p('Alpha opens the document.') +
      `<w:p><w:r><w:t xml:space="preserve">Beta keeps </w:t></w:r><w:ins w:id="1" ${ANN}><w:r><w:t>inserted words</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> at the end.</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t xml:space="preserve">Gamma keeps </w:t></w:r><w:del w:id="2" ${ANN}><w:r><w:delText>removed words</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> around.</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t xml:space="preserve">Delta has </w:t></w:r><w:ins w:id="3" ${BEN}><w:r><w:t>second reviewer</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> text.</w:t></w:r></w:p>` +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test.describe('docs Reviewing Pane', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-review-pane-')))
    docPath = join(dir, 'tracked.docx')
    writeFileSync(docPath, await trackedDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('lists every change, accepts one from the list, inks the second reviewer', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-reviewing-pane',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.locator('.doc-page p', { hasText: 'Alpha opens' }).first().waitFor()

      await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
      await page.getByRole('button', { name: 'Reviewing Pane', exact: true }).click()
      const pane = page.locator('.review-pane')
      await expect(pane).toBeVisible()
      // The Reviewing Pane occupies the workspace's outermost right edge.
      const paneBox = (await pane.boundingBox())!
      const scrollBox = (await page.locator('.editor-scroll').boundingBox())!
      const workspaceBox = (await pane.locator('..').boundingBox())!
      expect(paneBox.x).toBeGreaterThanOrEqual(scrollBox.x + scrollBox.width - 1)
      expect(
        Math.abs(paneBox.x + paneBox.width - workspaceBox.x - workspaceBox.width),
      ).toBeLessThanOrEqual(1)
      const cards = pane.locator('.review-row')
      await expect(cards).toHaveCount(3)
      await expect(pane.locator('.review-pane-count')).toHaveText('3')
      await expect(pane.locator('.review-pane-sum-line')).toHaveText([
        'Insertions:2',
        'Deletions:1',
        'Moves:0',
        'Formatting:0',
        'Comments:0',
      ])
      await expect(pane.locator('.review-pane-author')).toHaveText(['Ann', 'Ben'])
      await expect(cards.nth(0)).toContainText('inserted words')
      await expect(cards.nth(0).locator('.review-row-head')).toContainText('Inserted')
      await expect(cards.nth(0).locator('.review-row-head')).toContainText('Ann')
      await expect(cards.nth(1)).toContainText('removed words')
      await expect(cards.nth(2)).toContainText('second reviewer')

      // the second reviewer's insertion does not share the first reviewer's ink
      await page.getByRole('button', { name: 'Markup', exact: true }).click()
      await page.locator('.layout-menu button', { hasText: 'All Markup' }).click()
      const annInk = await page
        .locator('.doc-ins', { hasText: 'inserted words' })
        .evaluate((el) => getComputedStyle(el).color)
      const benInk = await page
        .locator('.doc-ins', { hasText: 'second reviewer' })
        .evaluate((el) => getComputedStyle(el).color)
      expect(benInk).not.toBe(annInk)

      // a card click selects the change; Accept removes it from list and body
      await cards.nth(1).click()
      await expect(cards.nth(1)).toHaveClass(/active/)
      await cards
        .nth(1)
        .getByRole('button', { name: /Accept/ })
        .click()
      await expect(cards).toHaveCount(2)
      await expect(page.locator('.doc-del', { hasText: 'removed words' })).toHaveCount(0)
      await expect(pane.locator('.review-pane-count')).toHaveText('2')
      await expect(pane.locator('.review-pane-sum-line').nth(1)).toHaveText('Deletions:0')

      // Specific People: unticking Ben hides his change from view and pane
      await page.getByRole('button', { name: 'Markup', exact: true }).click()
      await page.locator('.layout-menu button', { hasText: /^Ben/ }).click()
      await expect(cards).toHaveCount(1)
      await expect(pane.locator('.review-pane-hidden')).toHaveText('1 hidden by filters')
      await expect(page.locator('.doc-ins', { hasText: 'second reviewer' })).toHaveCSS(
        'text-decoration-line',
        'none',
      )

      // Navigation stays on the left; the right Reviewing Pane toggles independently.
      await page.locator('.ribbon-tab', { hasText: /^View$/ }).click()
      await page.locator('.rb-big', { hasText: /^Navigation Pane$/ }).click()
      const navigation = page.locator('.nav-pane:not(.review-pane)')
      await expect(navigation).toBeVisible()
      await expect(pane).toBeVisible()
      const navigationBox = (await navigation.boundingBox())!
      const sharedScrollBox = (await page.locator('.editor-scroll').boundingBox())!
      const sharedPaneBox = (await pane.boundingBox())!
      expect(navigationBox.x + navigationBox.width).toBeLessThanOrEqual(sharedScrollBox.x + 1)
      expect(sharedPaneBox.x).toBeGreaterThanOrEqual(sharedScrollBox.x + sharedScrollBox.width - 1)
      await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
      await page.getByRole('button', { name: 'Reviewing Pane', exact: true }).click()
      await expect(pane).toHaveCount(0)
      await expect(navigation).toBeVisible()
      await page.getByRole('button', { name: 'Reviewing Pane', exact: true }).click()
      await expect(pane).toBeVisible()
      await expect(navigation).toBeVisible()
    } finally {
      await closeAndSaveVideo(launched, 'docs-reviewing-pane')
    }
  })
})
