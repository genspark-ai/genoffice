/**
 * Review → Show Markup → Balloons (Word's three modes). The default, "Show
 * Only Comments and Formatting in Balloons", keeps deletions inline and
 * balloons comments / format changes; "Show Revisions in Balloons" moves the
 * deletion to a Deleted: balloon and leaves a marker in the text; "Show All
 * Revisions Inline" drops the markup column and tags comments [initials N].
 * The choice is an application preference that persists.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const REV = 'w:author="Ann Lee" w:date="2026-09-01T10:00:00Z"'

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const filler = Array.from({ length: 3 }, (_, i) => p(`Filler line ${i + 1}.`)).join('')

async function trackedDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
  )
  zip.file(
    'word/comments.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:comments ${W}><w:comment w:id="0" w:initials="AL" ${REV}><w:p><w:r><w:t>Please check this sentence.</w:t></w:r></w:p></w:comment></w:comments>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>` +
      p('Alpha opens the document.') +
      filler +
      `<w:p><w:r><w:t xml:space="preserve">Gamma keeps </w:t></w:r><w:del w:id="2" ${REV}><w:r><w:delText>removed words</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> around.</w:t></w:r></w:p>` +
      filler +
      `<w:p><w:r><w:t xml:space="preserve">Delta made </w:t></w:r><w:r><w:rPr><w:b/><w:rPrChange w:id="3" ${REV}><w:rPr/></w:rPrChange></w:rPr><w:t>bold words</w:t></w:r><w:r><w:t xml:space="preserve"> here.</w:t></w:r></w:p>` +
      filler +
      `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Epsilon carries a comment.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>` +
      filler +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test.describe('docs Show Markup balloons', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-balloons-')))
    docPath = join(dir, 'tracked.docx')
    writeFileSync(docPath, await trackedDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('three balloon modes from the Markup menu; the choice persists', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-balloon-modes',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.locator('.doc-page p', { hasText: 'Alpha opens' }).first().waitFor()
      const root = page.locator('.app')
      const wrap = page.locator('.editor-scroll .page-wrap')
      const del = page.locator('.doc-del', { hasText: 'removed words' })
      const commentBubble = page.locator('.comment-bubble:not(.rev-bubble)')
      const delBubble = page.locator('.rev-bubble-del')
      const fmtBubble = page.locator('.rev-bubble-fmt')

      // default (Word): comments + formatting in balloons, deletion struck inline
      await expect(root).toHaveAttribute('data-balloons', 'comments')
      await expect(commentBubble).toHaveCount(1, { timeout: 20_000 })
      await expect(fmtBubble).toHaveCount(1)
      await expect(delBubble).toHaveCount(0)
      await expect(del).toBeVisible()
      await expect(del).toHaveCSS('text-decoration-line', 'line-through')
      await expect(wrap).toHaveClass(/has-markup-area/)
      // markup strip (200) + Word's separate comment column (260)
      expect(await wrap.evaluate((el) => el.style.getPropertyValue('--markup-w'))).toBe('460px')

      await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
      const markup = page.getByRole('button', { name: 'Markup', exact: true })
      await markup.click()
      const items = page.locator('.layout-menu button[role="menuitemradio"]')
      await expect(items).toHaveText([
        'Show Revisions in Balloons',
        'Show All Revisions Inline',
        'Show Only Comments and Formatting in Balloons ✓',
      ])
      await expect(items.nth(2)).toHaveAttribute('aria-checked', 'true')

      // Show Revisions in Balloons: deletion leaves the flow for a Deleted: balloon,
      // a marker stays at its spot and right-clicking it offers the change's menu
      await items.nth(0).click()
      await expect(root).toHaveAttribute('data-balloons', 'revisions')
      await expect(root).toHaveClass(/rev-balloon/)
      await expect(delBubble).toHaveCount(1, { timeout: 20_000 })
      await expect(delBubble).toContainText('removed words')
      await expect(fmtBubble).toHaveCount(1)
      await expect(commentBubble).toHaveCount(1)
      await expect(del).toBeHidden()
      const marker = page.locator('.doc-del-marker')
      await expect(marker).toHaveCount(1)
      await expect(marker).toBeVisible()
      await marker.click({ button: 'right' })
      await expect(page.locator('.ctx-menu')).toBeVisible()
      await expect(page.locator('.ctx-menu')).toContainText(/Accept/)
      await page.keyboard.press('Escape')

      // Show All Revisions Inline: no balloons, no markup column, [AL1] comment tag
      await markup.click()
      await page.locator('.layout-menu button', { hasText: 'Show All Revisions Inline' }).click()
      await expect(root).toHaveAttribute('data-balloons', 'inline')
      await expect(root).not.toHaveClass(/rev-balloon/)
      await expect(page.locator('.comment-bubble')).toHaveCount(0)
      await expect(wrap).not.toHaveClass(/has-markup-area/)
      await expect(del).toBeVisible()
      const tag = page.locator('.doc-comment-marker')
      await expect(tag).toHaveText('[AL1]')
      await expect(tag).toHaveAttribute('title', /Ann Lee: Please check this sentence\./)
      await tag.click()
      await expect(page.locator('.comments-pane')).toBeVisible()
      await expect(page.locator('.comments-pane [data-comment-id="0"] .comment-card')).toHaveClass(
        /active/,
      )
      expect(await page.evaluate(() => localStorage.getItem('aidocs.balloonMode'))).toBe('inline')
      // the pane replaces the margin cards while open (Word): close it so they can come back
      await page.locator('.comments-pane-close').click()

      // back to the default: tag gone, bubbles back, deletion still inline
      await markup.click()
      await page
        .locator('.layout-menu button', {
          hasText: 'Show Only Comments and Formatting in Balloons',
        })
        .click()
      await expect(root).toHaveAttribute('data-balloons', 'comments')
      await expect(tag).toHaveCount(0)
      await expect(commentBubble).toHaveCount(1, { timeout: 20_000 })
      await expect(delBubble).toHaveCount(0)
      await expect(del).toBeVisible()
    } finally {
      await closeAndSaveVideo(launched, 'docs-balloon-modes')
    }
  })
})
