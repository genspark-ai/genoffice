/**
 * Word's Simple Markup review view: a document carrying tracked changes opens
 * with final text (deletions hidden, insertions plain), a red changed-line bar
 * in the left margin for every changed line and a balloon glyph per comment in
 * the right margin. Clicking a bar switches to All Markup at that change;
 * clicking a glyph focuses the thread in the Comments panel.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

interface AidocsWindow {
  __aidocs?: { editor?: { state: { selection: { $from: { parent: { textContent: string } } } } } }
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const REV = 'w:author="Ann" w:date="2026-09-01T10:00:00Z"'

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
    `<?xml version="1.0" encoding="UTF-8"?><w:comments ${W}><w:comment w:id="0" w:initials="A" ${REV}><w:p><w:r><w:t>Please check this sentence.</w:t></w:r></w:p></w:comment></w:comments>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>` +
      p('Alpha opens the document.') +
      filler +
      `<w:p><w:r><w:t xml:space="preserve">Beta keeps </w:t></w:r><w:ins w:id="1" ${REV}><w:r><w:t>inserted words</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> at the end.</w:t></w:r></w:p>` +
      filler +
      `<w:p><w:r><w:t xml:space="preserve">Gamma keeps </w:t></w:r><w:del w:id="2" ${REV}><w:r><w:delText>removed words</w:delText></w:r></w:del><w:r><w:t xml:space="preserve"> around.</w:t></w:r></w:p>` +
      filler +
      `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Delta carries a comment.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>` +
      filler +
      // table: plain row, inserted row, plain row, deleted row (the deleted row's bar anchors on the row above)
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>Row kept</w:t></w:r></w:p></w:tc></w:tr>' +
      `<w:tr><w:trPr><w:ins w:id="3" ${REV}/></w:trPr><w:tc><w:p><w:ins w:id="4" ${REV}><w:r><w:t>Row added</w:t></w:r></w:ins></w:p></w:tc></w:tr>` +
      '<w:tr><w:tc><w:p><w:r><w:t>Row kept too</w:t></w:r></w:p></w:tc></w:tr>' +
      `<w:tr><w:trPr><w:del w:id="5" ${REV}/></w:trPr><w:tc><w:p><w:del w:id="6" ${REV}><w:r><w:delText>Row removed</w:delText></w:r></w:del></w:p></w:tc></w:tr>` +
      '</w:tbl>' +
      filler +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test.describe('docs Simple Markup view', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-simple-markup-')))
    docPath = join(dir, 'tracked.docx')
    writeFileSync(docPath, await trackedDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('opens in Simple Markup with change bars and balloons; a bar click shows All Markup', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-simple-markup',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.locator('.doc-page p', { hasText: 'Alpha opens' }).first().waitFor()

      // default view for a document with w:ins / w:del
      const root = page.locator('.app')
      await expect(root).toHaveClass(/rev-display-simple/)
      const bars = page.locator('.change-bar')
      await expect(bars.first()).toBeVisible({ timeout: 20_000 })
      // one bar per changed line: insertion, hidden deletion, inserted row, hidden deleted row
      await expect(bars).toHaveCount(4)
      await expect(page.locator('tr.row-rev-del')).toBeHidden()
      await expect(page.locator('tr.row-rev-ins')).toBeVisible()
      const ins = page.locator('.doc-ins', { hasText: 'inserted words' })
      const del = page.locator('.doc-del', { hasText: 'removed words' })
      await expect(ins).toBeVisible()
      await expect(del).toBeHidden()
      await expect(ins).toHaveCSS('text-decoration-line', 'none')
      // no bubble column: the comment collapses to a glyph in the right page margin
      await expect(page.locator('.comment-bubble')).toHaveCount(0)
      const balloon = page.locator('.comment-balloon')
      await expect(balloon).toHaveCount(1)
      await expect(balloon).toHaveAttribute('title', /Ann/)
      const paperRight = await page
        .locator('.ProseMirror')
        .first()
        .evaluate((el) => el.getBoundingClientRect().right)
      const balloonBox = await balloon.boundingBox()
      expect(balloonBox).not.toBeNull()
      expect(balloonBox!.x + balloonBox!.width).toBeLessThanOrEqual(paperRight + 1)

      // Markup menu in Word's order with the current view ticked
      await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
      const markup = page.getByRole('button', { name: 'Markup', exact: true })
      await markup.click()
      await expect(page.locator('.layout-menu button')).toHaveText([
        'Simple Markup ✓',
        'All Markup',
        'No Markup',
        'Original',
      ])
      await markup.click()
      await expect(page.locator('.layout-menu')).toHaveCount(0)

      // balloon glyph -> Comments panel focused on the thread
      await balloon.click()
      await expect(page.locator('.comments-pane')).toBeVisible()
      await expect(page.locator('.comments-pane [data-comment-id="0"] .comment-card')).toHaveClass(
        /active/,
      )

      // changed-line bar of the deletion -> All Markup: the deletion is back as
      // a "Deleted:" margin balloon (print layout) and comments regain their bubbles
      await bars.nth(1).click()
      await expect(root).not.toHaveClass(/rev-display-/)
      await expect(page.locator('.rev-bubble-del', { hasText: 'removed words' })).toBeVisible()
      await expect(page.locator('.comment-balloon')).toHaveCount(0)
      await expect(page.locator('.comment-bubble:not(.rev-bubble)')).toHaveCount(1)

      // back in Simple Markup, Next lands on the visible insertion and stays;
      // the following Next reaches the hidden deletion and falls back to All Markup
      await markup.click()
      await page.locator('.layout-menu button', { hasText: 'Simple Markup' }).click()
      await expect(root).toHaveClass(/rev-display-simple/)
      await page.locator('.doc-page p', { hasText: 'Alpha opens' }).first().click()
      const nextChange = page.locator('button[data-tip="Go to the next change"]')
      await nextChange.click()
      await expect(root).toHaveClass(/rev-display-simple/)
      await expect(ins).toBeVisible()
      await nextChange.click()
      await expect(root).not.toHaveClass(/rev-display-/)
      await expect(page.locator('.rev-bubble-del', { hasText: 'removed words' })).toBeVisible()

      // a row bar carries a node position: the click still lands the caret in that row
      await markup.click()
      await page.locator('.layout-menu button', { hasText: 'Simple Markup' }).click()
      await expect(root).toHaveClass(/rev-display-simple/)
      await expect(bars).toHaveCount(4)
      await bars.nth(2).click()
      await expect(root).not.toHaveClass(/rev-display-/)
      expect(
        await page.evaluate(() => {
          const ed = (window as unknown as AidocsWindow).__aidocs?.editor
          return ed?.state.selection.$from.parent.textContent ?? ''
        }),
      ).toBe('Row added')
    } finally {
      await closeAndSaveVideo(launched, 'docs-simple-markup')
    }
  })
})
