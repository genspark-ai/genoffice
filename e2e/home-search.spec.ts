import { test, expect } from '@playwright/test'
import { mkdirSync, mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

/**
 * Home file search: names, folders and extracted content over the default
 * save folder, with CJK bigram matching and highlighted snippets.
 */

const HAN_REPORT = '\u62a5\u544a' // two CJK characters used as the content probe
const HAN_BODY =
  '\u672c\u5e74\u5ea6\u65b0\u80fd\u6e90\u6c7d\u8f66\u5e02\u573a\u8c03\u7814' + HAN_REPORT

async function minimalDocx(text: string): Promise<Buffer> {
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
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test.describe('home file search', () => {
  let root: string

  test.beforeEach(async () => {
    root = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-search-')))
    mkdirSync(join(root, 'Finance'))
    writeFileSync(join(root, 'Finance', 'quarterly-plan.docx'), await minimalDocx(HAN_BODY))
    writeFileSync(
      join(root, 'meeting-notes.md'),
      '# Notes\n\nDiscussed the annual budget review.\n',
    )
    writeFileSync(
      join(root, 'landing.html'),
      '<html><body><h1>Launch</h1><p>Budget approved.</p></body></html>',
    )
    writeFileSync(join(root, 'unrelated.md'), 'nothing to see here')
  })

  test.afterEach(() => {
    rmSync(root, { recursive: true, force: true })
  })

  test('finds files by content and name, highlights hits, opens a result', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root },
      videoDir: 'home-search',
    })
    const { page } = launched
    try {
      const box = page.locator('.file-search input')
      await expect(box).toBeVisible()

      // content match across two formats; the index fills in the background
      await box.fill('budget')
      const rows = page.locator('.search-row')
      await expect(rows).toHaveCount(2, { timeout: 30_000 })
      await expect(page.locator('.search-name', { hasText: 'meeting-notes.md' })).toBeVisible()
      await expect(page.locator('.search-name', { hasText: 'landing.html' })).toBeVisible()
      await expect(page.locator('.search-snippet .search-hit').first()).toHaveText(/budget/i)
      await expect(page.locator('.recents-heading .file-count')).toHaveText('2 results')

      // CJK bigram inside a longer run, extracted from a docx
      await box.fill(HAN_REPORT)
      await expect(rows).toHaveCount(1, { timeout: 30_000 })
      await expect(page.locator('.search-name')).toHaveText('quarterly-plan.docx')
      await expect(page.locator('.search-snippet .search-hit')).toHaveText(HAN_REPORT)
      await page.screenshot({ path: screenshotPath('home-search-cjk') })

      // a CJK phrase typed one character too far still reaches the file
      await box.fill(HAN_BODY.slice(-4) + '\u6211')
      await expect(rows).toHaveCount(1)
      await expect(page.locator('.search-snippet .search-hit')).toHaveText(HAN_BODY.slice(-4))

      // name prefix match with the name highlighted, filter pills narrow it
      await box.fill('quart')
      await expect(rows).toHaveCount(1)
      await expect(page.locator('.search-name .search-hit')).toHaveText('quart')
      await page.locator('.filter-pill', { hasText: 'Markdown' }).click()
      await expect(page.locator('.search-results .empty-hint')).toContainText('quart')
      await page.locator('.filter-pill', { hasText: 'All' }).click()
      await expect(rows).toHaveCount(1)

      // a folder-name hit lights up the location label
      await box.fill('finance')
      await expect(rows).toHaveCount(1)
      await expect(page.locator('.search-path .search-hit')).toHaveText('Finance')

      // Escape clears; ⌘F / Ctrl+F focuses
      await box.press('Escape')
      await expect(box).toHaveValue('')
      await expect(page.locator('.search-results')).toHaveCount(0)
      await expect(page.locator('.recents-heading .section-label')).toHaveText('Recent')
      await page.locator('body').click()
      await page.keyboard.press(process.platform === 'darwin' ? 'Meta+f' : 'Control+f')
      await expect(box).toBeFocused()

      // the sort button beside the search box lands on the Jev reranking block of Settings
      await page.locator('.file-search-group .file-search-settings').click()
      await expect(page.locator('.set-nav-item.active')).toHaveText('AI Media & Search')
      await expect(
        page.locator('.set-pane-subhead', { hasText: 'Local file search' }),
      ).toBeInViewport()
      await page.locator('.set-close').click()

      // opening a result switches to an editor tab
      await box.fill('meeting')
      await expect(rows).toHaveCount(1)
      await rows.first().click()
      await expect(page.locator('.tab-bar .tab-item', { hasText: 'meeting-notes' })).toBeVisible({
        timeout: 20_000,
      })
    } finally {
      await closeAndSaveVideo(launched, 'home-search')
    }
  })
})
