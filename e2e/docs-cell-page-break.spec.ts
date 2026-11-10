/**
 * A page break typed inside a table cell has to turn the page.
 *
 * genoffice#1902: in a two-column layout built from a table, Ctrl+Enter inside a cell
 * rendered its marker and did nothing else — no page turn, and the new page
 * missing from print and PDF export for the same reason.
 *
 * The break reaches a cell in two shapes, and the editor produces the first:
 * Ctrl+Enter splits the paragraph and sets `page-break-before` on the new one,
 * which only the block-level reader honours, so inside a cell nobody was
 * reading it. A `w:br w:type="page"` from the source is the other shape, and
 * `styles.css` hides it inside table cells — so it carries no position and
 * cannot be measured.
 *
 * The arithmetic is unit-tested in pagination-cell-page-break.test.ts; this
 * drives the real editor and asserts what a user would check: the row breaks
 * and the content below it lands on the next page.
 */
import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

interface EditorWindow extends Window {
  __aidocs?: { editor: unknown }
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

/** the smallest document Word and the engine both accept: one paragraph */
async function minimalDocx(): Promise<Buffer> {
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
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body><w:p><w:r><w:t>Anchor paragraph.</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** wait until pagination stops mutating: page count + last-page bottom stable */
async function settledPageCount(page: Page): Promise<number> {
  await page.waitForSelector('.doc-page', { timeout: 30_000 })
  await page.evaluate(() => document.fonts.ready.then(() => undefined))
  const deadline = Date.now() + 30_000
  let prev = ''
  let stableSince = Date.now()
  while (Date.now() < deadline) {
    const cur = await page.evaluate(() => {
      // pages open at gap widgets in ONE continuous flow (.doc-page is the single
      // ProseMirror root). A zero-height .page-gap-cut marker also opens a page —
      // an in-row cut in a multi-cell row renders as one; the paper layer is
      // painted from exactly this set (pageFramesFromGaps), overlays excluded
      const gaps = document.querySelectorAll(
        '.page-gap:not(.page-gap-carry):not(.page-cut-overlay), .page-gap-cut:not(.page-cut-overlay)',
      )
      const root = document.querySelector('.doc-page')
      const bottom = root ? Math.round(root.getBoundingClientRect().height) : 0
      return `${gaps.length + 1}:${bottom}`
    })
    if (cur !== prev) {
      prev = cur
      stableSince = Date.now()
    } else if (Date.now() - stableSince > 800) {
      return Number(cur.split(':')[0])
    }
    await page.waitForTimeout(120)
  }
  return Number(prev.split(':')[0])
}

test.describe('page break inside a table cell', () => {
  test('breaks the row so the content below lands on a new page', async () => {
    test.setTimeout(300_000)
    const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-cell-break-')))
    const docPath = join(dir, 'anchor.docx')
    writeFileSync(docPath, await minimalDocx())

    const app = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-cell-page-break',
      openFile: docPath,
    })
    const page = await waitForPageWithUrl(app.app, '://docs/')
    try {
      await page.locator('.doc-page p', { hasText: 'Anchor paragraph.' }).first().waitFor()

      // 1×3 — two content columns and a narrow divider, the shape genoffice#1902 uses,
      // appended at the end. Short enough that the row fits page 1 whole: the
      // only thing that can then add a page is the break itself
      await page.evaluate(() => {
        const w = window as unknown as EditorWindow
        const ed = w.__aidocs!.editor as {
          state: { doc: { content: { size: number } } }
          commands: { insertContentAt: (pos: number, content: unknown) => boolean }
        }
        const para = (text: string) => ({ type: 'docParagraph', content: [{ type: 'text', text }] })
        const cell = (text: string) => ({ type: 'docTableCell', content: [para(text)] })
        ed.commands.insertContentAt(ed.state.doc.content.size, [
          {
            type: 'docTable',
            content: [
              {
                type: 'docTableRow',
                content: [
                  cell('first column, which will run past the page boundary '.repeat(12)),
                  cell('divider'),
                  cell('second column follows the same break '.repeat(12)),
                ],
              },
            ],
          },
        ])
      })
      await page.waitForSelector('.doc-table td', { timeout: 30_000 })
      const before = await settledPageCount(page)

      // the chord the editor binds for a page break, with the caret in a cell.
      // End first: at a paragraph's start the editor emits the br shape (hidden
      // inside cells by styles.css) instead of the page-break-before this spec
      // is about, so the caret has to sit inside the paragraph.
      await page.locator('.doc-table td').first().click()
      await page.keyboard.press('End')
      // the keymap binds Mod-Enter: Command on macOS, Control elsewhere
      await page.keyboard.press(process.platform === 'darwin' ? 'Meta+Enter' : 'Control+Enter')

      // the instruction must exist before pagination can honour it — a missing
      // flag means the input never landed, not that the reader is broken
      const marked = await page.evaluate(() => {
        const w = window as unknown as EditorWindow
        let found = 0
        const walk = (node: {
          type?: string
          attrs?: Record<string, unknown>
          content?: unknown[]
        }) => {
          if (node.attrs?.pageBreakBefore === true) found++
          for (const child of node.content ?? []) walk(child as typeof node)
        }
        walk(
          (
            w.__aidocs!.editor as { state: { doc: { toJSON: () => unknown } } }
          ).state.doc.toJSON() as never,
        )
        return found
      })
      expect(marked, 'Ctrl+Enter must set page-break-before inside the cell').toBeGreaterThan(0)

      const after = await settledPageCount(page)
      expect(after, `the break must add a page — before=${before}, after=${after}`).toBeGreaterThan(
        before,
      )
    } finally {
      await closeAndSaveVideo(app, 'docs-cell-page-break')
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
