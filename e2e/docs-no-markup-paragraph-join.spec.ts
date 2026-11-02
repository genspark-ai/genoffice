/**
 * A paragraph whose mark is a tracked deletion: All Markup keeps the two
 * paragraphs apart (struck pilcrow), No Markup shows the final text, where the
 * second paragraph continues the first one on the same line (Word).
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const REV = 'w:author="Ann" w:date="2026-09-01T10:00:00Z"'

async function docx(): Promise<Buffer> {
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
      `<w:p><w:pPr><w:rPr><w:del w:id="1" ${REV}/></w:rPr></w:pPr><w:r><w:t>Head paragraph.</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Tail paragraph.</w:t></w:r></w:p>` +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const lineTops = () =>
  Array.from(document.querySelectorAll('.ProseMirror > p')).map((p) => {
    const r = document.createRange()
    r.selectNodeContents(p.firstChild!)
    return Math.round(r.getClientRects()[0]?.top ?? -1)
  })

test.describe('docs No Markup paragraph join', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-para-join-')))
    docPath = join(dir, 'join.docx')
    writeFileSync(docPath, await docx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('two line boxes in All Markup, one in No Markup; the caret still enters the tail', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-no-markup-paragraph-join',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      const head = page.locator('.ProseMirror > p', { hasText: 'Head paragraph.' })
      const tail = page.locator('.ProseMirror > p', { hasText: 'Tail paragraph.' })
      await head.waitFor()
      await expect(head).toHaveClass(/doc-para-mark-del/)
      const apart = await page.evaluate(lineTops)
      expect(apart[0]).toBeLessThan(apart[1])

      await page.evaluate(() =>
        (window as unknown as { __setRevisionDisplay: (m: string) => void }).__setRevisionDisplay(
          'none',
        ),
      )
      await expect(page.locator('.app')).toHaveClass(/rev-display-none/)
      await expect(head).toHaveClass(/doc-para-join-head/)
      await expect(tail).toHaveClass(/doc-para-join-tail/)
      await expect(head).toHaveCSS('display', 'inline')
      await expect.poll(() => page.evaluate(lineTops)).toEqual([apart[0], apart[0]])

      // the tail is inline now: its box spans from the shared line, so aim at
      // the centre of its first text rect instead of the box corner
      const target = await tail.evaluate((el) => {
        const r = document.createRange()
        r.selectNodeContents(el.firstChild!)
        const b = r.getClientRects()[0]
        return { x: b.left + b.width / 2, y: b.top + b.height / 2 }
      })
      await page.mouse.click(target.x, target.y)
      // ProseMirror reads the DOM selection back asynchronously
      await expect
        .poll(() =>
          page.evaluate(
            () =>
              (
                window as unknown as {
                  __aidocs?: {
                    editor?: {
                      state: { selection: { $from: { parent: { textContent: string } } } }
                    }
                  }
                }
              ).__aidocs?.editor?.state.selection.$from.parent.textContent,
          ),
        )
        .toBe('Tail paragraph.')

      await page.evaluate(() =>
        (window as unknown as { __setRevisionDisplay: (m: string) => void }).__setRevisionDisplay(
          'all',
        ),
      )
      await expect(head).not.toHaveClass(/doc-para-join-head/)
      await expect.poll(() => page.evaluate(lineTops)).toEqual(apart)
    } finally {
      await closeAndSaveVideo(launched)
    }
  })
})
