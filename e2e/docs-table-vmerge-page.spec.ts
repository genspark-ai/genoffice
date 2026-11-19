/**
 * A vertically merged group crossing a page boundary must not inflate its row.
 *
 * genoffice#2114: a table with rowspan groups over many rows rendered later
 * group-head rows at page-tall heights \u2014 one line of text and then blank
 * space. Three defects fed one another: an in-cell gap band anchored inside a
 * merged row's cell (its height shifted every row the merge covers), the
 * vMerge continuation rows never charged to the column, and \u2014 the amplifier \u2014
 * the measure pass's display:none on gap table rows removed them from the
 * table grid, so a rowspan crossing one extended over more rendered rows and
 * every later row's cells shifted a column, collapsing a fixed-layout column
 * to zero width (one character per line, ~20 lines per list item, a page-tall
 * row) which the row-fill then patched into the DOM.
 *
 * This drives the real editor with the issue's shape \u2014 rowspan-7 groups,
 * ordered-list description cells \u2014 and asserts what a reader would check:
 * every row stays at its content height, no page-tall row appears, and no
 * full-height gap band is inserted inside a merged table's cells.
 */
import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, realpathSync, writeFileSync } from 'node:fs'
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

/** wait until the row-height signature stops changing (pagination settled) */
async function settledRowHeights(page: Page): Promise<number[]> {
  await page.waitForSelector('.doc-table td', { timeout: 30_000 })
  await page.evaluate(() => document.fonts.ready.then(() => undefined))
  const deadline = Date.now() + 60_000
  let prev = ''
  let stableSince = Date.now()
  while (Date.now() < deadline) {
    const cur = await page.evaluate(() => {
      const heights = Array.from(document.querySelectorAll('.doc-table tr'))
        .map((tr) => Math.round(tr.getBoundingClientRect().height))
        .join(',')
      // the page count rides along: the final gap batch can land after the
      // row heights settle, and the assertions read both
      const gaps = document.querySelectorAll(
        '.page-gap:not(.page-gap-carry):not(.page-cut-overlay), .page-gap-cut:not(.page-cut-overlay)',
      ).length
      return `${heights}|${gaps}`
    })
    if (cur !== prev) {
      prev = cur
      stableSince = Date.now()
    } else if (Date.now() - stableSince > 2500) {
      // a pagination pass can run over 1s without touching the DOM; re-check
      // after another quiet second before trusting the quiet
      await page.waitForTimeout(1000)
      const recheck = await page.evaluate(() => {
        const heights = Array.from(document.querySelectorAll('.doc-table tr'))
          .map((tr) => Math.round(tr.getBoundingClientRect().height))
          .join(',')
        const gaps = document.querySelectorAll(
          '.page-gap:not(.page-gap-carry):not(.page-cut-overlay), .page-gap-cut:not(.page-cut-overlay)',
        ).length
        return `${heights}|${gaps}`
      })
      if (recheck === cur) return cur.split('|')[0].split(',').map(Number)
      prev = recheck
      stableSince = Date.now()
    }
    await page.waitForTimeout(150)
  }
  return prev.split('|')[0].split(',').map(Number)
}

test.describe('a vertically merged table crossing pages', () => {
  test('keeps every row at its content height', async () => {
    test.setTimeout(300_000)
    const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-vmerge-')))
    const docPath = join(dir, 'anchor.docx')
    writeFileSync(docPath, await minimalDocx())

    const app = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-table-vmerge-page',
      openFile: docPath,
    })
    const page = await waitForPageWithUrl(app.app, '://docs/')
    try {
      await page.locator('.doc-page p', { hasText: 'Anchor paragraph.' }).first().waitFor()
      // the issue's shape: rowspan-7 groups in the first column, description
      // cells with ordered-list items, tall enough that groups straddle pages
      await page.evaluate(() => {
        const w = window as unknown as EditorWindow
        const ed = w.__aidocs!.editor as {
          state: { doc: { content: { size: number } } }
          commands: { insertContentAt: (pos: number, content: unknown) => boolean }
        }
        const para = (text: string) => ({ type: 'docParagraph', content: [{ type: 'text', text }] })
        const li = (text: string) => ({
          type: 'docListItem',
          attrs: { kind: 'ordered', ilvl: 0 },
          content: [{ type: 'text', text }],
        })
        const cell = (items: unknown[], extra: Record<string, unknown> = {}) => ({
          type: 'docTableCell',
          attrs: extra,
          content: items,
        })
        // explicit column widths (Word's w:tcW shape): the fixed layout is what
        // lets a column-shifted cell collapse to zero width
        const body: unknown[] = [
          {
            type: 'docTableRow',
            content: [
              cell([para('\u6a21\u5757')], { colwidth: [120] }),
              cell([para('\u5efa\u8bbe\u529f\u80fd')], { colwidth: [170] }),
              cell([para('\u529f\u80fd\u63cf\u8ff0')], { colwidth: [700] }),
            ],
          },
        ]
        for (let i = 1; i <= 30; i++) {
          const desc =
            i % 3 === 0
              ? [
                  li(
                    `1. \u652f\u6301\u6309\u533a\u57df\u3001\u7ebf\u8def\u3001\u65f6\u6bb5\u591a\u7ef4\u5ea6\u67e5\u8be2\u7b2c ${i} \u9879\u4e1a\u52a1\u6570\u636e\uff0c\u652f\u6301\u6309\u6761\u4ef6\u7b5b\u9009\u4e0e\u7ec4\u5408\u67e5\u8be2\u3002`,
                  ),
                  li(
                    `2. \u63d0\u4f9b\u65b0\u589e\u3001\u7f16\u8f91\u3001\u5220\u9664\u3001\u542f\u7528\u3001\u505c\u7528\u7b49\u64cd\u4f5c\uff0c\u64cd\u4f5c\u5168\u7a0b\u7559\u75d5\uff0c\u652f\u6301\u6309\u64cd\u4f5c\u4eba\u4e0e\u65f6\u95f4\u68c0\u7d22\u3002`,
                  ),
                  li(
                    `3. \u652f\u6301\u6309\u65e5\u3001\u5468\u3001\u6708\u81ea\u52a8\u6c47\u603b\u7edf\u8ba1\uff0c\u5f02\u5e38\u6570\u636e\u81ea\u52a8\u6807\u7ea2\u63d0\u9192\uff0c\u5e76\u63a8\u9001\u81f3\u8d23\u4efb\u4eba\u786e\u8ba4\u5904\u7406\u3002`,
                  ),
                  li(
                    `4. \u652f\u6301\u6279\u91cf\u5bfc\u5165\u5bfc\u51fa Excel \u6a21\u677f\uff0c\u5bfc\u5165\u65f6\u81ea\u52a8\u6821\u9a8c\u5fc5\u586b\u9879\u4e0e\u6570\u636e\u683c\u5f0f\uff0c\u9519\u8bef\u884c\u9000\u56de\u3002`,
                  ),
                  li(
                    `5. \u652f\u6301\u4e0e\u7b2c\u4e09\u65b9\u7cfb\u7edf\u5bf9\u63a5\uff0c\u63d0\u4f9b\u6807\u51c6 API \u63a5\u53e3\u4e0e\u8c03\u7528\u51ed\u8bc1\u7ba1\u7406\uff0c\u8c03\u7528\u65e5\u5fd7\u53ef\u67e5\u3002`,
                  ),
                ]
              : [
                  li(
                    `${i}. \u652f\u6301\u6570\u636e\u91c7\u96c6\u3001\u6cbb\u7406\u3001\u53ef\u89c6\u5316\u4e0e\u63a8\u9001\u3002`,
                  ),
                ]
          const cells: unknown[] = []
          if (i % 7 === 1)
            cells.push(
              cell([para(`\u6a21\u5757${Math.ceil(i / 7)}`)], { rowspan: 7, colwidth: [120] }),
            )
          cells.push(cell([para(`\u529f\u80fd\u9879${i}`)]))
          cells.push(cell(desc))
          body.push({ type: 'docTableRow', content: cells })
        }
        ed.commands.insertContentAt(ed.state.doc.content.size, [
          { type: 'docTable', content: body },
        ])
      })
      const heights = await settledRowHeights(page)

      // every row stays at its content height: the tallest legal row here is a
      // three-item description cell with wrapped lines (~600px); the inflated
      // rows of the bug measured whole pages
      const tallest = Math.max(...heights)
      expect(tallest).toBeLessThan(900)

      // the table spans several pages, so groups do cross page boundaries \u2014
      // the shape the bug needs \u2014 and no row may be left page-tall by the fill
      const pages = await page.evaluate(() => {
        const gaps = document.querySelectorAll(
          '.page-gap:not(.page-gap-carry):not(.page-cut-overlay), .page-gap-cut:not(.page-cut-overlay)',
        )
        return gaps.length + 1
      })
      expect(pages).toBeGreaterThan(2)
      expect(pages).toBeLessThan(15)

      // no full-height gap band inside any cell of a merged table: the bug
      // anchored 220px bands between the spans of a list item
      const bandsInMergedCells = await page.evaluate(() => {
        let count = 0
        for (const table of Array.from(document.querySelectorAll('.doc-table'))) {
          const merged = Array.from(table.querySelectorAll('td')).some(
            (td) => Number((td as HTMLTableCellElement).getAttribute('rowspan') || 1) > 1,
          )
          if (!merged) continue
          count += table.querySelectorAll('td .page-gap:not(.page-gap-cut)').length
        }
        return count
      })
      expect(bandsInMergedCells).toBe(0)
    } finally {
      await closeAndSaveVideo(app, 'docs-table-vmerge-page')
    }
  })
})
