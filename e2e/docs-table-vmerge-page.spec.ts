/**
 * A vertically merged group crossing a page boundary must not inflate its row.
 *
 * genoffice#2114: a table with rowspan groups over many rows rendered later
 * group-head rows at page-tall heights — one line of text and then blank
 * space. Three defects fed one another: an in-cell gap band anchored inside a
 * merged row's cell (its height shifted every row the merge covers), the
 * vMerge continuation rows never charged to the column, and — the amplifier —
 * the measure pass's display:none on gap table rows removed them from the
 * table grid, so a rowspan crossing one extended over more rendered rows and
 * every later row's cells shifted a column, collapsing a fixed-layout column
 * to zero width (one character per line, ~20 lines per list item, a page-tall
 * row) which the row-fill then patched into the DOM.
 *
 * This drives the real editor with the issue's shape — rowspan-7 groups,
 * ordered-list description cells — and asserts what a reader would check:
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
              cell([para('模块')], { colwidth: [120] }),
              cell([para('建设功能')], { colwidth: [170] }),
              cell([para('功能描述')], { colwidth: [700] }),
            ],
          },
        ]
        for (let i = 1; i <= 30; i++) {
          const desc =
            i % 3 === 0
              ? [
                  li(
                    `1. 支持按区域、线路、时段多维度查询第 ${i} 项业务数据，支持按条件筛选与组合查询。`,
                  ),
                  li(
                    `2. 提供新增、编辑、删除、启用、停用等操作，操作全程留痕，支持按操作人与时间检索。`,
                  ),
                  li(
                    `3. 支持按日、周、月自动汇总统计，异常数据自动标红提醒，并推送至责任人确认处理。`,
                  ),
                  li(
                    `4. 支持批量导入导出 Excel 模板，导入时自动校验必填项与数据格式，错误行退回。`,
                  ),
                  li(`5. 支持与第三方系统对接，提供标准 API 接口与调用凭证管理，调用日志可查。`),
                ]
              : [li(`${i}. 支持数据采集、治理、可视化与推送。`)]
          const cells: unknown[] = []
          if (i % 7 === 1)
            cells.push(cell([para(`模块${Math.ceil(i / 7)}`)], { rowspan: 7, colwidth: [120] }))
          cells.push(cell([para(`功能项${i}`)]))
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

      // the table spans several pages, so groups do cross page boundaries —
      // the shape the bug needs — and no row may be left page-tall by the fill
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
