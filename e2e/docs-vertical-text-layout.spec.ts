/**
 * Vertical text as Word lays it out: a table of a vertical section stays
 * upright and on the page, a paragraph of a vertical cell keeps its line length
 * (its spacing turns with the text), and a Horizontal in Vertical run is turned
 * a quarter left of its line at its own size. Each case is checked on the page
 * and in the exported PDF.
 */
import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { docxWith, para } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-vertical-text-layout'
const sectPr = (direction = '') =>
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  `<w:cols w:space="720"/>${direction}<w:docGrid w:linePitch="360"/></w:sectPr>`
const VERTICAL = '<w:textDirection w:val="tbRl"/>'
const LONG = 'A long line of text that runs down the page when the section is vertical. '
const cell = (content: string, width: number, direction = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${direction}</w:tcPr>${content}</w:tc>`
const table = (cells: string, widths: number[], rowPr = '') =>
  `<w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((a, b) => a + b)}" w:type="dxa"/></w:tblPr>` +
  `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
  `<w:tr>${rowPr ? `<w:trPr>${rowPr}</w:trPr>` : ''}${cells}</w:tr></w:tbl>`
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const upright = (text: string, id: number, fit = false) =>
  `<w:r><w:rPr><w:eastAsianLayout w:id="${id}" w:vert="1"${fit ? ' w:vertCompress="1"' : ''}/></w:rPr><w:t>${text}</w:t></w:r>`

type Box = {
  left: number
  top: number
  right: number
  bottom: number
  width: number
  height: number
}
const box = (page: Page, selector: string, text?: string): Promise<Box> =>
  (text ? page.locator(selector, { hasText: text }) : page.locator(selector))
    .first()
    .evaluate((el) => {
      const r = el.getBoundingClientRect()
      return {
        left: r.left,
        top: r.top,
        right: r.right,
        bottom: r.bottom,
        width: r.width,
        height: r.height,
      }
    })

/** where the text of an element is drawn (its box may start earlier, e.g. beside a table) */
const inkBox = (page: Page, selector: string, text: string): Promise<Box> =>
  page
    .locator(selector, { hasText: text })
    .first()
    .evaluate((el) => {
      const range = document.createRange()
      range.selectNodeContents(el)
      const r = range.getBoundingClientRect()
      return {
        left: r.left,
        top: r.top,
        right: r.right,
        bottom: r.bottom,
        width: r.width,
        height: r.height,
      }
    })

type PdfItem = { str: string; x: number; y: number; a: number; b: number; d: number }
async function pdfItems(file: string): Promise<{ pages: number; items: PdfItem[] }> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(file)) }).promise
  const content = await (await doc.getPage(1)).getTextContent()
  const items = content.items
    .filter(
      (it): it is { transform: number[]; str: string } & typeof it =>
        'transform' in it && 'str' in it && (it as { str: string }).str.trim().length > 0,
    )
    .map((it) => ({
      str: it.str.trim(),
      x: it.transform[4],
      y: it.transform[5],
      a: it.transform[0],
      b: it.transform[1],
      d: it.transform[3],
    }))
  return { pages: doc.numPages, items }
}

async function open(docPath: string): Promise<{ launched: LaunchedApp; page: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p').first()).toBeVisible({ timeout: 30_000 })
  return { launched, page }
}

async function exportPdf(launched: LaunchedApp, page: Page, out: string) {
  await launched.app.evaluate(({ dialog, shell }, file) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: file })
    shell.showItemInFolder = () => {}
  }, out)
  expect(
    await page.evaluate(() =>
      (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
    ),
  ).toBe(true)
  return pdfItems(out)
}

async function withDoc(
  name: string,
  body: string,
  check: (ctx: { launched: LaunchedApp; page: Page; dir: string }) => Promise<void>,
) {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-vertical-layout-')))
  const docPath = join(dir, `${name}.docx`)
  writeFileSync(docPath, await docxWith(body))
  const { launched, page } = await open(docPath)
  try {
    await check({ launched, page, dir })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
    rmSync(launched.userDataDir, { recursive: true, force: true })
  }
}

test('a table in a vertical section stays upright on the page, with the next text below it', async () => {
  test.setTimeout(240_000)
  const body =
    para(LONG.repeat(3)) +
    table(cell(para('cell one'), 4000) + cell(para('cell two'), 4000), [4000, 4000]) +
    para('banana') +
    sectPr(VERTICAL)
  await withDoc('section-table', body, async ({ launched, page, dir }) => {
    const settled = async () => {
      const t = await box(page, '.doc-page table')
      const b = await inkBox(page, '.doc-page > p', 'banana')
      return Math.abs(b.top - t.bottom) < 3
    }
    await expect.poll(settled, { timeout: 30_000 }).toBe(true)
    const paper = await box(page, '.doc-page')
    const scale = paper.width / 816
    const inch = 96 * scale
    const tbl = await box(page, '.doc-page table')
    const long = await box(page, '.doc-page p', 'A long line')
    const banana = await inkBox(page, '.doc-page > p', 'banana')
    const one = await box(page, '.doc-page td', 'cell one')
    const two = await box(page, '.doc-page td', 'cell two')

    // inside the margins of the page, at the top of the lines
    expect(tbl.left).toBeGreaterThan(paper.left + inch - 2)
    expect(tbl.right).toBeLessThan(paper.right - inch + 2)
    expect(Math.abs(tbl.top - (paper.top + inch))).toBeLessThan(3)
    // upright: the cells side by side, their text horizontal
    expect(Math.abs(one.top - two.top)).toBeLessThan(1)
    expect(two.left).toBeGreaterThan(one.left + 2 * inch)
    expect(one.width).toBeGreaterThan(one.height * 3)
    expect(
      await page
        .locator('.doc-page td p', { hasText: 'cell one' })
        .evaluate((el) => getComputedStyle(el).writingMode),
    ).toBe('horizontal-tb')
    // its far edge is where the text before it ends; it takes no room of the flow,
    // so the next paragraph is the next line, starting below the table
    expect(Math.abs(tbl.right - long.left)).toBeLessThan(3)
    expect(banana.right).toBeLessThan(tbl.right + 2)
    expect(banana.right).toBeGreaterThan(tbl.right - 0.5 * inch)
    expect(banana.height).toBeGreaterThan(banana.width)
    await page.screenshot({ path: test.info().outputPath('section-table.png') })

    const pdf = await exportPdf(launched, page, join(dir, 'section-table.pdf'))
    expect(pdf.pages).toBe(1)
    const flat = pdf.items.filter((it) => Math.abs(it.b) < 0.01)
    const cells = ['cell one', 'cell two'].map(
      (text) => flat.find((it) => it.str === text) ?? flat.find((it) => text.startsWith(it.str)),
    )
    expect(cells[0]).toBeDefined()
    expect(cells[1]).toBeDefined()
    // 612pt wide paper with 72pt margins; the top line of an 792pt tall page
    for (const c of cells) {
      expect(c!.x).toBeGreaterThan(72)
      expect(c!.x).toBeLessThan(540)
      expect(c!.y).toBeGreaterThan(792 - 72 - 30)
    }
    expect(cells[1]!.x - cells[0]!.x).toBeCloseTo(200, -1)
    const turned = pdf.items.filter((it) => Math.abs(it.b) > 0.5).map((it) => it.str)
    expect(turned.join(' ')).toContain('banana')
  })
})

test('text in a vertical cell breaks at the space in a row as tall as its horizontal neighbour', async () => {
  test.setTimeout(240_000)
  const tbRl = '<w:textDirection w:val="tbRl"/>'
  const body =
    para(LONG) +
    table(cell(para('cell one'), 4000, tbRl) + cell(para('cell two'), 4000), [4000, 4000]) +
    para('between') +
    table(
      cell(para('internationalization'), 1200, tbRl) + cell(para('x'), 3000),
      [1200, 3000],
      '<w:trHeight w:val="400" w:hRule="exact"/>',
    ) +
    para('banana') +
    sectPr()
  await withDoc('vertical-cell', body, async ({ launched, page, dir }) => {
    const turnedPara = page.locator('.doc-page td .cell-vert > p').first()
    await expect(turnedPara).toHaveAttribute('data-vbox', /--vb-ml/, { timeout: 30_000 })
    // the columns of the paragraph: one per line of the turned text
    const columns = () =>
      turnedPara.evaluate((el) => {
        const range = document.createRange()
        range.selectNodeContents(el)
        const lefts = new Set<number>()
        const words: string[] = []
        const text = el.textContent ?? ''
        const node = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode()!
        for (let i = 0; i < text.length; i++) {
          range.setStart(node, i)
          range.setEnd(node, i + 1)
          const left = Math.round(range.getBoundingClientRect().left)
          if (!lefts.has(left)) words.push('')
          lefts.add(left)
          words[words.length - 1] += text[i]
        }
        return words.map((w) => w.trim())
      })
    await expect.poll(columns, { timeout: 30_000 }).toEqual(['cell', 'one'])

    const vertical = await box(page, '.doc-page td', 'cell one')
    const neighbour = await box(page, '.doc-page td', 'cell two')
    const neighbourPara = await box(page, '.doc-page td p', 'cell two')
    const plain = await box(page, '.doc-page > p', 'between')
    // the row is the horizontal cell's: one line and its space after, not a taller one
    expect(Math.abs(vertical.height - neighbour.height)).toBeLessThan(1)
    expect(neighbour.height).toBeLessThan(neighbourPara.height * 2)
    expect(neighbourPara.height).toBeCloseTo(plain.height, 0)
    // the space after the paragraph lies beside it, not under it
    const margins = await turnedPara.evaluate((el) => {
      const cs = getComputedStyle(el)
      return { left: parseFloat(cs.marginLeft), bottom: parseFloat(cs.marginBottom) }
    })
    expect(margins.left).toBeGreaterThan(5)
    expect(margins.bottom).toBe(0)

    // a row of an exact height: the turned text stays in its cell
    const exactCell = await box(page, '.doc-page td', 'internationalization')
    const exactWrap = await box(page, '.doc-page td .cell-clip', 'internationalization')
    expect(exactWrap.right).toBeLessThan(exactCell.right + 1)
    expect(exactWrap.left).toBeGreaterThan(exactCell.left - 1)
    await page.screenshot({ path: test.info().outputPath('vertical-cell.png') })

    const pdf = await exportPdf(launched, page, join(dir, 'vertical-cell.pdf'))
    expect(pdf.pages).toBe(1)
    // the first table's turned text, above the paragraph between the tables
    const between = pdf.items.find((it) => it.str === 'between')!
    const turned = pdf.items.filter((it) => Math.abs(it.b) > 0.5 && it.y > between.y)
    const byColumn = new Map<number, string>()
    for (const it of turned) {
      const x = Math.round(it.x)
      byColumn.set(x, (byColumn.get(x) ?? '') + it.str)
    }
    expect(
      [...byColumn.entries()]
        .sort((p, q) => q[0] - p[0])
        .map(([, text]) => text.replace(/\s+/g, '')),
    ).toEqual(['cell', 'one'])
  })
})

test('a Horizontal in Vertical run stands upright in vertical text and reads upwards in horizontal text, fitted or not', async () => {
  test.setTimeout(240_000)
  const tbRl = '<w:textDirection w:val="tbRl"/>'
  const inCell = (label: string, text: string, id: number, fit: boolean) =>
    cell(`<w:p>${run(`${label} Ab `)}${upright(text, id, fit)}${run(' cd')}</w:p>`, 1300, tbRl)
  const body =
    table(
      inCell('c1', '2026', 21, false) +
        inCell('c2', '2026', 22, true) +
        inCell('c3', '123456', 23, true),
      [1300, 1300, 1300],
      '<w:trHeight w:val="2600" w:hRule="atLeast"/>',
    ) +
    para('plain line') +
    `<w:p>${run('h1 Horizontal ')}${upright('2026', 31)}${run(' end')}</w:p>` +
    `<w:p>${run('h2 Horizontal ')}${upright('2026', 32, true)}${run(' end')}</w:p>` +
    sectPr()
  await withDoc('horizontal-in-vertical', body, async ({ launched, page, dir }) => {
    const runs = page.locator('.doc-page [data-ea-vert]')
    await expect(runs).toHaveCount(5, { timeout: 30_000 })
    const drawn = (i: number) =>
      runs.nth(i).evaluate((el) => {
        const r = el.getBoundingClientRect()
        const cs = getComputedStyle(el)
        return {
          width: r.width,
          height: r.height,
          middle: r.top + r.height / 2,
          left: r.left,
          right: r.right,
          mode: cs.writingMode,
          display: cs.display,
          combine: cs.getPropertyValue('text-combine-upright'),
        }
      })
    const [c1, c2, c3, h1, h2] = await Promise.all([0, 1, 2, 3, 4].map(drawn))

    // in a line running down: upright, as wide as its text, the same with and without the fitting
    for (const c of [c1, c2, c3]) {
      expect(c.mode).toBe('horizontal-tb')
      expect(c.display).toBe('inline-block')
      expect(c.combine).toBe('none')
      expect(c.width).toBeGreaterThan(c.height)
    }
    expect(Math.abs(c2.width - c1.width)).toBeLessThan(0.5)
    expect(Math.abs(c2.height - c1.height)).toBeLessThan(0.5)
    expect(c3.width).toBeGreaterThan(c1.width * 1.4)
    // in a horizontal line: turned to read upwards, and the line grows to hold it
    for (const h of [h1, h2]) {
      expect(h.mode).toBe('sideways-lr')
      expect(h.height).toBeGreaterThan(h.width)
      expect(Math.abs(h.height - c1.width)).toBeLessThan(1)
    }
    const plain = await box(page, '.doc-page > p', 'plain line')
    const line = await box(page, '.doc-page > p', 'h1 Horizontal')
    expect(line.height).toBeGreaterThan(plain.height + 8)
    expect(line.height).toBeGreaterThanOrEqual(h1.height)
    expect(Math.abs(h1.middle - (line.top + (line.height - (plain.height - 16)) / 2))).toBeLessThan(
      plain.height,
    )
    await page.screenshot({ path: test.info().outputPath('horizontal-in-vertical.png') })

    const pdf = await exportPdf(launched, page, join(dir, 'horizontal-in-vertical.pdf'))
    const years = pdf.items.filter((it) => it.str === '2026')
    expect(years).toHaveLength(4)
    const flat = years.filter((it) => Math.abs(it.b) < 0.01)
    const up = years.filter((it) => it.b > 0.5)
    // the two in the cells are drawn straight and unsqueezed, the two in the lines a quarter left
    expect(flat).toHaveLength(2)
    expect(up).toHaveLength(2)
    for (const it of flat) expect(Math.abs(it.a)).toBeCloseTo(Math.abs(it.d), 1)
    expect(flat[0].a).toBeCloseTo(flat[1].a, 2)
    const six = pdf.items.find((it) => it.str === '123456')!
    expect(Math.abs(six.b)).toBeLessThan(0.01)
    expect(six.a).toBeCloseTo(flat[0].a, 2)
  })
})
