/**
 * Format > Borders and Shading > Options: the Left and Right distance of a
 * paragraph border move the line out into the margin while the text stays
 * where it was, are saved as w:space on w:pBdr and are read back; a file that
 * brings the same distances draws the same box, and the print sheet and the
 * exported PDF keep it.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import {
  copyFileSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'

const LONG =
  'apple ' +
  'The quick brown fox jumps over the lazy dog and keeps running along the river bank. '.repeat(4)

const side = (name: string, space: number) =>
  `<w:${name} w:val="single" w:sz="4" w:space="${space}" w:color="auto"/>`
const boxed = (text: string) =>
  `<w:p><w:pPr><w:pBdr>${side('top', 1)}${side('left', 8)}${side('bottom', 1)}${side('right', 12)}</w:pBdr></w:pPr>` +
  `<w:r><w:t>${text}</w:t></w:r></w:p>`

interface AidocsWindow {
  __exportPdf?: () => Promise<boolean>
}

async function clickMenu(app: ElectronApplication, path: string[]) {
  const click = () =>
    app.evaluate(({ Menu }, path) => {
      let items = Menu.getApplicationMenu()?.items
      let item: Electron.MenuItem | undefined
      for (const label of path) {
        item = items?.find((candidate) => candidate.label.startsWith(label))
        items = item?.submenu?.items
      }
      if (!item?.enabled) return false
      item.click()
      return true
    }, path)
  await expect.poll(click, { timeout: 15_000 }).toBe(true)
}

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')!.async('string')

const paragraph = (page: Page, text: string, root = '.editor-scroll .doc-page') =>
  page.locator(`${root} p`, { hasText: text }).first()

/** the paragraph box and its text, in unzoomed px from the page's left edge */
const geometry = (page: Page, text: string, root?: string) =>
  paragraph(page, text, root).evaluate((el) => {
    const paper = el.closest('.doc-page') as HTMLElement
    const pr = paper.getBoundingClientRect()
    const zoom = pr.width / paper.offsetWidth
    const range = document.createRange()
    range.selectNodeContents(el)
    const lines = Array.from(range.getClientRects()).filter((r) => r.width > 0)
    const r = el.getBoundingClientRect()
    const round = (n: number) => Math.round((n / zoom) * 10) / 10
    return {
      boxLeft: round(r.left - pr.left),
      boxRight: round(r.right - pr.left),
      textLeft: round(Math.min(...lines.map((l) => l.left)) - pr.left),
      textRight: round(Math.max(...lines.map((l) => l.right)) - pr.left),
      lines: new Set(lines.map((l) => Math.round(l.top))).size,
      height: round(r.height),
      leftLine: getComputedStyle(el).borderLeftStyle,
    }
  })

async function openDialog(app: ElectronApplication, page: Page, text: string) {
  await paragraph(page, text).click()
  await clickMenu(app, ['Format', 'Borders and Shading…'])
  const dialog = page.locator('.borders-shading-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

async function pdfTextLeft(file: string, word: string): Promise<number> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(file)) }).promise
  const content = await (await doc.getPage(1)).getTextContent()
  const item = content.items.find(
    (it): it is { transform: number[]; str: string } =>
      'str' in it && (it as { str: string }).str.startsWith(word),
  )
  return item ? item.transform[4] : Number.NaN
}

test('the left and right distance move the border line, not the text', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-border-side-distance-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para(LONG) + para('banana') + boxed('cherry') + boxed('damson') + para('elder') + SECT_PR,
    ),
  )
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-para-border-side-distance',
    openFile: docPath,
  })
  const PX = 96 / 72
  try {
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(paragraph(page, 'apple')).toBeVisible({ timeout: 30_000 })

    // the file's own distances: the box stands 8 pt and 12 pt outside the text column
    const plain = await geometry(page, 'banana')
    const cherry = await geometry(page, 'cherry')
    expect(cherry.leftLine).toBe('solid')
    expect(cherry.boxLeft).toBeCloseTo(plain.boxLeft - 8 * PX, 1)
    expect(cherry.boxRight).toBeCloseTo(plain.boxRight + 12 * PX, 1)
    expect(cherry.textLeft).toBeCloseTo(plain.textLeft + 1, 1)
    // one border group: no line between the two, and their sides line up
    await expect(paragraph(page, 'cherry')).toHaveClass(/pbdr-suppress-bottom/)
    await expect(paragraph(page, 'damson')).toHaveClass(/pbdr-suppress-top/)
    const damson = await geometry(page, 'damson')
    expect([damson.boxLeft, damson.boxRight]).toEqual([cherry.boxLeft, cherry.boxRight])

    // a box with no side distance, then the distances from Options
    let dialog = await openDialog(launched.app, page, 'apple')
    await dialog.locator('button[data-setting="box"]').click()
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await geometry(page, 'apple')).leftLine).toBe('solid')
    const tight = await geometry(page, 'apple')
    expect(tight.lines).toBeGreaterThan(1)
    expect(tight.boxLeft).toBeCloseTo(plain.boxLeft, 1)

    dialog = await openDialog(launched.app, page, 'apple')
    await dialog.locator('.pb-options-toggle').click()
    const left = dialog.locator('input[name="space-l"]')
    const right = dialog.locator('input[name="space-r"]')
    await expect(left).toBeEnabled()
    await expect(right).toBeEnabled()
    await expect(left).toHaveValue('0')
    // Word's range ends at 31 pt
    await left.fill('40')
    await expect(left).toHaveValue('31')
    await left.fill('8')
    await right.fill('12')
    await page.screenshot({ path: test.info().outputPath('side-distance-dialog.png') })
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toBeHidden()

    const moved = {
      ...tight,
      boxLeft: Math.round((tight.boxLeft - 8 * PX) * 10) / 10,
      boxRight: Math.round((tight.boxRight + 12 * PX) * 10) / 10,
    }
    await expect.poll(() => geometry(page, 'apple')).toEqual(moved)
    // the same picture as the paragraph the file brought
    expect([moved.boxLeft, moved.boxRight]).toEqual([cherry.boxLeft, cherry.boxRight])
    expect(moved.textLeft).toBe(cherry.textLeft)
    await page.screenshot({ path: test.info().outputPath('side-distance-page.png') })

    await saveFromMenu(launched.app)
    const pBdr = `<w:pBdr>${side('top', 0)}${side('left', 8)}${side('bottom', 0)}${side('right', 12)}</w:pBdr>`
    await expect
      .poll(async () => (await documentXml(docPath)).includes(`<w:p><w:pPr>${pBdr}</w:pPr>`), {
        timeout: 20_000,
      })
      .toBe(true)
    // the paragraphs the file brought keep their w:pBdr
    expect((await documentXml(docPath)).split(side('left', 8)).length - 1).toBe(3)

    // the print sheet lays the same box out
    await clickMenu(launched.app, ['File', 'Print'])
    const sheet = page.locator('.print-dialog')
    await expect(sheet).toBeVisible({ timeout: 30_000 })
    const pv = '.pagination-preview .pv-page .doc-page'
    await expect(paragraph(page, 'apple', pv)).toBeAttached()
    const printed = await geometry(page, 'apple', pv)
    expect(printed.lines).toBe(moved.lines)
    expect(printed.textLeft - printed.boxLeft).toBeCloseTo(moved.textLeft - moved.boxLeft, 1)
    expect(printed.boxRight - printed.textLeft).toBeCloseTo(moved.boxRight - moved.textLeft, 1)
    await sheet.locator('button', { hasText: 'Cancel' }).click()
    await expect(sheet).toBeHidden()

    // the exported PDF keeps the text at the margin (72 pt) inside its 0.75 pt line
    const pdf = join(dir, 'letter.pdf')
    await launched.app.evaluate(({ dialog: d, shell }, out) => {
      d.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdf)
    expect(await page.evaluate(() => (window as unknown as AidocsWindow).__exportPdf!())).toBe(true)
    expect(await pdfTextLeft(pdf, 'apple')).toBeCloseTo(72.75, 0)
    expect(await pdfTextLeft(pdf, 'cherry')).toBeCloseTo(72.75, 0)
    expect(await pdfTextLeft(pdf, 'banana')).toBeCloseTo(72, 0)
    copyFileSync(pdf, test.info().outputPath('side-distance.pdf'))

    await closeAndSaveVideo(launched, 'docs-para-border-side-distance')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-para-border-side-distance-reopen',
      openFile: docPath,
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(paragraph(page, 'apple')).toBeVisible({ timeout: 30_000 })
    await expect.poll(() => geometry(page, 'apple')).toEqual(moved)
    dialog = await openDialog(launched.app, page, 'apple')
    await dialog.locator('.pb-options-toggle').click()
    await expect(dialog.locator('input[name="space-l"]')).toHaveValue('8')
    await expect(dialog.locator('input[name="space-r"]')).toHaveValue('12')

    // back to no distance: the line returns to the text column
    await dialog.locator('input[name="space-l"]').fill('0')
    await dialog.locator('input[name="space-r"]').fill('0')
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => geometry(page, 'apple')).toEqual(tight)
  } finally {
    await closeAndSaveVideo(launched, 'docs-para-border-side-distance')
    rmSync(dir, { recursive: true, force: true })
  }
})

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const NUMBERING =
  `<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="${W_NS}">` +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>' +
  '<w:lvlText w:val="-"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'

async function docxWithList(body: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body))
  const edit = async (name: string, from: string, to: string) =>
    zip.file(name, (await zip.file(name)!.async('string')).replace(from, to))
  zip.file('word/numbering.xml', NUMBERING)
  await edit(
    '[Content_Types].xml',
    '</Types>',
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
  )
  await edit(
    'word/_rels/document.xml.rels',
    '</Relationships>',
    '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('an indented, a listed and a right-to-left paragraph keep their text beside the moved line', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-border-side-kinds-')))
  const docPath = join(dir, 'kinds.docx')
  const p = (text: string, pPr: string, spaced: boolean) =>
    `<w:p><w:pPr>${pPr}<w:pBdr>${side('top', 1)}${side('left', spaced ? 8 : 0)}${side('bottom', 1)}${side('right', spaced ? 12 : 0)}</w:pBdr></w:pPr>` +
    `<w:r><w:t>${text}</w:t></w:r></w:p>`
  const kinds: Array<[string, string, string]> = [
    ['fig', 'grape', '<w:ind w:left="720" w:right="360"/>'],
    ['honey', 'iris', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'],
    ['jade', 'kiwi', '<w:bidi/>'],
  ]
  // pPr children in schema order: numPr, pBdr, bidi, ind
  const build = (text: string, pPr: string, spaced: boolean) =>
    pPr.startsWith('<w:numPr')
      ? p(text, pPr, spaced)
      : p(text, '', spaced).replace('</w:pBdr>', `</w:pBdr>${pPr}`)
  writeFileSync(
    docPath,
    await docxWithList(
      kinds
        .map(
          ([tight, spaced, pPr]) =>
            build(tight, pPr, false) + para('-') + build(spaced, pPr, true) + para('-'),
        )
        .join('') + SECT_PR,
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-para-border-side-kinds',
    openFile: docPath,
  })
  const PX = 96 / 72
  const block = (page: Page, text: string) =>
    page.locator('.editor-scroll .doc-page :is(p, .doc-li)', { hasText: text }).first()
  const measure = (page: Page, text: string) =>
    block(page, text).evaluate((el) => {
      const range = document.createRange()
      range.selectNodeContents(el)
      const lines = Array.from(range.getClientRects()).filter((r) => r.width > 0)
      const r = el.getBoundingClientRect()
      const round = (n: number) => Math.round(n * 10) / 10
      return {
        boxLeft: round(r.left),
        boxRight: round(r.right),
        textLeft: round(Math.min(...lines.map((l) => l.left))),
        textRight: round(Math.max(...lines.map((l) => l.right))),
      }
    })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(block(page, 'fig')).toBeVisible({ timeout: 30_000 })
    await expect(block(page, 'iris')).toHaveClass(/doc-li/)
    for (const [tightText, spacedText] of kinds) {
      const tight = await measure(page, tightText)
      const spaced = await measure(page, spacedText)
      expect(spaced.boxLeft, spacedText).toBeCloseTo(tight.boxLeft - 8 * PX, 0)
      expect(spaced.boxRight, spacedText).toBeCloseTo(tight.boxRight + 12 * PX, 0)
      // the first glyphs differ, the edge the text starts from does not
      const start = spacedText === 'kiwi' ? 'textRight' : 'textLeft'
      expect(spaced[start], spacedText).toBeCloseTo(tight[start], 0)
    }
    await page.screenshot({ path: test.info().outputPath('side-distance-kinds.png') })
  } finally {
    await closeAndSaveVideo(launched, 'docs-para-border-side-kinds')
    rmSync(dir, { recursive: true, force: true })
  }
})
