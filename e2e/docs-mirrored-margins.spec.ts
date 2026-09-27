import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Layout > Margins > Mirrored (w:mirrorMargins): w:left is the inside margin
 * and w:right the outside one. Odd pages keep them as left/right, even pages
 * swap them — on the canvas, on the ruler (which follows the caret's page),
 * in Custom Margins (Inside/Outside + Multiple pages) and in the exported PDF.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown }
  __exportPdf?: () => Promise<boolean>
}

const PT_PER_TWIP = 1 / 20
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

async function twoPageDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>',
  )
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>Running header</w:t></w:r></w:p></w:hdr>`,
  )
  zip.file(
    'word/footer1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="${W_NS}"><w:p><w:r><w:t>Running footer</w:t></w:r></w:p></w:ftr>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
      '<w:p><w:r><w:t>Alpha on the first page</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Beta on the second page</w:t></w:r></w:p>' +
      '<w:sectPr><w:headerReference w:type="default" r:id="rId2"/><w:footerReference w:type="default" r:id="rId3"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** x (pt) of the first text item on each page of a PDF */
async function pdfTextLefts(path: string): Promise<number[]> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(path)) }).promise
  const lefts: number[] = []
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i)
    const content = await page.getTextContent()
    const items = content.items.filter(
      (it): it is { transform: number[]; str: string } =>
        'transform' in it && 'str' in it && (it as { str: string }).str.trim().length > 0,
    )
    lefts.push(Math.min(...items.map((it) => it.transform[4])))
  }
  return lefts
}

const colShift = (page: Page, text: string) =>
  page
    .locator('.doc-page p', { hasText: text })
    .first()
    .evaluate((el) => parseFloat(el.style.getPropertyValue('--col-dx')) || 0)

/** left edge of a header/footer strip relative to the paper, in unzoomed px */
const stripInset = (page: Page, selector: string) =>
  page
    .locator(selector)
    .first()
    .evaluate((el) => {
      const paper = el.closest('.page-wrap')!.querySelector('.doc-page') as HTMLElement
      const pr = paper.getBoundingClientRect()
      return (el.getBoundingClientRect().left - pr.left) / (pr.width / paper.offsetWidth)
    })

const leftRulerZone = (page: Page) =>
  page
    .locator('.ruler-zone')
    .first()
    .evaluate((el) => parseFloat(el.style.width))

/** Word for Mac model: clicking the active tab collapses the ribbon, so only switch when needed */
async function showTab(page: Page, name: RegExp) {
  const tab = page.locator('.ribbon-tab', { hasText: name })
  if (!/\bactive\b/.test((await tab.getAttribute('class')) ?? '')) await tab.click()
}

test.describe('docs mirrored margins', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-mirror-')))
    docPath = join(dir, 'mirror.docx')
    writeFileSync(docPath, await twoPageDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('even pages swap inside/outside on the canvas, ruler, dialog and export', async () => {
    test.setTimeout(180_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-mirrored-margins',
      openFile: docPath,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await page.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      const alpha = page.locator('.doc-page p', { hasText: 'Alpha' }).first()
      const beta = page.locator('.doc-page p', { hasText: 'Beta' }).first()
      await expect(beta).toBeVisible()
      await expect(page.locator('.doc-page .page-gap')).toHaveCount(1)

      await showTab(page, /^View$/)
      await page.getByRole('button', { name: 'Ruler', exact: true }).click()
      await expect(page.locator('.ruler')).toBeVisible()

      // Mirrored preset: inside 1.25in (1800), outside 1in (1440)
      await showTab(page, /^Layout$/)
      await page.getByRole('button', { name: 'Margins', exact: true }).click()
      await page.locator('[data-rb-panel] button', { hasText: 'Mirrored' }).click()
      await expect
        .poll(() =>
          page
            .locator('.doc-page')
            .first()
            .evaluate((el) => parseFloat(getComputedStyle(el).paddingLeft)),
        )
        .toBe(120)
      // the even page's blocks translate by outside - inside
      await expect.poll(() => colShift(page, 'Beta')).toBe(-24)
      expect(await colShift(page, 'Alpha')).toBe(0)
      const zoom = await page
        .locator('.doc-page')
        .first()
        .evaluate((el) => el.getBoundingClientRect().width / (el as HTMLElement).offsetWidth)
      const alphaLeft = (await alpha.boundingBox())!.x
      const betaLeft = (await beta.boundingBox())!.x
      expect(betaLeft - alphaLeft).toBeCloseTo(-24 * zoom, 0)

      // header/footer strips sit on their page's own margin: 1.25in on page 1, 1in on page 2
      await expect
        .poll(() => stripInset(page, '.page-wrap .page-hf-header:not(.page-gap-hf)'))
        .toBeCloseTo(120, 0)
      await expect
        .poll(() => stripInset(page, '.page-wrap .page-gap-hf[data-hf-kind="footer"]'))
        .toBeCloseTo(120, 0)
      await expect
        .poll(() => stripInset(page, '.page-wrap .page-gap-hf[data-hf-kind="header"]'))
        .toBeCloseTo(96, 0)
      await expect
        .poll(() => stripInset(page, '.page-wrap .page-hf-footer:not(.page-gap-hf)'))
        .toBeCloseTo(96, 0)

      // the ruler follows the caret's page: left margin 1.25in on page 1, 1in on page 2
      await alpha.click()
      await expect.poll(() => leftRulerZone(page)).toBe(120)
      await beta.click()
      await expect.poll(() => leftRulerZone(page)).toBe(96)
      await alpha.click()
      await expect.poll(() => leftRulerZone(page)).toBe(120)

      // the gallery marks the preset, Custom Margins reads Inside/Outside and Mirror margins
      await page.getByRole('button', { name: 'Margins', exact: true }).click()
      await expect(page.locator('[data-rb-panel] button.active b')).toHaveText('Mirrored')
      await page.getByRole('button', { name: 'Custom Margins…' }).click()
      const dialog = page.locator('.modal', { hasText: 'Custom Page Margins' })
      await expect(dialog.getByRole('textbox', { name: 'Inside', exact: true })).toHaveValue(
        '1.25"',
      )
      await expect(dialog.getByRole('textbox', { name: 'Outside', exact: true })).toHaveValue('1"')
      const pages = dialog.getByRole('combobox')
      await expect(pages).toHaveValue('mirror')
      await pages.selectOption('normal')
      await expect(dialog.getByRole('textbox', { name: 'Left', exact: true })).toHaveValue('1.25"')
      await pages.selectOption('mirror')
      await dialog.getByRole('button', { name: 'Cancel' }).click()

      // export goes through the per-page sheets: page 2's text starts at the outside margin
      await launched.app.evaluate(
        ({ dialog: d, shell }, out) => {
          d.showSaveDialog = async () => ({ canceled: false, filePath: out })
          shell.showItemInFolder = () => {}
        },
        join(dir, 'mirror.pdf'),
      )
      expect(await page.evaluate(() => (window as unknown as AidocsWindow).__exportPdf!())).toBe(
        true,
      )
      const lefts = await pdfTextLefts(join(dir, 'mirror.pdf'))
      expect(lefts).toHaveLength(2)
      expect(lefts[0]).toBeCloseTo(1800 * PT_PER_TWIP, 0)
      expect(lefts[1]).toBeCloseTo(1440 * PT_PER_TWIP, 0)

      // the exported PDF opens in its own tab: return to the document, close the
      // preview the export opened, then Normal from the dialog clears the shift
      const shellPage = await waitForPageWithUrl(launched.app, 'shell/out')
      await shellPage.locator('.tab-bar .tab-item', { hasText: 'mirror.docx' }).click()
      await expect(page.locator('.pagination-preview')).toBeVisible()
      await page.keyboard.press('Escape')
      await expect(page.locator('.pagination-preview')).toBeHidden()
      await showTab(page, /^Layout$/)
      await page.getByRole('button', { name: 'Margins', exact: true }).click()
      await page.getByRole('button', { name: 'Custom Margins…' }).click()
      await dialog.getByRole('combobox').selectOption('normal')
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await expect.poll(() => colShift(page, 'Beta')).toBe(0)
      await beta.click()
      await expect.poll(() => leftRulerZone(page)).toBe(120)
    } finally {
      await closeAndSaveVideo(launched, 'docs-mirrored-margins')
    }
  })
})
