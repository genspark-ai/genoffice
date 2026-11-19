import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import type { Page } from '@playwright/test'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Measurement unit preference (Word ▸ Preferences ▸ General): the ruler
 * graduation, the Custom Margins dialog and the Layout indent fields all
 * follow one unit; fields accept any unit suffix on input.
 */

interface AidocsWindow {
  __aidocs?: { editor?: unknown }
}

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
    '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Units paragraph.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const rulerLabels = (page: Page) =>
  page.locator('.ruler-num').evaluateAll((els) => els.map((el) => el.textContent))

test.describe('docs measurement unit preference', () => {
  let dir: string
  let docPath: string

  test.beforeEach(async () => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-units-')))
    docPath = join(dir, 'units.docx')
    writeFileSync(docPath, await minimalDocx())
  })

  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('ruler, margin dialog and indent fields follow the chosen unit', async () => {
    test.setTimeout(150_000)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-measurement-units',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.waitForFunction(
        () => Boolean((window as unknown as AidocsWindow).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      const openPreferences = () =>
        app.evaluate(({ webContents }) => {
          for (const wc of webContents.getAllWebContents())
            if (wc.getURL().includes('://docs/')) wc.send('menu:command', 'preferences')
        })
      const pickUnit = async (unit: string) => {
        await openPreferences()
        const dialog = page.getByRole('dialog', { name: 'Preferences' })
        await expect(dialog).toBeVisible()
        await dialog.getByRole('combobox').selectOption(unit)
        await dialog.getByRole('button', { name: 'Close' }).click()
        await expect(dialog).toBeHidden()
      }

      await page.locator('.ribbon-tab', { hasText: /^View$/ }).click()
      await page.getByRole('button', { name: 'Ruler', exact: true }).click()
      await expect(page.locator('.ruler')).toBeVisible()

      // inches: numbered from the left margin, one label per inch on both sides
      await pickUnit('in')
      await expect.poll(() => rulerLabels(page)).toEqual(['1', '2', '3', '4', '5', '6', '7', '1'])

      // centimeters: 19 labels across the body and right margin, 2 in the left margin
      await pickUnit('cm')
      await expect.poll(async () => (await rulerLabels(page)).length).toBe(21)
      const labels = await rulerLabels(page)
      expect(labels.slice(0, 3)).toEqual(['1', '2', '3'])
      expect(labels.slice(-2)).toEqual(['1', '2'])

      // Custom Margins shows cm and accepts an inch entry
      await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
      await page.getByRole('button', { name: 'Margins', exact: true }).click()
      await page.getByRole('button', { name: 'Custom Margins…' }).click()
      const top = page.getByRole('textbox', { name: 'Top', exact: true })
      await expect(top).toHaveValue('2.54 cm')
      await top.fill('2 cm')
      await top.press('Tab')
      await expect(top).toHaveValue('2 cm')
      await page.getByRole('button', { name: 'OK', exact: true }).click()
      await expect
        .poll(() =>
          page
            .locator('.doc-page')
            .first()
            .evaluate((el) => parseFloat(getComputedStyle(el).paddingTop)),
        )
        .toBeCloseTo(75.6, 0)

      // Layout indent field: unit display, any-unit input, negative allowed
      await page.locator('.doc-page p', { hasText: 'Units paragraph' }).first().click()
      const indentLeft = page.getByRole('textbox', { name: 'Indent Left' })
      await expect(indentLeft).toHaveValue('0 cm')
      await indentLeft.fill('1"')
      await indentLeft.press('Enter')
      const para = page.locator('.doc-page p', { hasText: 'Units paragraph' }).first()
      await expect(para).toHaveCSS('margin-inline-start', '96px')
      await expect(page.getByRole('textbox', { name: 'Indent Left' })).toHaveValue('2.54 cm')

      // points: the same field re-renders in pt
      await pickUnit('pt')
      await expect(page.getByRole('textbox', { name: 'Indent Left' })).toHaveValue('72 pt')
      await expect.poll(async () => (await rulerLabels(page))[0]).toBe('36')
    } finally {
      await closeAndSaveVideo(launched, 'docs-measurement-units')
    }
  })
})
