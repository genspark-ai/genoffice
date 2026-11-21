/**
 * Home > Text Effects is a gallery of the ready-made looks with the way into
 * the dialog under them. A look's colour is saved as Word saves a text fill
 * (w14:textFill beside w:color), and a fill that came with a file is drawn,
 * shown in the dialog and left as it is by a save that does not touch it.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-text-effects-gallery'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'

async function docxWith(body: string): Promise<Buffer> {
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
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const run = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
const solidFill = (hex: string) =>
  `<w14:textFill><w14:solidFill><w14:srgbClr w14:val="${hex}"/></w14:solidFill></w14:textFill>`

async function clickMenu(app: ElectronApplication, path: string[]) {
  const click = () =>
    app.evaluate(({ Menu }, path) => {
      let items = Menu.getApplicationMenu()?.items
      let item: Electron.MenuItem | undefined
      for (const label of path) {
        item = items?.find((candidate) => candidate.label === label)
        items = item?.submenu?.items
      }
      if (!item?.enabled) return false
      item.click()
      return true
    }, path)
  await expect.poll(click, { timeout: 15_000 }).toBe(true)
}

const documentXml = async (docPath: string) =>
  (await (
    await JSZip.loadAsync(readFileSync(docPath))
  )
    .file('word/document.xml')
    ?.async('string')) ?? ''

/** selects document positions: the first paragraph's text starts at 1 */
const select = (page: Page, from: number, to: number) =>
  page.evaluate(
    ([from, to]) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection({ from, to }).run()
    },
    [from, to],
  )

/** how the page draws the piece of text that reads `text` */
const drawn = (page: Page, text: string) =>
  page.evaluate((text) => {
    const span = Array.from(document.querySelectorAll<HTMLElement>('.doc-page p span')).find(
      (el) => el.textContent === text,
    )
    if (!span) return null
    const style = getComputedStyle(span)
    return { color: style.color, stroke: style.getPropertyValue('-webkit-text-stroke-color') }
  }, text)

const openDoc = async (docPath: string, text: string) => {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p', { hasText: text })).toBeVisible({ timeout: 30_000 })
  return { launched, page }
}

const ORANGE = 'rgb(237, 125, 49)'
const GREEN = 'rgb(112, 173, 71)'

test('Home > Text Effects: a look from the gallery is drawn and saved as a text fill; the last entry opens the dialog', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-fx-gallery-')))
  const docPath = join(dir, 'gallery.docx')
  writeFileSync(
    docPath,
    await docxWith(`<w:p>${run('Plain fancy words')}</w:p><w:p>${run('Last')}</w:p>`),
  )
  let { launched, page } = await openDoc(docPath, 'Plain fancy words')
  try {
    const { app } = launched
    const button = page.locator('[data-ribbon-body] button[aria-label="Text Effects"]')
    const gallery = page.locator('.text-fx-gallery')
    const dialog = page.locator('.text-fx-dialog')

    await select(page, 7, 12)
    await button.click()
    await expect(gallery.locator('.text-fx-tile')).toHaveCount(6)
    await expect(gallery.locator('.text-fx-more')).toHaveText('Text Effects…')
    await page.screenshot({ path: test.info().outputPath('text-effects-gallery.png') })
    await gallery.locator('.text-fx-tile[data-preset="orange-outline"]').click()
    await expect(gallery).toBeHidden()
    await expect.poll(() => drawn(page, 'fancy')).toEqual({ color: ORANGE, stroke: ORANGE })
    await page.screenshot({ path: test.info().outputPath('text-effects-gallery-page.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            xml.includes('<w:color w:val="ED7D31"/>'),
            // the fill closes the run's properties, after the outline
            xml.includes(`</w14:textOutline>${solidFill('ED7D31')}</w:rPr>`),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true])

    // the entry under the looks opens the dialog, which shows the fill
    await select(page, 7, 12)
    await button.click()
    await gallery.locator('.text-fx-more').click()
    await expect(gallery).toBeHidden()
    await expect(dialog).toBeVisible()
    const fill = dialog.locator('[data-fx="fill"] .text-fx-swatch')
    await expect(fill.and(dialog.locator('[aria-checked="true"]'))).toHaveAttribute(
      'aria-label',
      '#ED7D31',
    )
    await expect(dialog.locator('[data-fx="outline"] .text-fx-choice.active')).toHaveText('0.75 pt')
    await fill.and(dialog.locator('[aria-label="#70AD47"]')).click()
    await page.screenshot({ path: test.info().outputPath('text-effects-fill-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => drawn(page, 'fancy')).toEqual({ color: GREEN, stroke: ORANGE })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            xml.includes('<w:color w:val="70AD47"/>'),
            xml.includes(`</w14:textOutline>${solidFill('70AD47')}</w:rPr>`),
            xml.includes(solidFill('ED7D31')),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, false])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  ;({ launched, page } = await openDoc(docPath, 'Plain fancy words'))
  try {
    await expect.poll(() => drawn(page, 'fancy')).toEqual({ color: GREEN, stroke: ORANGE })
    await select(page, 7, 12)
    await clickMenu(launched.app, ['Format', 'Text Effects…'])
    const dialog = page.locator('.text-fx-dialog')
    await expect(
      dialog.locator('[data-fx="fill"] .text-fx-swatch[aria-checked="true"]'),
    ).toHaveAttribute('aria-label', '#70AD47')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a text fill that came with a file is drawn, shown in the dialog and kept by a save', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-fx-fill-')))
  const docPath = join(dir, 'fill.docx')
  // the fill alone, without the w:color Word writes beside it
  const filled = run('filled', solidFill('C00000'))
  writeFileSync(
    docPath,
    await docxWith(`<w:p>${run('A ')}${filled}${run(' word')}</w:p><w:p>${run('Last')}</w:p>`),
  )
  const { launched, page } = await openDoc(docPath, 'A filled word')
  try {
    const { app } = launched
    await expect.poll(async () => (await drawn(page, 'filled'))?.color).toBe('rgb(192, 0, 0)')

    await select(page, 3, 9)
    await clickMenu(app, ['Format', 'Text Effects…'])
    const dialog = page.locator('.text-fx-dialog')
    await expect(dialog).toBeVisible()
    // a colour the palette does not hold is offered beside it, picked
    const fill = dialog.locator('[data-fx="fill"] .text-fx-swatch')
    await expect(fill).toHaveCount(10)
    await expect(
      dialog.locator('[data-fx="fill"] .text-fx-swatch[aria-checked="true"]'),
    ).toHaveAttribute('aria-label', '#C00000')
    await page.screenshot({ path: test.info().outputPath('text-effects-file-fill.png') })

    // a glow is added; the fill of the file is not touched
    await dialog.locator('[data-fx="glow"] .text-fx-choice', { hasText: /^5 pt$/ }).click()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await drawn(page, 'filled'))?.color).toBe('rgb(192, 0, 0)')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            xml.includes('<w14:glow w14:rad="63500">'),
            xml.includes(`</w14:glow>${solidFill('C00000')}</w:rPr>`),
            xml.includes('<w:color'),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, false])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
