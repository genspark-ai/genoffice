/**
 * Controls that were shown greyed out and now work: the Styles pane's Delete
 * and Organizer…, the Field dialog's Options…, Borders and Shading's
 * Horizontal Line…, the Language dialog's Detect language automatically, and
 * Mailings > Update Labels on a sheet of labels.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-greyed-controls-styles-fields'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const NORMAL =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
const custom = (id: string, rPr: string) =>
  `<w:style w:type="paragraph" w:customStyle="1" w:styleId="${id}"><w:name w:val="${id}"/>` +
  `<w:basedOn w:val="Normal"/><w:qFormat/><w:rPr>${rPr}</w:rPr></w:style>`

/** a document whose styles part holds Normal and the given custom styles */
async function docxWithStyles(body: string, styles: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body))
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>${NORMAL}${styles}</w:styles>`,
  )
  const rels = 'word/_rels/document.xml.rels'
  zip.file(
    rels,
    (await zip.file(rels)!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    ),
  )
  const types = '[Content_Types].xml'
  zip.file(
    types,
    (await zip.file(types)!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** clicks a menu bar item once the renderer has told the menu it is available */
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

/** puts the caret right after `text`, or with `select` selects it */
const caretAt = (page: Page, text: string, select = false) =>
  page.evaluate(
    ([text, select]) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: {
                doc: {
                  descendants(
                    f: (node: { isText: boolean; text?: string }, pos: number) => void,
                  ): void
                }
              }
              chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } }
            }
          }
        }
      ).__aidocs
      let from = -1
      editor.state.doc.descendants((node, pos) => {
        const i = node.isText ? (node.text ?? '').indexOf(text as string) : -1
        if (from < 0 && i >= 0) from = pos + i
      })
      if (from < 0) throw new Error(`no text "${text}"`)
      const to = from + (text as string).length
      editor
        .chain()
        .focus()
        .setTextSelection({ from: select ? from : to, to })
        .run()
    },
    [text, select] as const,
  )

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''
const documentXml = (docPath: string) => part(docPath, 'word/document.xml')
const stylesXml = (docPath: string) => part(docPath, 'word/styles.xml')

test('Styles: Delete takes a style out of the document, Organizer copies styles to and from another file', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-style-organizer-')))
  const docPath = join(dir, 'styles.docx')
  const otherPath = join(dir, 'other.docx')
  writeFileSync(
    docPath,
    await docxWithStyles(
      '<w:p><w:pPr><w:pStyle w:val="MyQuote"/></w:pPr><w:r><w:t>Quoted words</w:t></w:r></w:p>' +
        para('Plain words') +
        SECT_PR,
      custom('MyQuote', '<w:i/>') + custom('Keeper', '<w:b/>'),
    ),
  )
  writeFileSync(
    otherPath,
    await docxWithStyles(para('Other') + SECT_PR, custom('Fancy', '<w:color w:val="FF0000"/>')),
  )
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    page.on('dialog', (d) => void d.accept())
    const quoted = page.locator('.doc-page p', { hasText: 'Quoted words' }).first()
    const plain = page.locator('.doc-page p', { hasText: 'Plain words' }).first()
    await expect(quoted).toBeVisible({ timeout: 30_000 })
    const css = (el: typeof quoted, prop: 'fontStyle' | 'color') =>
      el.evaluate((node, prop) => getComputedStyle(node)[prop], prop)
    await expect.poll(() => css(quoted, 'fontStyle')).toBe('italic')
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, otherPath)

    await caretAt(page, 'Plain')
    await clickMenu(app, ['Format', 'Style…'])
    const pane = page.locator('.styles-pane')
    const del = pane.locator('button[name="deleteStyle"]')
    // a built-in style cannot be deleted
    await expect(del).toBeDisabled()
    await caretAt(page, 'Quoted')
    await expect(del).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('styles-pane-delete.png') })
    await del.click()
    await expect.poll(() => css(quoted, 'fontStyle')).toBe('normal')
    await expect(pane.locator('.style-card[data-style-id="MyQuote"]')).toHaveCount(0)

    await pane.locator('button[name="organizer"]').click()
    const organizer = page.locator('.organizer-dialog')
    const docList = organizer.locator('select[name="docStyles"]')
    const otherList = organizer.locator('select[name="otherStyles"]')
    await expect(docList.locator('option')).toHaveText(['Normal', 'Keeper'])
    await expect(otherList).toBeDisabled()
    await organizer.locator('button[name="file"]').click()
    await expect(otherList.locator('option')).toHaveText(['Normal', 'Fancy'])
    await otherList.selectOption('Fancy')
    await organizer.locator('button[name="copy"]').click()
    await expect(docList.locator('option')).toHaveText(['Normal', 'Keeper', 'Fancy'])
    await docList.selectOption('Keeper')
    await organizer.locator('button[name="copy"]').click()
    await expect(otherList.locator('option')).toHaveText(['Normal', 'Fancy', 'Keeper'])
    await organizer.locator('button[name="rename"]').click()
    await organizer.locator('input[name="newName"]').fill('Keeper Two')
    await organizer.locator('button[name="renameOk"]').click()
    await expect(docList.locator('option')).toHaveText(['Normal', 'Keeper Two', 'Fancy'])
    await page.screenshot({ path: test.info().outputPath('organizer.png') })
    await organizer.locator('button[name="close"]').click()
    await expect(organizer).toBeHidden()
    await expect
      .poll(async () => /w:styleId="Keeper"/.test(await stylesXml(otherPath)), { timeout: 15_000 })
      .toBe(true)

    // the copied style is there to apply
    await caretAt(page, 'Plain')
    await pane.locator('.style-card[data-style-id="Fancy"]').click()
    await expect.poll(() => css(plain, 'color'), { timeout: 15_000 }).toBe('rgb(255, 0, 0)')
    await page.screenshot({ path: test.info().outputPath('styles-after.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const styles = await stylesXml(docPath)
          const body = await documentXml(docPath)
          return [
            styles.includes('MyQuote'),
            /w:styleId="Fancy"/.test(styles),
            styles.includes('<w:name w:val="Keeper Two"/>'),
            body.includes('MyQuote'),
            /<w:pStyle w:val="Fancy"\/>/.test(body),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([false, true, true, false, true])

    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    const plainAgain = reopened.locator('.doc-page p', { hasText: 'Plain words' }).first()
    await expect(plainAgain).toBeVisible({ timeout: 30_000 })
    await expect.poll(() => css(plainAgain, 'color'), { timeout: 15_000 }).toBe('rgb(255, 0, 0)')
    await expect
      .poll(() =>
        css(reopened.locator('.doc-page p', { hasText: 'Quoted words' }).first(), 'fontStyle'),
      )
      .toBe('normal')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Field Options adds switches that shape the result; Horizontal Line goes in; typed text is marked with its language', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-field-options-')))
  const docPath = join(dir, 'fields.docx')
  writeFileSync(
    docPath,
    await docxWith(
      ['Pages: ', 'Line here', 'Typed: ']
        .map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`)
        .join('') + SECT_PR,
    ),
  )
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(paragraph('Pages:')).toBeVisible({ timeout: 30_000 })

    // Field > Options…: only for a field whose result is worked out here
    const field = page.locator('.field-dialog').first()
    const options = page.locator('.field-options-dialog')
    const optionsButton = field.locator('button[name="options"]')
    await caretAt(page, 'Pages: ')
    await clickMenu(app, ['Insert', 'Field…'])
    await field.getByLabel('Field code').first().fill('TOC')
    await expect(optionsButton).toBeDisabled()
    await field.getByLabel('Field code').first().fill('NUMPAGES')
    await expect(optionsButton).toBeEnabled()
    await optionsButton.click()
    const code = options.locator('input[name="optionsCode"]')
    await options.locator('select[name="numeric"]').selectOption('\\# 0.00')
    await options.locator('button[name="add"]').click()
    await expect(code).toHaveValue('NUMPAGES \\# 0.00')
    await options.locator('button[name="undo"]').click()
    await expect(code).toHaveValue('NUMPAGES')
    await options.locator('select[name="format"]').selectOption('\\* ROMAN')
    await options.locator('button[name="add"]').click()
    await expect(code).toHaveValue('NUMPAGES \\* ROMAN')
    await page.screenshot({ path: test.info().outputPath('field-options.png') })
    await options.locator('button.btn-primary').click()
    await expect(options).toBeHidden()
    await expect(field.getByLabel('Field code')).toHaveValue('NUMPAGES \\* ROMAN')
    await field.locator('button.btn-primary').click()
    await expect(paragraph('Pages:')).toHaveText('Pages: I')

    // Borders and Shading > Horizontal Line…
    await caretAt(page, 'Line here')
    await clickMenu(app, ['Format', 'Borders and Shading…'])
    const borders = page.locator('.borders-shading-dialog')
    await page.screenshot({ path: test.info().outputPath('borders-shading.png') })
    await borders.locator('button[name="horizontalLine"]').click()
    await expect(borders).toBeHidden()
    const rule = page.locator('.doc-page .doc-inline-rule')
    await expect(rule).toHaveCount(1)
    const ruleBox = await rule.boundingBox()
    const lineBox = await paragraph('Line here').boundingBox()
    expect(ruleBox!.y).toBeGreaterThan(lineBox!.y)
    expect(ruleBox!.width).toBeGreaterThan(400)

    // Language > Detect language automatically
    const language = page.locator('.language-dialog')
    const detect = language.locator('input[name="detect"]')
    await caretAt(page, 'Typed:', true)
    await clickMenu(app, ['Tools', 'Language…'])
    await expect(detect).not.toBeChecked()
    await detect.check()
    await page.screenshot({ path: test.info().outputPath('language-detect.png') })
    await language.locator('button.primary').click()
    await expect(language).toBeHidden()
    await caretAt(page, 'Typed: ')
    // the editor takes focus a frame later: keys typed before that go nowhere
    await expect
      .poll(() => page.evaluate(() => !!document.activeElement?.closest('.ProseMirror')))
      .toBe(true)
    await page.keyboard.type('le chat est dans la maison et il ne veut pas sortir avec nous')
    await expect(paragraph('Typed:')).toContainText('sortir avec nous')
    await caretAt(page, 'maison', true)
    await clickMenu(app, ['Tools', 'Language…'])
    await expect(detect).toBeChecked()
    await expect(language.locator('select.language-list')).toHaveValue('fr-FR')
    await detect.uncheck()
    await language.locator('button.primary').click()
    await page.screenshot({ path: test.info().outputPath('field-line-language.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            /NUMPAGES \\\* ROMAN/.test(xml),
            /<v:rect\b[^>]*o:hr="t"/.test(xml),
            /<w:lang w:val="fr-FR"\/>/.test(xml),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true])
    // the page count is known here: the field is saved with it, not left for Word to update
    expect(await documentXml(docPath)).toMatch(
      /<w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText[^>]*> NUMPAGES \\\* ROMAN [^<]*<\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t[^>]*>I</,
    )

    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(reopened.locator('.doc-page p', { hasText: 'Pages:' }).first()).toHaveText(
      'Pages: I',
      { timeout: 30_000 },
    )
    await expect(reopened.locator('.doc-page .doc-inline-rule')).toHaveCount(1)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Update Labels copies the first label of a label sheet into the others; the merge fills them record by record', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-update-labels-')))
  const docPath = join(dir, 'labels.docx')
  writeFileSync(docPath, await docxWith(para('Dear') + SECT_PR))
  const listPath = join(dir, 'people.csv')
  writeFileSync(listPath, 'Name\nAda\nGrace\nAlan\nLinus\nBo\n')
  const outPath = join(dir, 'merged.docx')
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const greeting = page.locator('.doc-page p', { hasText: 'Dear' }).first()
    await expect(greeting).toBeVisible({ timeout: 30_000 })
    await app.evaluate(
      ({ dialog }, [list, out]) => {
        dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [list] })
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
      },
      [listPath, outPath],
    )
    await page.locator('.ribbon-tab', { hasText: /^Mailings$/ }).click()
    const big = (name: string) =>
      page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })
    await expect(big('Update Labels')).toBeDisabled()

    await greeting.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ')
    const panel = page.locator('.modal-mailmerge')
    await clickMenu(app, ['Tools', 'Labels…'])
    await panel.getByRole('button', { name: 'Select Recipient List…' }).click()
    await panel.locator('.mailmerge-field', { hasText: /^Name$/ }).click()
    await panel.locator('input[name="mailmerge-stock"]').nth(1).check()
    const grid = panel.locator('.mailmerge-pattern input[type="number"]')
    await grid.nth(0).fill('2')
    await grid.nth(1).fill('3')
    await page.screenshot({ path: test.info().outputPath('labels-panel.png') })
    await panel.locator('button[name="createLabelSheet"]').click()
    await expect(panel).toBeHidden()

    const labels = page.locator('.doc-page td', { hasText: '«Name»' })
    await expect(labels).toHaveCount(6)
    await expect(big('Update Labels')).toBeEnabled()

    // the first label changes; Update Labels brings the others in step
    await labels.first().locator('p').first().click()
    await page.keyboard.press('Home')
    await page.keyboard.type('To ')
    await expect(page.locator('.doc-page td', { hasText: 'To Dear' })).toHaveCount(1)
    await big('Update Labels').click()
    await expect(page.locator('.doc-page td', { hasText: 'To Dear' })).toHaveCount(6)
    await page.screenshot({ path: test.info().outputPath('labels-updated.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            (xml.match(/<w:instrText[^>]*>\s*NEXT\s*<\/w:instrText>/g) ?? []).length,
            (xml.match(/MERGEFIELD/g) ?? []).length,
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([5, 6])
    // a label sheet prints no grid: Word draws the table's lines over a cell's "none", not over "nil"
    const borderless = (xml: string) => [
      (xml.match(/<w:tcBorders><w:top w:val="nil"\/><w:left w:val="nil"\/>/g) ?? []).length,
      /<w:tcBorders>(?:(?!<\/w:tcBorders>).)*w:val="none"/.test(xml),
    ]
    expect(borderless(await documentXml(docPath))).toEqual([6, false])

    await big('Finish & Merge').click()
    await panel.locator('input[name="mailmerge-output"]').nth(2).check()
    await panel.locator('button.btn-primary').click()
    await expect(panel.locator('.mailmerge-done')).toBeVisible({ timeout: 30_000 })
    const merged = await documentXml(outPath)
    expect(merged.match(/Ada|Grace|Alan|Linus|Bo\b/g)).toEqual([
      'Ada',
      'Grace',
      'Alan',
      'Linus',
      'Bo',
    ])
    expect((merged.match(/To Dear/g) ?? []).length).toBe(6)
    expect(merged).not.toContain('NEXT')
    expect(borderless(merged)).toEqual([6, false])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
