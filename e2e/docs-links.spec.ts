/**
 * Linked content and Edit > Links…: Insert > File with Link to File writes an
 * INCLUDETEXT field whose result is what was inserted; a picture inserted
 * linked names its file beside its own bytes; the Links dialog lists both,
 * reads their sources again, locks them, and breaks them; link fields that
 * came with the file are listed and can be locked and broken.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, pictureRun, saveFromMenu } from './docs-objects-fixture'

const PNG_WIDE =
  'iVBORw0KGgoAAAANSUhEUgAAACgAAAAUCAIAAABwJOjsAAAAJ0lEQVR42mM8ISfHMBCAiWGAwKjFoxaPWjxq8ajFoxaPWjxqMf0BAF+MASwgP5VpAAAAAElFTkSuQmCC'
const PNG_SQUARE =
  'iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAIAAAAC64paAAAAHUlEQVR42mOUkzvBQC5gYqAAjGoe1TyqeVQz5ZoBOL0BLMMULg8AAAAASUVORK5CYII='

/** the pictures of the page; the editor's own zero-size separators are images too */
const PICTURES = '.doc-page img:not(.ProseMirror-separator)'

type Item = { label: string; enabled: boolean; submenu?: Item[] }

const menuItem = (app: ElectronApplication, path: string[]): Promise<Item | undefined> =>
  app.evaluate(({ Menu }, trail) => {
    let items = Menu.getApplicationMenu()?.items ?? []
    let found: Electron.MenuItem | undefined
    for (const label of trail) {
      found = items.find((item) => item.label === label)
      items = found?.submenu?.items ?? []
    }
    return found ? { label: found.label, enabled: found.enabled } : undefined
  }, path)

const clickMenu = (app: ElectronApplication, path: string[]) =>
  app.evaluate(({ Menu }, trail) => {
    let items = Menu.getApplicationMenu()?.items ?? []
    let found: Electron.MenuItem | undefined
    for (const label of trail) {
      found = items.find((item) => item.label === label)
      items = found?.submenu?.items ?? []
    }
    found?.click()
  }, path)

const pickFile = (app: ElectronApplication, file: string) =>
  app.evaluate(({ dialog }, picked) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [picked] })
  }, file)

const documentXml = async (file: string): Promise<string> =>
  (await (await JSZip.loadAsync(readFileSync(file))).file('word/document.xml')?.async('string')) ??
  ''
const relsXml = async (file: string): Promise<string> =>
  (await (
    await JSZip.loadAsync(readFileSync(file))
  )
    .file('word/_rels/document.xml.rels')
    ?.async('string')) ?? ''

const fieldChars = (xml: string): string =>
  [...xml.matchAll(/w:fldCharType="(begin|separate|end)"/g)].map((m) => m[1][0]).join('')

async function openLinks(app: ElectronApplication, page: Page) {
  await expect.poll(async () => (await menuItem(app, ['Edit', 'Links…']))?.enabled).toBe(true)
  await clickMenu(app, ['Edit', 'Links…'])
  const dialog = page.locator('.links-dialog')
  await expect(dialog).toBeVisible({ timeout: 15_000 })
  return dialog
}

const caretAtEndOf = async (page: Page, text: string) => {
  await page.locator('.doc-page p', { hasText: text }).first().click()
  await page.keyboard.press('End')
}

test('a file inserted with Link to File is an INCLUDETEXT field the Links dialog updates, locks and breaks', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-links-')))
  const docPath = join(dir, 'host.docx')
  const sourcePath = join(dir, 'source.docx')
  writeFileSync(docPath, await docxWith(para('Host first') + para('Host last')))
  writeFileSync(sourcePath, await docxWith(para('Linked alpha') + para('Linked beta')))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-links',
    openFile: docPath,
  })
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Host first' })).toBeVisible({
      timeout: 30_000,
    })
    // nothing is linked yet: Word greys the item
    await expect.poll(async () => (await menuItem(app, ['Edit', 'Links…']))?.enabled).toBe(false)

    await caretAtEndOf(page, 'Host first')
    await pickFile(app, sourcePath)
    await clickMenu(app, ['Insert', 'File…'])
    const options = page.locator('.insert-file-range-dialog')
    await expect(options).toBeVisible({ timeout: 15_000 })
    await options.locator('input[name="link"]').check()
    await page.screenshot({ path: test.info().outputPath('insert-file-link.png') })
    await options.locator('button.primary').click()
    const body = page.locator('.doc-page')
    await expect(body).toContainText('Linked beta')

    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('INCLUDETEXT')
    let xml = await documentXml(docPath)
    expect(xml).toContain(`INCLUDETEXT "${sourcePath}"`)
    expect(fieldChars(xml)).toBe('bse')
    // the result stands between the separator and the end
    const order = ['"separate"', 'Linked alpha', 'Linked beta', '"end"', 'Host last'].map((s) =>
      xml.indexOf(s),
    )
    expect([...order].sort((a, b) => a - b)).toEqual(order)
    expect(order[0]).toBeGreaterThan(0)

    // the source changes on disk; Update Now brings the new text in
    writeFileSync(sourcePath, await docxWith(para('Linked gamma')))
    let dialog = await openLinks(app, page)
    const rows = dialog.locator('button.links-row')
    await expect(rows).toHaveCount(1)
    await expect(rows.first()).toContainText('source.docx')
    await expect(rows.first()).toContainText('Document')
    await expect(rows.first()).toContainText('Manual')
    await expect(dialog.locator('.links-source')).toHaveText(sourcePath)
    // an INCLUDETEXT field has no automatic update of its own
    await expect(dialog.locator('input[name="links-update"]').first()).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('links-dialog.png') })
    await dialog.locator('button[name="update"]').click()
    await expect(rows.first()).toContainText('Updated')
    await expect(body).toContainText('Linked gamma')
    await expect(body).not.toContainText('Linked alpha')
    await expect(body).toContainText('Host last')

    // locked: the result stands
    await dialog.locator('input[name="links-locked"]').check()
    writeFileSync(sourcePath, await docxWith(para('Linked delta')))
    await dialog.locator('button[name="update"]').click()
    await expect(rows.first()).toContainText('Locked, not updated')
    await expect(body).toContainText('Linked gamma')
    await dialog.locator('input[name="links-locked"]').uncheck()

    // a source that is gone leaves the result and says so
    rmSync(sourcePath)
    await dialog.locator('button[name="update"]').click()
    await expect(rows.first()).toContainText('Source not found')
    await expect(body).toContainText('Linked gamma')
    await dialog.locator('button[name="close"]').click()
    await expect(dialog).toHaveCount(0)

    // F9 with the caret in the linked text reads the source again
    writeFileSync(sourcePath, await docxWith(para('Linked delta')))
    await page.locator('.doc-page p', { hasText: 'Linked gamma' }).first().click()
    await page.keyboard.press('F9')
    await expect(body).toContainText('Linked delta')
    await expect(body).not.toContainText('Linked gamma')

    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('Linked delta')
    xml = await documentXml(docPath)
    expect(fieldChars(xml)).toBe('bse')
    expect(xml).not.toContain('Linked gamma')

    // reopened: still a link, which Break Link turns into its text
    await closeAndSaveVideo(launched, 'docs-links')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-links',
      openFile: docPath,
    })
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page')).toContainText('Linked delta', { timeout: 30_000 })
    dialog = await openLinks(app, page)
    await expect(dialog.locator('button.links-row')).toHaveCount(1)
    await dialog.locator('input[name="links-locked"]').check()
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('w:fldLock="true"')
    page.once('dialog', (confirm) => void confirm.accept())
    await dialog.locator('button[name="break"]').click()
    await expect(dialog.locator('button.links-row')).toHaveCount(0)
    await dialog.locator('button[name="close"]').click()
    await expect.poll(async () => (await menuItem(app, ['Edit', 'Links…']))?.enabled).toBe(false)
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).not.toContain('INCLUDETEXT')
    xml = await documentXml(docPath)
    expect(fieldChars(xml)).toBe('')
    expect(xml).toContain('Linked delta')
  } finally {
    await closeAndSaveVideo(launched, 'docs-links')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a picture inserted linked names its file, is read again by Update Now and keeps its bytes when the link is broken', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-links-')))
  const docPath = join(dir, 'host.docx')
  const picturePath = join(dir, 'photo.png')
  writeFileSync(docPath, await docxWith(para('Host first') + para('Host last')))
  writeFileSync(picturePath, Buffer.from(PNG_WIDE, 'base64'))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-links',
    openFile: docPath,
  })
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Host first' })).toBeVisible({
      timeout: 30_000,
    })
    await caretAtEndOf(page, 'Host first')
    await pickFile(app, picturePath)
    // Link to File is an option of Picture from File, as in Word: the menu has no entry of its own
    const pictureItems = await app.evaluate(({ Menu }) => {
      const insert = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Insert')
      const pictures = insert?.submenu?.items.find((item) => item.label === 'Pictures')
      return pictures?.submenu?.items.map((item) => item.label)
    })
    expect(pictureItems).toEqual(['Picture from File…'])
    await clickMenu(app, ['Insert', 'Pictures', 'Picture from File…'])
    const options = page.locator('.insert-picture-options-dialog')
    await expect(options).toContainText('photo.png')
    await expect(options.locator('input[name="link"]')).not.toBeChecked()
    await options.locator('input[name="link"]').check()
    await page.screenshot({ path: test.info().outputPath('picture-options.png') })
    await options.locator('button[type="submit"]').click()
    await expect(options).toBeHidden()
    const img = page.locator(PICTURES)
    await expect(img).toHaveCount(1)
    await expect(img).toHaveAttribute('src', new RegExp(PNG_WIDE.slice(0, 40)))

    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('r:link=')
    const blip = /<a:blip\b[^>]*>/.exec(await documentXml(docPath))![0]
    const embed = /r:embed="([^"]+)"/.exec(blip)?.[1]
    const link = /r:link="([^"]+)"/.exec(blip)?.[1]
    expect(embed).toBeTruthy()
    expect(link).toBeTruthy()
    const rels = await relsXml(docPath)
    const target = pathToFileURL(picturePath).href
    expect(rels).toMatch(
      new RegExp(`<Relationship Id="${link}"[^>]*Target="${target}"[^>]*TargetMode="External"`),
    )
    expect(rels).toMatch(new RegExp(`<Relationship Id="${embed}"[^>]*Target="media/`))

    // reopened: the picture shows from its own bytes and is listed as a link
    await closeAndSaveVideo(launched, 'docs-links')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-links',
      openFile: docPath,
    })
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator(PICTURES)).toHaveCount(1, { timeout: 30_000 })
    writeFileSync(picturePath, Buffer.from(PNG_SQUARE, 'base64'))
    const dialog = await openLinks(app, page)
    const rows = dialog.locator('button.links-row')
    await expect(rows).toHaveCount(1)
    await expect(rows.first()).toContainText('photo.png')
    await expect(rows.first()).toContainText('Picture')
    await expect(dialog.locator('input[name="links-locked"]')).toBeDisabled()
    await dialog.locator('button[name="update"]').click()
    await expect(rows.first()).toContainText('Updated')
    await expect(page.locator(PICTURES)).toHaveAttribute('src', new RegExp(PNG_SQUARE.slice(0, 40)))
    await page.screenshot({ path: test.info().outputPath('links-picture.png') })
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const media = Object.keys(zip.files).filter((name) => name.startsWith('word/media/'))
          const bytes = await Promise.all(media.map((name) => zip.file(name)!.async('base64')))
          return bytes.includes(PNG_SQUARE) && /r:link=/.test(await documentXml(docPath))
        },
        { timeout: 30_000 },
      )
      .toBe(true)

    page.once('dialog', (confirm) => void confirm.accept())
    await dialog.locator('button[name="break"]').click()
    await expect(rows).toHaveCount(0)
    await dialog.locator('button[name="close"]').click()
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).not.toContain('r:link=')
    expect(await documentXml(docPath)).toMatch(/<a:blip\b[^>]*r:embed="/)
    await expect(page.locator(PICTURES)).toHaveCount(1)
  } finally {
    await closeAndSaveVideo(launched, 'docs-links')
    rmSync(dir, { recursive: true, force: true })
  }
})

const fieldRuns = (instr: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  result +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

test('link fields a file came with are listed, and can be locked, set to update automatically and broken', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-links-')))
  const docPath = join(dir, 'fields.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('Host first') +
        // a picture brought in by a field: the paragraph stays a picture block
        `<w:p>${fieldRuns('INCLUDEPICTURE "https://example.com/logo.png" \\* MERGEFORMATINET', pictureRun(1, 20, 10, 60, 40))}</w:p>` +
        `<w:p><w:r><w:t xml:space="preserve">Total: </w:t></w:r>${fieldRuns('LINK Excel.Sheet.12 "C:\\\\data\\\\book.xlsx" "Sheet1!R1C1" \\t', '<w:r><w:t>42</w:t></w:r>')}</w:p>` +
        para('Host last'),
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-links',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Host last' })).toBeVisible({
      timeout: 30_000,
    })
    await expect(page.locator(PICTURES)).toHaveCount(1)
    await expect(page.locator('.doc-page')).toContainText('Total: 42')
    const dialog = await openLinks(app, page)
    const rows = dialog.locator('button.links-row')
    await expect(rows).toHaveCount(2)
    await expect(rows.nth(0)).toContainText('logo.png')
    await expect(rows.nth(0)).toContainText('Picture')
    await expect(rows.nth(1)).toContainText('book.xlsx')
    await expect(rows.nth(1)).toContainText('Sheet1!R1C1')
    await expect(rows.nth(1)).toContainText('Object')
    await page.screenshot({ path: test.info().outputPath('links-fields.png') })

    // the field of the picture block: nothing here reads a web address, it can be locked
    await expect(dialog.locator('button[name="update"]')).toBeDisabled()
    await dialog.locator('input[name="links-locked"]').check()
    await expect(rows.nth(0)).toContainText('Locked')

    // the LINK field: Automatic is its \a switch
    await rows.nth(1).click()
    await expect(dialog.locator('button[name="update"]')).toBeDisabled()
    const [automatic, manual] = [
      dialog.locator('input[name="links-update"]').nth(0),
      dialog.locator('input[name="links-update"]').nth(1),
    ]
    await expect(manual).toBeChecked()
    await automatic.check()
    await expect(rows.nth(1)).toContainText('Automatic')
    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 30_000 })
      .toMatch(/"Sheet1!R1C1" \\t \\a/)
    let xml = await documentXml(docPath)
    expect(xml).toMatch(/w:fldLock="true"\/><\/w:r><w:r><w:instrText[^>]*> INCLUDEPICTURE/)
    expect(fieldChars(xml)).toBe('bsebse')

    page.once('dialog', (confirm) => void confirm.accept())
    await dialog.locator('button[name="break"]').click()
    await expect(rows).toHaveCount(1)
    await expect(rows.first()).toContainText('logo.png')
    page.once('dialog', (confirm) => void confirm.accept())
    await dialog.locator('button[name="break"]').click()
    await expect(rows).toHaveCount(0)
    await dialog.locator('button[name="close"]').click()
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).not.toContain('instrText')
    xml = await documentXml(docPath)
    expect(fieldChars(xml)).toBe('')
    expect(xml).toContain('Total: 42')
    expect(xml).toContain('<w:drawing>')
    await expect(page.locator(PICTURES)).toHaveCount(1)
    await expect(page.locator('.doc-page')).toContainText('Total: 42')
  } finally {
    await closeAndSaveVideo(launched, 'docs-links')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('several links selected at once are locked and broken together; Open Source and Change Source take one', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-links-')))
  const docPath = join(dir, 'several.docx')
  const link = (book: string, result: string) =>
    `<w:p><w:r><w:t xml:space="preserve">${book}: </w:t></w:r>${fieldRuns(`LINK Excel.Sheet.12 "C:\\\\data\\\\${book}.xlsx" "Sheet1!R1C1" \\t`, `<w:r><w:t>${result}</w:t></w:r>`)}</w:p>`
  writeFileSync(
    docPath,
    await docxWith(
      para('Host first') +
        link('alpha', '1') +
        link('beta', '2') +
        link('gamma', '3') +
        para('End'),
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-links',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'End' })).toBeVisible({ timeout: 30_000 })
    const dialog = await openLinks(app, page)
    const rows = dialog.locator('button.links-row')
    const selected = dialog.locator('button.links-row[aria-selected="true"]')
    await expect(rows).toHaveCount(3)
    await expect(selected).toHaveCount(1)
    await expect(dialog.locator('button[name="open"]')).toBeEnabled()

    await rows.nth(2).click({ modifiers: ['Shift'] })
    await expect(selected).toHaveCount(3)
    await expect(dialog.locator('button[name="open"]')).toBeDisabled()
    await expect(dialog.locator('button[name="change"]')).toBeDisabled()
    await expect(dialog.locator('.links-source')).toHaveText('')
    await page.screenshot({ path: test.info().outputPath('links-several.png') })
    await dialog.locator('input[name="links-locked"]').check()
    for (let i = 0; i < 3; i++) await expect(rows.nth(i)).toContainText('Locked')
    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml(docPath)).match(/w:fldLock="true"/g)?.length ?? 0, {
        timeout: 30_000,
      })
      .toBe(3)

    // the middle one is taken out of the selection: the other two go, it stays
    await rows.nth(1).click({ modifiers: ['ControlOrMeta'] })
    await expect(selected).toHaveCount(2)
    await expect(rows.nth(1)).toHaveAttribute('aria-selected', 'false')
    page.once('dialog', (confirm) => void confirm.accept())
    await dialog.locator('button[name="break"]').click()
    await expect(rows).toHaveCount(1)
    await expect(rows.first()).toContainText('beta.xlsx')
    await expect(selected).toHaveCount(1)
    await dialog.locator('button[name="close"]').click()
    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).not.toContain('alpha.xlsx')
    const xml = await documentXml(docPath)
    expect(fieldChars(xml)).toBe('bse')
    expect(xml).toContain('beta.xlsx')
    expect(xml).not.toContain('gamma.xlsx')
    await expect(page.locator('.doc-page')).toContainText('alpha: 1')
    await expect(page.locator('.doc-page')).toContainText('gamma: 3')
  } finally {
    await closeAndSaveVideo(launched, 'docs-links')
    rmSync(dir, { recursive: true, force: true })
  }
})
