/**
 * Templates: New from Template opens a gallery of the templates folder, a
 * template saved from a document shows there with a picture of its first
 * page, and a macro-enabled template makes a document without its macros.
 */
import { test, expect, type ElectronApplication } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const DOCUMENT_MAIN =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
const MACRO_TEMPLATE_MAIN = 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml'
const VBA_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'

/** a macro-enabled template: the document package declared as one, with a VBA project part */
async function dotm(body: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body))
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types
      .replace(DOCUMENT_MAIN, MACRO_TEMPLATE_MAIN)
      .replace(
        '</Types>',
        '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>' +
          '<Override PartName="/word/vbaData.xml" ContentType="application/vnd.ms-word.vbaData+xml"/></Types>',
      ),
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      `<Relationship Id="rIdVba" Type="${VBA_REL}" Target="vbaProject.bin"/></Relationships>`,
    ),
  )
  zip.file('word/vbaProject.bin', Buffer.from('stand-in for a VBA project'))
  zip.file(
    'word/vbaData.xml',
    '<wne:vbaSuppData xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"/>',
  )
  zip.file(
    'word/_rels/vbaProject.bin.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/wordVbaData" Target="vbaData.xml"/></Relationships>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const clickFile = (app: ElectronApplication, label: string) =>
  expect
    .poll(
      () =>
        app.evaluate(({ Menu }, label) => {
          const file = Menu.getApplicationMenu()?.items.find((item) => item.label === 'File')
          const item = file?.submenu?.items.find((candidate) => candidate.label === label)
          if (!item?.enabled) return false
          item.click()
          return true
        }, label),
      { timeout: 15_000 },
    )
    .toBe(true)

const saveTo = (app: ElectronApplication, file: string) =>
  app.evaluate(({ dialog }, file) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: file })
  }, file)

const parts = async (file: string) => Object.keys((await JSZip.loadAsync(readFileSync(file))).files)
const part = async (file: string, name: string) =>
  ((await JSZip.loadAsync(readFileSync(file))).file(name)?.async('string') ?? '') as Promise<string>

/** LibreOffice as a reader that is not this application: whether it converts the file; null where it is not installed */
function officeReads(file: string, outDir: string): boolean | null {
  try {
    execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', outDir, file], {
      stdio: 'ignore',
      timeout: 120_000,
    })
    return existsSync(
      join(
        outDir,
        file
          .split('/')
          .pop()!
          .replace(/\.docx$/, '.pdf'),
      ),
    )
  } catch (err) {
    return (err as NodeJS.ErrnoException).code === 'ENOENT' ? null : false
  }
}

test('a template saved from a document shows in the gallery with its first page, and makes a new document', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-template-gallery-')))
  const docPath = join(dir, 'letter.docx')
  const madePath = join(dir, 'made.docx')
  writeFileSync(docPath, await docxWith(para('Dear reader') + para('Yours truly')))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-template-gallery',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Dear reader' })).toBeVisible({
      timeout: 30_000,
    })
    const templates = join(await app.evaluate(({ app }) => app.getPath('userData')), 'Templates')
    const gallery = page.locator('.template-gallery-dialog')

    // nothing saved yet: the gallery says so, and still leads to any file
    await clickFile(app, 'New from Template…')
    await expect(gallery).toBeVisible()
    await expect(gallery.locator('.template-tile')).toHaveCount(0)
    await expect(gallery.locator('.template-gallery-empty')).toBeVisible()
    await expect(gallery.locator('button.primary')).toBeDisabled()
    await expect(gallery.locator('button[name="browse"]')).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('gallery-empty.png') })
    // a cancelled file picker comes back to the gallery
    await app.evaluate(({ dialog }) => {
      dialog.showOpenDialog = async () => ({ canceled: true, filePaths: [] })
    })
    await gallery.locator('button[name="browse"]').click()
    await page.waitForTimeout(500)
    await expect(gallery).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(gallery).toHaveCount(0)

    // Save as Template writes into the templates folder, with a picture of the first page
    const templatePath = join(templates, 'Letter.dotx')
    await saveTo(app, templatePath)
    await clickFile(app, 'Save as Template…')
    await expect.poll(() => existsSync(templatePath), { timeout: 30_000 }).toBe(true)
    expect(await parts(templatePath)).toContain('docProps/thumbnail.jpeg')
    expect(await part(templatePath, '_rels/.rels')).toContain('metadata/thumbnail')
    const jpeg = await (
      await JSZip.loadAsync(readFileSync(templatePath))
    )
      .file('docProps/thumbnail.jpeg')!
      .async('nodebuffer')
    expect([jpeg[0], jpeg[1]]).toEqual([0xff, 0xd8])
    // a second template, put there by hand, that carries no picture
    writeFileSync(join(templates, 'Plain.dotx'), readFileSync(docPath))
    writeFileSync(join(templates, 'notes.txt'), 'not a template')

    // the document moves on, so that what the template makes can be told from it
    await page.locator('.doc-page p', { hasText: 'Dear reader' }).click()
    await page.keyboard.press('End')
    await page.keyboard.type(' changed')
    // saved, so replacing the document has nothing to ask about
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('changed')

    await clickFile(app, 'New from Template…')
    await expect(gallery.locator('.template-tile .template-name')).toHaveText(['Letter', 'Plain'])
    const letter = gallery.locator('.template-tile[data-template="Letter.dotx"]')
    await expect
      .poll(() =>
        letter.locator('img').evaluate((img: HTMLImageElement) => img.complete && img.naturalWidth),
      )
      .toBe(320)
    await expect(gallery.locator('.template-tile[data-template="Plain.dotx"] img')).toHaveCount(0)
    await expect(letter).toHaveAttribute('aria-selected', 'true')
    await page.screenshot({ path: test.info().outputPath('gallery.png') })
    await gallery.locator('button.primary').click()
    await expect(gallery).toHaveCount(0)
    await expect(page.locator('.doc-page')).toContainText('Dear reader')
    await expect(page.locator('.doc-page')).not.toContainText('changed')

    // unsaved changes in an untitled document: declining to replace it comes back to the gallery
    await page.locator('.doc-page p', { hasText: 'Yours truly' }).click()
    await page.keyboard.press('End')
    await page.keyboard.type(' kept')
    await app.evaluate(({ dialog }) => {
      dialog.showMessageBox = (async () => ({
        response: 2,
        checkboxChecked: false,
      })) as typeof dialog.showMessageBox
    })
    await clickFile(app, 'New from Template…')
    await expect(gallery.locator('.template-tile')).toHaveCount(2)
    await gallery.locator('button.primary').click()
    await page.waitForTimeout(500)
    await expect(gallery).toBeVisible()
    await expect(page.locator('.doc-page')).toContainText('kept')
    await page.keyboard.press('Escape')
    await expect(gallery).toHaveCount(0)

    // what it made is an untitled document: saving asks where, and writes a .docx without the picture
    await saveTo(app, madePath)
    await clickFile(app, 'Save As…')
    await expect.poll(() => existsSync(madePath), { timeout: 30_000 }).toBe(true)
    expect(await part(madePath, '[Content_Types].xml')).toContain(DOCUMENT_MAIN)
    expect(await parts(madePath)).not.toContain('docProps/thumbnail.jpeg')
    expect(await part(madePath, '_rels/.rels')).not.toContain('thumbnail')
    expect(await part(madePath, 'word/document.xml')).toContain('Dear reader')
    expect(await parts(templatePath)).toContain('docProps/thumbnail.jpeg')
  } finally {
    await closeAndSaveVideo(launched, 'docs-template-gallery')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a macro-enabled template makes a document that holds no macros', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-template-dotm-')))
  const docPath = join(dir, 'start.docx')
  const madePath = join(dir, 'from-dotm.docx')
  writeFileSync(docPath, await docxWith(para('Start document')))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-template-gallery',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Start document' })).toBeVisible({
      timeout: 30_000,
    })
    const templates = join(await app.evaluate(({ app }) => app.getPath('userData')), 'Templates')
    mkdirSync(templates, { recursive: true })
    const dotmPath = join(templates, 'Report.dotm')
    writeFileSync(dotmPath, await dotm(para('Macro report') + para('Second line')))
    expect(await parts(dotmPath)).toContain('word/vbaProject.bin')

    await clickFile(app, 'New from Template…')
    const gallery = page.locator('.template-gallery-dialog')
    const tile = gallery.locator('.template-tile[data-template="Report.dotm"]')
    await expect(tile).toBeVisible()
    await tile.dblclick()
    await expect(page.locator('.doc-page')).toContainText('Macro report', { timeout: 30_000 })
    await expect(page.locator('.doc-page')).not.toContainText('Start document')
    await expect(page.locator('body')).toContainText(
      'The new document was made without the macros of the template',
    )

    await page.locator('.doc-page p', { hasText: 'Second line' }).click()
    await page.keyboard.press('End')
    await page.keyboard.type(' edited')
    await saveTo(app, madePath)
    await clickFile(app, 'Save As…')
    await expect.poll(() => existsSync(madePath), { timeout: 30_000 }).toBe(true)
    const names = await parts(madePath)
    expect(names.filter((name) => /vba/i.test(name))).toEqual([])
    const types = await part(madePath, '[Content_Types].xml')
    expect(types).toContain(DOCUMENT_MAIN)
    expect(types).not.toMatch(/macroEnabled|vbaData/)
    expect(await part(madePath, 'word/_rels/document.xml.rels')).not.toContain('vbaProject')
    expect(await part(madePath, 'word/document.xml')).toContain('Second line edited')
    // the template itself is as it was
    expect(await parts(dotmPath)).toContain('word/vbaProject.bin')
    expect(officeReads(madePath, dir)).not.toBe(false)

    // the document reopens as an ordinary one
    await closeAndSaveVideo(launched, 'docs-template-gallery')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-template-gallery',
      openFile: madePath,
    })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(reopened.locator('.doc-page')).toContainText('Second line edited', {
      timeout: 30_000,
    })
  } finally {
    await closeAndSaveVideo(launched, 'docs-template-gallery')
    rmSync(dir, { recursive: true, force: true })
  }
})
