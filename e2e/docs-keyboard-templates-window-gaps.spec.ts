/**
 * Keys a document carries run while it is open and Customize Keyboard saves
 * into it; a protected template asks for its password; a template saved from
 * Draft view still gets its picture; Reduce File Size stores BMP and WebP
 * pictures again as PNG / JPEG; a window dragged out of the tab strip is one
 * of the windows the keyboard goes round; Insert > Icons finds icons by the
 * category names of the UI language.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
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
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'
import { encryptOoxml } from '../packages/electron-utils/src/ooxml-encryption'

const VIDEO_DIR = 'docs-keyboard-templates-window-gaps'
const IS_MAC = process.platform === 'darwin'

const scratch = (name: string) =>
  realpathSync(mkdtempSync(join(tmpdir(), `genoffice-e2e-${name}-`)))

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

const ribbonTab = (page: Page, name: string) =>
  page
    .locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) })
    .first()
    .click()

const menuLabels = (app: ElectronApplication, menu: string) =>
  app.evaluate(
    ({ Menu }, menu) =>
      Menu.getApplicationMenu()
        ?.items.find((item) => item.label === menu)
        ?.submenu?.items.map((item) => (item.type === 'separator' ? '---' : item.label)) ?? [],
    menu,
  )

/** selects the whole text of the paragraph that reads `text` */
const selectLine = (page: Page, text: string) =>
  page.evaluate((text) => {
    interface Node {
      isTextblock: boolean
      textContent: string
    }
    const { editor } = (
      window as never as {
        __aidocs: {
          editor: {
            state: { doc: { descendants(visit: (node: Node, pos: number) => void): void } }
            chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } }
          }
        }
      }
    ).__aidocs
    let from = -1
    editor.state.doc.descendants((node, pos) => {
      if (from < 0 && node.isTextblock && node.textContent === text) from = pos + 1
    })
    if (from < 0) throw new Error(`no paragraph reads "${text}"`)
    editor
      .chain()
      .focus()
      .setTextSelection({ from, to: from + text.length })
      .run()
  }, text)

const part = async (file: string, name: string) =>
  ((await JSZip.loadAsync(readFileSync(file))).file(name)?.async('string') ?? '') as Promise<string>

/** the run properties of the paragraph holding `text`, as saved */
const savedRunProps = async (docPath: string, text: string) => {
  const paragraphs = (await part(docPath, 'word/document.xml')).split(/<w:p[ >]/)
  const paragraph = paragraphs.find((p) => p.includes(text)) ?? ''
  return (paragraph.match(/<w:rPr>.*?<\/w:rPr>/g) ?? []).join('')
}

const KEYS_REL = 'http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations'
const KEYS_TYPE = 'application/vnd.ms-word.keyMapCustomizations+xml'
const WNE = 'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"'
// Control+Shift+Y, Control+Shift+X and Control+Shift+W: a command this app has, a macro, a command it has not
const STRIKE_KEY =
  '<wne:keymap wne:kcmPrimary="0359"><wne:fci wne:fciName="Strikethrough" wne:swArg="0000"/></wne:keymap>'
const MACRO_KEY =
  '<wne:keymap wne:kcmPrimary="0358"><wne:macro wne:macroName="PROJECT.NEWMACROS.STAMP"/></wne:keymap>'
const UNKNOWN_KEY =
  '<wne:keymap wne:kcmPrimary="0357"><wne:fci wne:fciName="MailMergeWizard" wne:swArg="0000"/></wne:keymap>'
const OTHER_KEYBOARD =
  '<wne:keymapsBad><wne:keymap wne:chmPrimary="06DF"><wne:fci wne:fciName="Bold"/></wne:keymap></wne:keymapsBad>'
const CUSTOMIZATIONS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<wne:tcg ${WNE}>` +
  `<wne:keymaps>${STRIKE_KEY}${MACRO_KEY}${UNKNOWN_KEY}</wne:keymaps>${OTHER_KEYBOARD}</wne:tcg>`

async function withCustomizations(docx: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx)
  zip.file('word/customizations.xml', CUSTOMIZATIONS)
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      `<Relationship Id="rIdKeys" Type="${KEYS_REL}" Target="customizations.xml"/></Relationships>`,
    ),
  )
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      `<Override PartName="/word/customizations.xml" ContentType="${KEYS_TYPE}"/></Types>`,
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function openDoc(docPath: string, shows: string, lang?: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    ...(lang ? { lang } : {}),
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p', { hasText: shows }).first()).toBeVisible({
    timeout: 30_000,
  })
  return { launched, page }
}

async function openCustomizeKeyboard(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Tools', 'Customize Keyboard…'])
  const dialog = page.locator('.modal-customize-keyboard')
  await expect(dialog).toBeVisible()
  await expect(dialog.locator('[data-command="font-dialog"]')).toBeAttached()
  return dialog
}

test('keys a document carries run while it is open, and Customize Keyboard saves into the document', async () => {
  test.setTimeout(300_000)
  const dir = scratch('doc-keys')
  const docPath = join(dir, 'keys.docx')
  const plainPath = join(dir, 'plain.docx')
  const lines = para('First line') + para('Second line') + para('Third line')
  writeFileSync(docPath, await withCustomizations(await docxWith(lines)))
  writeFileSync(plainPath, await docxWith(lines))
  let { launched, page } = await openDoc(docPath, 'Third line')
  try {
    const { app } = launched
    const machineFile = join(launched.userDataDir, 'docs-keyboard.json')

    // the key the document assigns runs its command
    await selectLine(page, 'Second line')
    await page.keyboard.press('Control+Shift+Y')
    await expect(page.locator('.doc-page p', { hasText: 'Second line' }).locator('s')).toHaveText(
      'Second line',
    )

    const dialog = await openCustomizeKeyboard(app, page)
    const press = dialog.locator('#ck-press')
    const note = dialog.locator('.ck-note')
    const assign = dialog.getByRole('button', { name: 'Assign' })
    const currentKeys = dialog.locator('.ck-current button')
    const saveIn = dialog.locator('.ck-save-in select')
    const pick = async (id: string) => {
      await dialog.locator(`[data-command="${id}"]`).click()
      await expect(dialog.locator(`[data-command="${id}"]`)).toHaveAttribute(
        'aria-selected',
        'true',
      )
    }
    await dialog.getByRole('option', { name: 'All Commands' }).click()

    // this computer knows nothing of the document's key; the document as the target shows it
    await pick('strikethrough')
    await expect(currentKeys.filter({ hasText: /Y$/ })).toHaveCount(0)
    await expect(saveIn.locator('option')).toHaveText([
      'This computer (all documents)',
      'keys.docx',
    ])
    await saveIn.selectOption('document')
    await expect(currentKeys.filter({ hasText: /Y$/ })).toHaveCount(1)

    // a new key goes into the document
    await pick('italic')
    await press.press('Control+Shift+F8')
    await expect(note).toHaveText('[unassigned]')
    await assign.click()
    await expect(currentKeys.filter({ hasText: 'F8' })).toHaveCount(1)

    // a command Word has no name for cannot be saved there
    await pick('numbering')
    await press.press('Control+Shift+F7')
    await expect(note).toContainText('cannot be saved in a document')
    await expect(assign).toBeDisabled()

    // Remove takes the document's own key map away
    await pick('strikethrough')
    await currentKeys.filter({ hasText: /Y$/ }).click()
    await dialog.getByRole('button', { name: 'Remove' }).click()
    await expect(currentKeys.filter({ hasText: /Y$/ })).toHaveCount(0)
    await page.screenshot({ path: test.info().outputPath('customize-keyboard-document.png') })
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()

    await selectLine(page, 'Third line')
    await page.keyboard.press('Control+Shift+F8')
    await expect(
      page.locator('.doc-page p', { hasText: 'Third line' }).locator('em, i'),
    ).toHaveText('Third line')
    await selectLine(page, 'First line')
    await page.keyboard.press('Control+Shift+Y')
    await expect(page.locator('.doc-page p', { hasText: 'First line' }).locator('s')).toHaveCount(0)

    await saveFromMenu(app)
    await expect
      .poll(() => savedRunProps(docPath, 'Third line'), { timeout: 30_000 })
      .toContain('<w:i/>')
    const saved = await part(docPath, 'word/customizations.xml')
    // what this app does not run is written as it was; the rest of the part is untouched
    expect(saved).toBe(
      CUSTOMIZATIONS.replace(
        STRIKE_KEY + MACRO_KEY + UNKNOWN_KEY,
        MACRO_KEY +
          UNKNOWN_KEY +
          '<wne:keymap wne:kcmPrimary="0377"><wne:fci wne:fciName="Italic" wne:swArg="0000"/></wne:keymap>',
      ),
    )
    expect(existsSync(machineFile)).toBe(false)
    await closeAndSaveVideo(launched, VIDEO_DIR)

    // read back from the file, on a machine that has no assignments of its own
    ;({ launched, page } = await openDoc(docPath, 'Third line'))
    await selectLine(page, 'First line')
    await page.keyboard.press('Control+Shift+F8')
    await expect(
      page.locator('.doc-page p', { hasText: 'First line' }).locator('em, i'),
    ).toHaveText('First line')
    await closeAndSaveVideo(launched, VIDEO_DIR)

    // a document without the part gets one with its first key
    ;({ launched, page } = await openDoc(plainPath, 'Third line'))
    const plain = await openCustomizeKeyboard(launched.app, page)
    await plain.getByRole('option', { name: 'All Commands' }).click()
    await plain.locator('.ck-save-in select').selectOption('document')
    await plain.locator('[data-command="bold"]').click()
    await plain.locator('#ck-press').press('Control+Shift+F9')
    await plain.getByRole('button', { name: 'Assign' }).click()
    await plain.getByRole('button', { name: 'OK' }).click()
    await saveFromMenu(launched.app)
    await expect
      .poll(() => part(plainPath, 'word/customizations.xml'), { timeout: 30_000 })
      .toContain(
        '<wne:keymaps><wne:keymap wne:kcmPrimary="0378"><wne:fci wne:fciName="Bold" wne:swArg="0000"/></wne:keymap></wne:keymaps>',
      )
    expect(await part(plainPath, 'word/_rels/document.xml.rels')).toContain(
      `Type="${KEYS_REL}" Target="customizations.xml"`,
    )
    expect(await part(plainPath, '[Content_Types].xml')).toContain(
      `<Override PartName="/word/customizations.xml" ContentType="${KEYS_TYPE}"/>`,
    )
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => undefined)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a protected template asks for its password, and a template saved from Draft view gets its picture', async () => {
  test.setTimeout(300_000)
  const dir = scratch('templates')
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(para('Dear reader') + para('Yours truly')))
  const { launched, page } = await openDoc(docPath, 'Dear reader')
  try {
    const { app } = launched
    const templates = join(launched.userDataDir, 'Templates')
    mkdirSync(templates, { recursive: true })
    writeFileSync(
      join(templates, 'Secret.dotx'),
      encryptOoxml(await docxWith(para('Behind a password')), 'open sesame'),
    )
    const gallery = page.locator('.template-gallery-dialog')
    const prompt = page.locator('.pwd-dialog')

    await clickMenu(app, ['File', 'New from Template…'])
    await gallery.locator('.template-tile[data-template="Secret.dotx"]').click()
    await gallery.locator('button.primary').click()
    await expect(prompt).toBeVisible({ timeout: 15_000 })
    await expect(prompt).toContainText('Secret.dotx')
    await prompt.locator('input').fill('wrong')
    await prompt.locator('button.btn-primary').click()
    await expect(prompt.locator('.has-error')).toBeVisible({ timeout: 30_000 })
    await expect(page.locator('.doc-page')).toContainText('Dear reader')
    await prompt.locator('input').fill('open sesame')
    await prompt.locator('button.btn-primary').click()
    await expect(prompt).toHaveCount(0, { timeout: 30_000 })
    await expect(page.locator('.doc-page')).toContainText('Behind a password', { timeout: 30_000 })
    await expect(page.locator('.doc-page')).not.toContainText('Dear reader')

    // Draft has no page to take a picture of: the first page is taken as it prints
    await clickMenu(app, ['View', 'Draft'])
    await expect(page.locator('.doc-zoom.view-draft')).toBeVisible()
    const templatePath = join(templates, 'FromDraft.dotx')
    await app.evaluate(({ dialog }, file) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: file })
    }, templatePath)
    await clickMenu(app, ['File', 'Save as Template…'])
    await expect.poll(() => existsSync(templatePath), { timeout: 60_000 }).toBe(true)
    const jpeg = await (
      await JSZip.loadAsync(readFileSync(templatePath))
    )
      .file('docProps/thumbnail.jpeg')
      ?.async('nodebuffer')
    expect(jpeg && [jpeg[0], jpeg[1]]).toEqual([0xff, 0xd8])
    // the view the user was in is the view the user is left in
    await expect(page.locator('.doc-zoom.view-draft')).toBeVisible()
    await expect(page.locator('.pagination-preview')).toHaveCount(0)

    await clickMenu(app, ['File', 'New from Template…'])
    const tile = gallery.locator('.template-tile[data-template="FromDraft.dotx"]')
    await expect
      .poll(() =>
        tile.locator('img').evaluate((img: HTMLImageElement) => img.complete && img.naturalWidth),
      )
      .toBe(320)
    // a Letter page: taller than wide, and not a blank sheet
    const picture = await tile.locator('img').evaluate((img: HTMLImageElement) => {
      const canvas = document.createElement('canvas')
      canvas.width = img.naturalWidth
      canvas.height = img.naturalHeight
      const ctx = canvas.getContext('2d')!
      ctx.drawImage(img, 0, 0)
      const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height)
      let dark = 0
      for (let i = 0; i < data.length; i += 4) if (data[i] + data[i + 1] + data[i + 2] < 384) dark++
      return { height: img.naturalHeight, dark, corner: [data[0], data[1], data[2]] }
    })
    expect(picture.height).toBeGreaterThan(380)
    expect(picture.dark).toBeGreaterThan(20)
    expect(Math.min(...picture.corner)).toBeGreaterThan(240)
    await page.screenshot({ path: test.info().outputPath('gallery-from-draft.png') })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

/** a 24-bit BMP of w x h in coloured squares */
function patternBmp(w: number, h: number, cell: number): Buffer {
  const row = Math.ceil((w * 3) / 4) * 4
  const out = Buffer.alloc(54 + row * h)
  out.write('BM', 0, 'ascii')
  out.writeUInt32LE(out.length, 2)
  out.writeUInt32LE(54, 10)
  out.writeUInt32LE(40, 14)
  out.writeInt32LE(w, 18)
  out.writeInt32LE(h, 22)
  out.writeUInt16LE(1, 26)
  out.writeUInt16LE(24, 28)
  out.writeUInt32LE(row * h, 34)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const shade = ((Math.floor(x / cell) * 37 + Math.floor(y / cell) * 91) % 200) + 30
      out.set([shade, 255 - shade, (shade * 2) % 256], 54 + y * row + x * 3)
    }
  }
  return out
}

// two frames of one pixel: an animation, however small
const ANIMATED_GIF = Buffer.from([
  ...Buffer.from('GIF89a', 'ascii'),
  ...[1, 0, 1, 0, 0x80, 0, 0],
  ...[0, 0, 0, 255, 255, 255],
  ...[0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 2, 0x44, 0x01, 0],
  ...[0x2c, 0, 0, 0, 0, 1, 0, 1, 0, 0, 2, 2, 0x4c, 0x01, 0],
  0x3b,
])

const emu = (px: number) => px * 9525
const pictureRun = (rId: string, id: number, w: number, h: number) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`

/** a document whose pictures are the given media files, each shown at 200 x 150 */
async function pictureDoc(media: Record<string, Buffer>): Promise<Buffer> {
  const names = Object.keys(media)
  const zip = await JSZip.loadAsync(
    await docxWith(
      para('Pictures') +
        names.map((_, i) => `<w:p>${pictureRun(`rIdPic${i}`, i + 1, 200, 150)}</w:p>`).join(''),
    ),
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      names
        .map(
          (name, i) =>
            `<Relationship Id="rIdPic${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`,
        )
        .join('') + '</Relationships>',
    ),
  )
  const types = await zip.file('[Content_Types].xml')!.async('string')
  const extensions = [...new Set(names.map((name) => name.split('.').pop()!))]
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '<Override ',
      extensions.map((e) => `<Default Extension="${e}" ContentType="image/${e}"/>`).join('') +
        '<Override ',
    ),
  )
  for (const [name, data] of Object.entries(media)) zip.file(`word/media/${name}`, data)
  return zip.generateAsync({ type: 'nodebuffer' })
}

const mediaOf = async (file: string) => {
  const zip = await JSZip.loadAsync(readFileSync(file))
  const out: Record<string, Buffer> = {}
  for (const name of Object.keys(zip.files).filter((f) => /^word\/media\/[^/]+$/.test(f)))
    out[name.slice('word/media/'.length)] = await zip.file(name)!.async('nodebuffer')
  return out
}

async function reduce(app: ElectronApplication, page: Page, quality: RegExp) {
  await clickMenu(app, ['File', 'Reduce File Size…'])
  const dialog = page.locator('.reduce-size-dialog')
  await dialog.getByLabel(quality).check()
  await dialog.locator('button.primary').click()
  await expect(dialog.locator('.reduce-size-result[data-status]')).toBeVisible({ timeout: 60_000 })
  return dialog
}

const shownPictures = (page: Page) =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLImageElement>('.doc-page img'))
      .filter((img) => img.complete && img.naturalWidth > 1)
      .map((img) => `${img.naturalWidth}x${img.naturalHeight}`),
  )

test('Reduce File Size stores a BMP again as PNG and a lossy WebP as JPEG, and leaves an animated GIF alone', async () => {
  test.setTimeout(300_000)
  const dir = scratch('reduce-formats')
  const docPath = join(dir, 'formats.docx')
  const webpPath = join(dir, 'webp.docx')
  writeFileSync(
    docPath,
    await pictureDoc({ 'photo.bmp': patternBmp(800, 600, 16), 'spin.gif': ANIMATED_GIF }),
  )
  const { launched, page } = await openDoc(docPath, 'Pictures')
  try {
    const { app } = launched
    await expect.poll(() => shownPictures(page), { timeout: 30_000 }).toEqual(['800x600'])
    const before = await mediaOf(docPath)

    const dialog = await reduce(app, page, /96 ppi/)
    await expect(dialog.locator('.reduce-size-result')).toHaveAttribute('data-status', 'done')
    await expect(dialog.locator('.reduce-size-result')).toContainText('1 of 2 pictures')
    await page.screenshot({ path: test.info().outputPath('reduce-formats.png') })
    await page.keyboard.press('Escape')

    const after = await mediaOf(docPath)
    expect(Object.keys(after).sort()).toEqual(['image1.png', 'photo.png', 'spin.gif'])
    const png = after['photo.png']
    expect([...png.subarray(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47])
    // 200 x 150 px on the page at 96 ppi
    expect([png.readUInt32BE(16), png.readUInt32BE(20)]).toEqual([200, 150])
    expect(png.length).toBeLessThan(before['photo.bmp'].length / 4)
    expect(after['spin.gif'].equals(before['spin.gif'])).toBe(true)
    const rels = await part(docPath, 'word/_rels/document.xml.rels')
    expect(rels).toContain('Target="media/photo.png"')
    expect(rels).not.toContain('photo.bmp')
    expect(await part(docPath, 'word/document.xml')).toContain(
      `<wp:extent cx="${emu(200)}" cy="${emu(150)}"/>`,
    )
    // the page shows the picture the file now holds, at the size it had
    await expect.poll(() => shownPictures(page), { timeout: 30_000 }).toEqual(['200x150'])

    // a lossy WebP, made by the browser itself, comes out a JPEG
    const webp = await page.evaluate(async () => {
      const canvas = new OffscreenCanvas(800, 600)
      const ctx = canvas.getContext('2d', { alpha: false })!
      // noise: a picture that takes bytes, so a quarter of its pixels is sure to take fewer
      const noise = ctx.createImageData(800, 600)
      let state = 7
      for (let i = 0; i < noise.data.length; i++) {
        state = (state * 1103515245 + 12345) & 0x7fffffff
        noise.data[i] = i % 4 === 3 ? 255 : (state >> 8) & 255
      }
      ctx.putImageData(noise, 0, 0)
      const blob = await canvas.convertToBlob({ type: 'image/webp', quality: 0.95 })
      return [blob.type, Array.from(new Uint8Array(await blob.arrayBuffer()))] as const
    })
    expect(webp[0]).toBe('image/webp')
    writeFileSync(webpPath, await pictureDoc({ 'photo.webp': Buffer.from(webp[1]) }))
    await page.evaluate(
      (path) =>
        (
          window as never as { __aidocs: { openPath(path: string): Promise<unknown> } }
        ).__aidocs.openPath(path),
      webpPath,
    )
    await expect.poll(() => shownPictures(page), { timeout: 30_000 }).toEqual(['800x600'])
    const second = await reduce(app, page, /96 ppi/)
    await expect(second.locator('.reduce-size-result')).toHaveAttribute('data-status', 'done')
    const stored = await mediaOf(webpPath)
    expect(Object.keys(stored).sort()).toEqual(['image1.png', 'photo.jpeg'])
    expect([stored['photo.jpeg'][0], stored['photo.jpeg'][1]]).toEqual([0xff, 0xd8])
    expect(await part(webpPath, '[Content_Types].xml')).toContain(
      '<Default Extension="jpeg" ContentType="image/jpeg"/>',
    )
    await expect.poll(() => shownPictures(page), { timeout: 30_000 }).toEqual(['200x150'])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a window dragged out of the tab strip is one of the windows the keyboard goes round, and the Window menu has Minimize All and Zoom All', async () => {
  test.setTimeout(240_000)
  const dir = scratch('windows')
  const docPath = join(dir, 'window.docx')
  writeFileSync(docPath, await docxWith(para('Window one') + para('Last paragraph')))
  const { launched, page } = await openDoc(docPath, 'Window one')
  try {
    const { app } = launched
    const tabs = (from: Page) =>
      from.evaluate(() =>
        (
          window as never as {
            desktop: { listDocsTabs(): Promise<{ id: string; focused: boolean }[]> }
          }
        ).desktop.listDocsTabs(),
      )
    // a second document, in a second tab of the same window
    await expect
      .poll(
        async () => {
          for (const candidate of app.windows()) {
            const opened = await candidate
              .evaluate(async () => {
                const api = (window as never as { aiOffice?: { newDoc(): Promise<void> } }).aiOffice
                if (!api) return false
                await api.newDoc()
                return true
              })
              .catch(() => false)
            if (opened) return true
          }
          return false
        },
        { timeout: 30_000 },
      )
      .toBe(true)
    const docsPages = () => app.windows().filter((p) => p.url().includes('://docs/'))
    await expect.poll(() => docsPages().length, { timeout: 30_000 }).toBe(2)
    const second = docsPages().find((p) => p !== page)!
    await expect(second.locator('.status-page')).toContainText(/\d/, { timeout: 30_000 })
    await expect.poll(async () => (await tabs(page)).length, { timeout: 30_000 }).toBe(2)
    expect((await tabs(page)).some((tab) => tab.id.startsWith('detached:'))).toBe(false)

    // Arrange All gives the other document a window of its own: it stays in the round
    await ribbonTab(second, 'View')
    const arrange = second.locator('.ribbon :is(.rb-big, .rb-small)', { hasText: /^Arrange All$/ })
    await arrange.scrollIntoViewIfNeeded()
    await arrange.click()
    await expect
      .poll(async () => (await tabs(page)).filter((tab) => tab.id.startsWith('detached:')).length, {
        timeout: 30_000,
      })
      .toBe(1)
    expect(await tabs(page)).toHaveLength(2)

    if (IS_MAC) {
      // the strip's document first, then the window that left it
      const at = async () => (await tabs(page)).findIndex((tab) => tab.focused)
      await second.locator('.doc-page').first().click()
      await expect.poll(at).toBe(0)
      await second.keyboard.press('Meta+F6')
      await expect.poll(at).toBe(1)
      await page.locator('.doc-page').first().click()
      await page.keyboard.press('Shift+Meta+F6')
      await expect.poll(at).toBe(0)

      const labels = await menuLabels(app, 'Window')
      const front = labels.indexOf('Bring All to Front')
      expect(labels.slice(front, front + 3)).toEqual([
        'Bring All to Front',
        'Minimize All',
        'Zoom All',
      ])
      const windows = () =>
        app.evaluate(({ BrowserWindow }) =>
          BrowserWindow.getAllWindows()
            .filter((w) => w.isVisible() || w.isMinimized())
            .map((w) => ({ minimized: w.isMinimized(), zoomed: w.isMaximized() })),
        )
      await clickMenu(app, ['Window', 'Zoom All'])
      await expect.poll(async () => (await windows()).every((w) => w.zoomed)).toBe(true)
      await clickMenu(app, ['Window', 'Zoom All'])
      await expect.poll(async () => (await windows()).some((w) => w.zoomed)).toBe(false)
      await clickMenu(app, ['Window', 'Minimize All'])
      await expect
        .poll(async () => (await windows()).every((w) => w.minimized), { timeout: 15_000 })
        .toBe(true)
      expect(await windows()).toHaveLength(2)
      await app.evaluate(({ BrowserWindow }) =>
        BrowserWindow.getAllWindows().forEach((w) => w.isMinimized() && w.restore()),
      )
    }
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Insert > Icons finds icons by the category names of the UI language', async () => {
  test.setTimeout(180_000)
  const dir = scratch('icons-lang')
  const docPath = join(dir, 'icons.docx')
  writeFileSync(docPath, await docxWith(para('apple')))
  const { launched, page } = await openDoc(docPath, 'apple', 'de')
  try {
    const { app } = launched
    await page.locator('.doc-page p', { hasText: 'apple' }).first().click()
    await clickMenu(app, ['Einfügen', 'Symbole…'])
    const dialog = page.locator('.icons-dialog')
    await expect(dialog).toBeVisible()
    const cells = dialog.locator('.icons-dialog-cell')
    const search = dialog.locator('.icons-dialog-search')

    await dialog.locator('[role="tab"][data-category="medical"]').click()
    const inCategory = await cells.count()
    expect(inCategory).toBeGreaterThan(5)
    // the German name of the category is no English keyword of any icon
    await search.fill('Medizin')
    await expect(cells).toHaveCount(inCategory)
    // a category name and an English keyword together
    await search.fill('medizin heart')
    await expect(dialog.locator('.icons-dialog-cell[data-icon="heart_pulse"]')).toBeVisible()
    expect(await cells.count()).toBeLessThan(inCategory)
    // English keywords alone still find their icons
    await search.fill('cardiology')
    await expect(cells).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('icons-search-de.png') })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
