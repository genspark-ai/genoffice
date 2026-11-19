/**
 * Two places where the saved file has to say what the page shows: Convert
 * Permanently of Font Substitution follows Undo and Redo into the parts the
 * editor does not hold, before a save and after one; and a preset of the
 * Layout > Position gallery puts the picture where the saved file puts it.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, rects, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-font-convert-undo-position-presets'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

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

const FONT_PARTS = ['word/document.xml', 'word/styles.xml', 'word/header1.xml']

/** the font slots of each part that names fonts, in the order the part holds them */
async function fontSlots(docPath: string): Promise<Record<string, string[]>> {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const out: Record<string, string[]> = {}
  for (const part of FONT_PARTS) {
    const xml = (await zip.file(part)?.async('string')) ?? ''
    out[part] = Array.from(xml.matchAll(/\sw:(?:ascii|hAnsi|eastAsia|cs)="([^"]*)"/g), (m) => m[1])
  }
  return out
}

const documentXml = async (docPath: string): Promise<string> =>
  (await (
    await JSZip.loadAsync(readFileSync(docPath))
  )
    .file('word/document.xml')
    ?.async('string')) ?? ''

test('Convert Permanently follows Undo and Redo into the saved styles and header, before a save and after one', async () => {
  test.setTimeout(240_000)
  const MISSING = 'Zzyzx Gone Sans'
  const HERE = 'Courier New'
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-font-undo-')))
  const docPath = join(dir, 'fonts.docx')
  const fonts = (name: string) => `<w:rPr><w:rFonts w:ascii="${name}" w:hAnsi="${name}"/></w:rPr>`
  const run = (name: string, text: string) => `<w:r>${fonts(name)}<w:t>${text}</w:t></w:r>`
  const source = await JSZip.loadAsync(
    await docxWith(
      `<w:p>${run(MISSING, 'in a font nobody has')}</w:p>` +
        `<w:p>${run(HERE, 'in a font that is here')}</w:p>` +
        '<w:p><w:pPr><w:pStyle w:val="Gone"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p>' +
        para('Plain') +
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>',
    ),
  )
  source.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      `<w:style w:type="paragraph" w:styleId="Gone"><w:name w:val="Gone"/><w:basedOn w:val="Normal"/>${fonts(MISSING)}</w:style>` +
      `<w:style w:type="paragraph" w:styleId="Here"><w:name w:val="Here"/><w:basedOn w:val="Normal"/>${fonts(HERE)}</w:style>` +
      '</w:styles>',
  )
  source.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${W} ${R}><w:p>${run(MISSING, 'head')}${run(HERE, 'head too')}</w:p></w:hdr>`,
  )
  source.file(
    '[Content_Types].xml',
    (await source.file('[Content_Types].xml')!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
    ),
  )
  source.file(
    'word/_rels/document.xml.rels',
    (await source.file('word/_rels/document.xml.rels')!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>',
    ),
  )
  writeFileSync(docPath, await source.generateAsync({ type: 'nodebuffer' }))
  const original = await fontSlots(docPath)
  expect(original).toEqual({
    'word/document.xml': [MISSING, MISSING, HERE, HERE],
    'word/styles.xml': [MISSING, MISSING, HERE, HERE],
    'word/header1.xml': [MISSING, MISSING, HERE, HERE],
  })

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const plain = page.locator('.doc-page[contenteditable="true"] p', { hasText: 'Plain' })
    await expect(plain).toBeVisible({ timeout: 30_000 })
    const dialog = page.locator('.font-substitution-dialog')
    const missingRows = dialog.locator(`tr[data-font="${MISSING}"]`)
    const convert = async () => {
      await clickMenu(app, ['Format', 'Font Substitution…'])
      await expect(missingRows).toHaveCount(1)
      await dialog.locator('button[name="convert"]').click()
      await expect(dialog).toBeHidden()
    }
    /** how many rows the dialog has for the missing font */
    const listed = async () => {
      await clickMenu(app, ['Format', 'Font Substitution…'])
      await expect(dialog).toBeVisible()
      const count = await missingRows.count()
      await dialog.locator('button.primary').click()
      await expect(dialog).toBeHidden()
      return count
    }

    // converted and undone before any save: the font is a missing font again
    await convert()
    expect(await listed()).toBe(0)
    await clickMenu(app, ['Edit', 'Undo'])
    // an edit of its own, so the save has something to write
    await plain.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' text')
    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('Plain text'), { timeout: 30_000 })
      .toBe(true)
    // the file names the missing font wherever it did, in the styles and the header too
    expect(await fontSlots(docPath)).toEqual(original)
    await expect.poll(listed).toBe(1)

    // converted, undone, redone: the save converts every part
    await convert()
    await clickMenu(app, ['Edit', 'Undo'])
    await clickMenu(app, ['Edit', 'Redo'])
    expect(await listed()).toBe(0)
    await saveFromMenu(app)
    const withoutMissing = async () => {
      const slots = await fontSlots(docPath)
      return FONT_PARTS.map((part) => slots[part].includes(MISSING))
    }
    await expect.poll(withoutMissing, { timeout: 30_000 }).toEqual([false, false, false])
    const converted = await fontSlots(docPath)
    for (const part of FONT_PARTS) {
      // only the slots that named the font changed: none was added, the other font is as it was
      expect(converted[part], part).toHaveLength(4)
      expect(converted[part].slice(2), part).toEqual([HERE, HERE])
    }
    await expect.poll(listed).toBe(0)

    // undone after that save: the next save writes the font back into every part
    await clickMenu(app, ['Edit', 'Undo'])
    await expect.poll(listed).toBe(1)
    await saveFromMenu(app)
    await expect.poll(() => fontSlots(docPath), { timeout: 30_000 }).toEqual(original)

    // and redone after that one: converted again
    await clickMenu(app, ['Edit', 'Redo'])
    await saveFromMenu(app)
    await expect.poll(() => fontSlots(docPath), { timeout: 30_000 }).toEqual(converted)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const emu = (px: number) => px * 9525

/** a run that is one in-line picture, size in px */
const inlinePicture = (id: number, w: number, h: number) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`

// bottom row first: a picture taken out of the line of text shows the difference most there
const PRESETS = (['bottom', 'center', 'top'] as const).flatMap((v) =>
  (['left', 'center', 'right'] as const).map((h) => ({ h, v })),
)
const PICTURE = '.doc-page[contenteditable="true"] .doc-protected img'

/** the corner of the picture, measured from the paper's */
async function onPaper(page: Page): Promise<number[]> {
  const [paper] = await rects(page, '.doc-page[contenteditable="true"]')
  const [picture] = await rects(page, PICTURE)
  return [Math.round(picture.left - paper.left), Math.round(picture.top - paper.top)]
}

/** where the picture is once the page has stopped laying it out: the same in four readings half a second apart */
async function settled(page: Page): Promise<number[]> {
  let last = ''
  let same = 0
  await expect
    .poll(
      async () => {
        const now = JSON.stringify(await onPaper(page))
        same = now === last ? same + 1 : 0
        last = now
        return same
      },
      { intervals: [500], timeout: 30_000 },
    )
    .toBeGreaterThanOrEqual(3)
  return JSON.parse(last) as number[]
}

test('each Position gallery preset shows the picture where the saved file puts it', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-position-presets-')))
  const docPath = join(dir, 'picture.docx')
  const [W_PX, H_PX] = [120, 90]
  writeFileSync(
    docPath,
    await docxWith(
      [1, 2, 3, 4, 5].map((n) => para(`Line ${n} of the text before the picture`)).join('') +
        `<w:p>${inlinePicture(1, W_PX, H_PX)}</w:p>` +
        para('After the picture') +
        para('Last'),
    ),
  )
  const moved: object[] = []
  let last: number[] = []
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const picture = page.locator(PICTURE)
    await expect(picture).toHaveCount(1, { timeout: 30_000 })
    const inLine = await settled(page)
    const [line] = await rects(page, '.doc-page[contenteditable="true"] p')
    const [paper] = await rects(page, '.doc-page[contenteditable="true"]')
    const textLeft = Math.round(line.left - paper.left)
    const position = page.locator('.rb-big', { hasText: /^Position$/ })
    const openGallery = async () => {
      await picture.click()
      // selecting a picture can show its own tab; the gallery is on Layout
      if (!(await position.isVisible()))
        await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
      await expect(position).toBeEnabled()
      await position.click()
    }
    for (const { h, v } of PRESETS) {
      // from a picture in the line of text each time, as the first preset finds it
      await openGallery()
      await page.locator('.layout-menu > button').first().click()
      await openGallery()
      await page.locator(`.pos-cell.ph-${h}.pv-${v}`).click()
      await page.locator('.doc-page[contenteditable="true"] p', { hasText: 'Line 1' }).click()
      const shown = await settled(page)

      await saveFromMenu(app)
      await expect
        .poll(async () => await documentXml(docPath), { timeout: 30_000 })
        .toContain(
          `<wp:positionH relativeFrom="margin"><wp:align>${h}</wp:align></wp:positionH>` +
            `<wp:positionV relativeFrom="margin"><wp:align>${v}</wp:align></wp:positionV>`,
        )
      // the page is drawn again from the saved file
      const saved = await settled(page)
      if (Math.abs(shown[0] - saved[0]) > 2 || Math.abs(shown[1] - saved[1]) > 2)
        moved.push({ preset: `${h},${v}`, shown, saved })
      last = saved

      // and that is the place the preset names: across the text area, and down the page for the lower rows
      const across = { left: 0, center: (line.width - W_PX) / 2, right: line.width - W_PX }[h]
      expect(Math.abs(saved[0] - (textLeft + across)), `${h},${v}`).toBeLessThanOrEqual(2)
      if (v === 'center') expect(saved[1] - inLine[1], `${h},${v}`).toBeGreaterThan(100)
      if (v === 'bottom') expect(saved[1] - inLine[1], `${h},${v}`).toBeGreaterThan(400)
    }
    await page.screenshot({ path: test.info().outputPath('last-preset.png') })
    // every preset showed the picture where the save then put it
    expect(moved).toEqual([])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
  }

  launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator(PICTURE)).toHaveCount(1, { timeout: 30_000 })
    // opened again, the picture is where the last preset left it
    const reopened = await settled(page)
    expect(Math.abs(reopened[0] - last[0])).toBeLessThanOrEqual(2)
    expect(Math.abs(reopened[1] - last[1])).toBeLessThanOrEqual(2)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
