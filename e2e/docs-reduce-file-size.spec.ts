/**
 * File > Reduce File Size: pictures stored at more pixels than the size they
 * are shown at needs are resampled and the document is saved, in the body, in
 * a line of text, in the header, and when the pictures are served lazily from
 * the file. Cropped pictures can be cut down to what they show, and the work
 * can be kept to the selected pictures.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { crc32, deflateSync } from 'node:zlib'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

/** a PNG of w x h in coloured squares of `cell` px: 1 is noise that does not compress */
function patternPng(w: number, h: number, cell: number, seed: number): Buffer {
  let state = seed
  const next = () => (state = (state * 1103515245 + 12345) & 0x7fffffff) >> 8
  const stride = w * 4 + 1
  const raw = Buffer.alloc(stride * h)
  for (let y = 0; y < h; y += cell) {
    for (let x = 0; x < w; x += cell) {
      const [r, g, b] = [next() & 255, next() & 255, next() & 255]
      for (let dx = 0; dx < cell && x + dx < w; dx++)
        raw.set([r, g, b, 255], y * stride + 1 + (x + dx) * 4)
    }
    for (let dy = 1; dy < cell && y + dy < h; dy++)
      raw.copy(raw, (y + dy) * stride, y * stride, (y + 1) * stride)
  }
  const chunk = (type: string, data: Buffer) => {
    const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
    const out = Buffer.alloc(body.length + 8)
    out.writeUInt32BE(data.length, 0)
    body.copy(out, 4)
    out.writeUInt32BE(crc32(body), body.length + 4)
    return out
  }
  const head = Buffer.alloc(13)
  head.writeUInt32BE(w, 0)
  head.writeUInt32BE(h, 4)
  head.set([8, 6, 0, 0, 0], 8)
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', head),
    chunk('IDAT', deflateSync(raw, { level: 1 })),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

const emu = (px: number) => px * 9525

/** a run that is one in-line picture, size in px */
const pictureRun = (rId: string, id: number, w: number, h: number) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`

const text = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const HEADER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header'
const rel = (id: string, type: string, target: string) =>
  `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`
const relsXml = (inner: string) =>
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`

/**
 * image1 is shown in the body (200x150) and larger in the header (400x300),
 * image2 sits in a line of text, image3 only in the header, image4 is stored
 * smaller than it is shown.
 */
async function fixture(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(
    await docxWith(
      para('One') +
        `<w:p>${pictureRun('rIdImg', 1, 200, 150)}</w:p>` +
        `<w:p>${text('before ')}${pictureRun('rIdImg2', 2, 40, 40)}${text(' after')}</w:p>` +
        `<w:p>${pictureRun('rIdImg4', 4, 200, 150)}</w:p>` +
        para('Last') +
        `<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="12240" w:h="15840"/>` +
        `<w:pgMar w:top="2880" w:right="1440" w:bottom="1440" w:left="1440" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>`,
    ),
  )
  const documentXml = await zip.file('word/document.xml')!.async('string')
  const ns = /<w:document ([^>]+)>/.exec(documentXml)![1]
  zip.file(
    'word/header1.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr ${ns}>` +
      `<w:p>${pictureRun('rId1', 11, 400, 300)}${pictureRun('rId2', 12, 240, 60)}</w:p></w:hdr>`,
  )
  zip.file(
    'word/_rels/header1.xml.rels',
    relsXml(
      rel('rId1', IMAGE_REL, 'media/image1.png') + rel('rId2', IMAGE_REL, 'media/image3.png'),
    ),
  )
  zip.file(
    'word/_rels/document.xml.rels',
    relsXml(
      rel('rIdImg', IMAGE_REL, 'media/image1.png') +
        rel('rIdImg2', IMAGE_REL, 'media/image2.png') +
        rel('rIdImg4', IMAGE_REL, 'media/image4.png') +
        rel('rIdHdr', HEADER_REL, 'header1.xml'),
    ),
  )
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
    ),
  )
  zip.file('word/media/image1.png', patternPng(1600, 1200, 8, 1))
  zip.file('word/media/image2.png', patternPng(800, 800, 8, 2))
  zip.file('word/media/image3.png', patternPng(1200, 300, 4, 3))
  zip.file('word/media/image4.png', patternPng(40, 30, 2, 4))
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

const ready = (page: Page) =>
  page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )

/** every picture on the first page, body and header: the size it is shown at and the pixels behind it */
const pictures = (page: Page) =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLImageElement>('.doc-page img, .page-hf-header img'))
      .filter((img) => img.complete && img.naturalWidth > 0)
      .map((img) => {
        const box = img.getBoundingClientRect()
        return {
          shown: `${Math.round(box.width)}x${Math.round(box.height)}`,
          stored: `${img.naturalWidth}x${img.naturalHeight}`,
        }
      })
      .sort((a, b) => a.shown.localeCompare(b.shown) || a.stored.localeCompare(b.stored)),
  )

interface Stored {
  /** part name → [pixel width, pixel height, bytes] */
  media: Record<string, number[]>
  extents: string[]
}

async function stored(docPath: string): Promise<Stored> {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const media: Stored['media'] = {}
  for (const name of Object.keys(zip.files).filter((f) => /^word\/media\/[^/]+$/.test(f))) {
    const bytes = await zip.file(name)!.async('nodebuffer')
    media[name] = [bytes.readUInt32BE(16), bytes.readUInt32BE(20), bytes.length]
  }
  const extents: string[] = []
  for (const part of ['word/document.xml', 'word/header1.xml']) {
    const xml = (await zip.file(part)?.async('string')) ?? ''
    extents.push(...(xml.match(/<(?:wp:extent|a:ext) [^>]*>/g) ?? []))
  }
  return { media, extents }
}

async function reduce(app: ElectronApplication, page: Page, quality: RegExp) {
  await clickMenu(app, ['File', 'Reduce File Size…'])
  const dialog = page.locator('.reduce-size-dialog')
  await dialog.getByLabel(quality).check()
  await dialog.locator('button.primary').click()
  await expect(dialog.locator('.reduce-size-result[data-status]')).toBeVisible({ timeout: 60_000 })
  return dialog
}

test('pictures in the body, in a line of text and in the header are resampled to the size they are shown at', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-reduce-size-')))
  const docPath = join(dir, 'pictures.docx')
  writeFileSync(docPath, await fixture())
  const before = await stored(docPath)
  const fileBytesBefore = statSync(docPath).size
  const shown = ['200x150', '200x150', '240x60', '400x300', '40x40']
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-reduce-file-size',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    await expect
      .poll(async () => (await pictures(page)).map((p) => p.stored), { timeout: 30_000 })
      .toEqual(['1600x1200', '40x30', '1200x300', '1600x1200', '800x800'])
    expect((await pictures(page)).map((p) => p.shown)).toEqual(shown)

    const dialog = await reduce(app, page, /96 ppi/)
    await expect(dialog.locator('.reduce-size-result')).toHaveAttribute('data-status', 'done')
    await expect(dialog.locator('.reduce-size-result')).toContainText(
      '3 of 4 pictures were made smaller',
    )
    await page.screenshot({ path: test.info().outputPath('reduced.png') })

    // the file: the shared picture keeps what its largest use needs, the small one is untouched
    const after = await stored(docPath)
    expect(after.extents).toEqual(before.extents)
    expect(after.media['word/media/image1.png'].slice(0, 2)).toEqual([400, 300])
    expect(after.media['word/media/image2.png'].slice(0, 2)).toEqual([40, 40])
    expect(after.media['word/media/image3.png'].slice(0, 2)).toEqual([240, 60])
    expect(after.media['word/media/image4.png']).toEqual(before.media['word/media/image4.png'])
    for (const name of ['image1', 'image2', 'image3'].map((n) => `word/media/${n}.png`))
      expect(after.media[name][2]).toBeLessThan(before.media[name][2])
    expect(statSync(docPath).size).toBeLessThan(fileBytesBefore)
    test.info().annotations.push({
      type: 'file size',
      description: `${fileBytesBefore} -> ${statSync(docPath).size} bytes`,
    })

    // the page shows the pictures of the file, at the size they were shown at
    await expect
      .poll(() => pictures(page), { timeout: 30_000 })
      .toEqual([
        { shown: '200x150', stored: '400x300' },
        { shown: '200x150', stored: '40x30' },
        { shown: '240x60', stored: '240x60' },
        { shown: '400x300', stored: '400x300' },
        { shown: '40x40', stored: '40x40' },
      ])

    // a second run has nothing left to do and says so
    await dialog.locator('button.primary').click()
    await expect(dialog.locator('.reduce-size-result')).toHaveAttribute('data-status', 'none', {
      timeout: 60_000,
    })
    await dialog.getByRole('button', { name: 'Close' }).click()
    await expect(dialog).toBeHidden()

    // editing on and saving again keeps the smaller pictures
    await page.locator('.doc-page[contenteditable="true"] p', { hasText: 'Last' }).click()
    await page.keyboard.press('End')
    await page.keyboard.type(' word')
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          return /Last word/.test(await zip.file('word/document.xml')!.async('string'))
        },
        { timeout: 30_000 },
      )
      .toBe(true)
    expect((await stored(docPath)).media).toEqual(after.media)
  } finally {
    await closeAndSaveVideo(launched, 'docs-reduce-file-size')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-reduce-file-size',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(page)
    await expect
      .poll(async () => (await pictures(page)).map((p) => p.shown), { timeout: 30_000 })
      .toEqual(shown)
  } finally {
    await closeAndSaveVideo(launched, 'docs-reduce-file-size')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a large picture served from the file on disk is resampled too', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-reduce-size-')))
  const docPath = join(dir, 'large.docx')
  const zip = await JSZip.loadAsync(
    await docxWith(
      para('One') +
        `<w:p>${pictureRun('rIdImg', 1, 300, 200)}</w:p>` +
        `<w:p>${pictureRun('rIdImg2', 2, 300, 200)}</w:p>` +
        para('Last'),
    ),
  )
  zip.file(
    'word/_rels/document.xml.rels',
    relsXml(
      rel('rIdImg', IMAGE_REL, 'media/image1.png') + rel('rIdImg2', IMAGE_REL, 'media/image2.png'),
    ),
  )
  // noise: over the 1 MB of media at which a document's pictures stay in the file
  zip.file('word/media/image1.png', patternPng(1200, 800, 1, 5))
  zip.file('word/media/image2.png', patternPng(300, 200, 1, 6))
  writeFileSync(docPath, await zip.generateAsync({ type: 'nodebuffer' }))
  const before = await stored(docPath)
  expect(before.media['word/media/image1.png'][2]).toBeGreaterThan(1024 * 1024)

  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-reduce-file-size',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const sources = () =>
      page.evaluate(() =>
        Array.from(document.querySelectorAll<HTMLImageElement>('.doc-page img')).map(
          (img) => img.src.split(':')[0],
        ),
      )
    await expect
      .poll(sources, { timeout: 30_000 })
      .toEqual(['genoffice-docx-media', 'genoffice-docx-media'])
    await expect
      .poll(() => pictures(page), { timeout: 30_000 })
      .toEqual([
        { shown: '300x200', stored: '1200x800' },
        { shown: '300x200', stored: '300x200' },
      ])

    const dialog = await reduce(app, page, /150 ppi/)
    await expect(dialog.locator('.reduce-size-result')).toHaveAttribute('data-status', 'done')
    await expect(dialog.locator('.reduce-size-result')).toContainText(
      '1 of 2 pictures were made smaller',
    )
    await dialog.getByRole('button', { name: 'Close' }).click()

    // 300x200 px at 96 dpi needs 469 x 313 px at 150 ppi; 1200x800 scaled to cover both is 470 x 313
    const after = await stored(docPath)
    expect(after.extents).toEqual(before.extents)
    expect(after.media['word/media/image1.png'].slice(0, 2)).toEqual([470, 313])
    expect(after.media['word/media/image1.png'][2]).toBeLessThan(
      before.media['word/media/image1.png'][2] / 2,
    )
    // the picture that stayed is written back whole from the file
    expect(after.media['word/media/image2.png']).toEqual(before.media['word/media/image2.png'])
    await expect
      .poll(() => pictures(page), { timeout: 30_000 })
      .toEqual([
        { shown: '300x200', stored: '300x200' },
        { shown: '300x200', stored: '470x313' },
      ])
  } finally {
    await closeAndSaveVideo(launched, 'docs-reduce-file-size')
    rmSync(dir, { recursive: true, force: true })
  }
})

/** a paragraph that is one picture showing the right half of its bitmap's lower part */
const croppedPicture = (rId: string, id: number, w: number, h: number) =>
  `<w:p>${pictureRun(rId, id, w, h).replace(
    '<a:stretch>',
    '<a:srcRect l="50000" t="25000"/><a:stretch>',
  )}</w:p>`

/** the pixels of what the page shows for a picture, as a small grid of colours */
const shownColours = (page: Page, index: number) =>
  page.evaluate(async (index) => {
    const img = document.querySelectorAll<HTMLImageElement>('.doc-page img')[index]
    const frame = (img.closest('.doc-protected') ?? img).getBoundingClientRect()
    const box = img.getBoundingClientRect()
    const canvas = document.createElement('canvas')
    canvas.width = 20
    canvas.height = 15
    const ctx = canvas.getContext('2d')!
    // the part of the bitmap that falls inside the frame the page gives the picture
    const sx = ((frame.left - box.left) / box.width) * img.naturalWidth
    const sy = ((frame.top - box.top) / box.height) * img.naturalHeight
    const sw = (frame.width / box.width) * img.naturalWidth
    const sh = (frame.height / box.height) * img.naturalHeight
    ctx.drawImage(await createImageBitmap(img), sx, sy, sw, sh, 0, 0, 20, 15)
    return Array.from(ctx.getImageData(0, 0, 20, 15).data)
  }, index)

const meanDifference = (a: number[], b: number[]) =>
  a.reduce((sum, v, i) => sum + Math.abs(v - b[i]), 0) / a.length

test('the cropped areas of the selected picture are deleted, then of every picture, and the page looks the same', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-reduce-crop-')))
  const docPath = join(dir, 'cropped.docx')
  const zip = await JSZip.loadAsync(
    await docxWith(
      para('One') +
        croppedPicture('rIdImg', 1, 200, 225) +
        croppedPicture('rIdImg2', 2, 200, 225) +
        para('Last'),
    ),
  )
  zip.file(
    'word/_rels/document.xml.rels',
    relsXml(
      rel('rIdImg', IMAGE_REL, 'media/image1.png') + rel('rIdImg2', IMAGE_REL, 'media/image2.png'),
    ),
  )
  // large squares, so that a picture cut at the wrong place is a different picture
  zip.file('word/media/image1.png', patternPng(800, 600, 100, 5))
  zip.file('word/media/image2.png', patternPng(800, 600, 100, 6))
  writeFileSync(docPath, await zip.generateAsync({ type: 'nodebuffer' }))
  const before = await stored(docPath)
  const crops = async () =>
    (
      await (
        await JSZip.loadAsync(readFileSync(docPath))
      )
        .file('word/document.xml')!
        .async('string')
    ).match(/<a:blip r:embed="[^"]*"\/>(<a:srcRect[^>]*>)?/g)
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-reduce-file-size',
    openFile: docPath,
  })
  try {
    const { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    await expect
      .poll(async () => (await pictures(page)).map((p) => p.stored), { timeout: 30_000 })
      .toEqual(['800x600', '800x600'])
    const looked = [await shownColours(page, 0), await shownColours(page, 1)]
    await page.screenshot({ path: test.info().outputPath('crop-before.png') })

    // with no picture selected the dialog offers the whole file only
    await clickMenu(app, ['File', 'Reduce File Size…'])
    const dialog = page.locator('.reduce-size-dialog')
    await expect(dialog.getByLabel('Selected pictures only')).toBeDisabled()
    await expect(dialog.getByLabel('All pictures in this file')).toBeChecked()
    // keeping the resolution and the cropped areas leaves nothing to do
    await dialog.getByLabel('Keep current resolution').check()
    await expect(dialog.locator('button.primary')).toBeDisabled()
    await dialog.getByRole('button', { name: 'Close' }).click()

    // the first picture selected: only its cropped area goes, at the resolution it has
    await page.locator('.doc-page img').first().click()
    await clickMenu(app, ['File', 'Reduce File Size…'])
    await dialog.getByLabel('Keep current resolution').check()
    await dialog.getByLabel('Delete cropped areas of pictures').check()
    await dialog.getByLabel('Selected pictures only').check()
    await page.screenshot({ path: test.info().outputPath('crop-dialog.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog.locator('.reduce-size-result')).toHaveAttribute('data-status', 'done', {
      timeout: 60_000,
    })
    await dialog.getByRole('button', { name: 'Close' }).click()

    let after = await stored(docPath)
    expect(after.extents).toEqual(before.extents)
    // the right half of the lower three quarters of 800 x 600
    expect(after.media['word/media/image1.png'].slice(0, 2)).toEqual([400, 450])
    expect(after.media['word/media/image2.png']).toEqual(before.media['word/media/image2.png'])
    expect(await crops()).toEqual([
      '<a:blip r:embed="rIdImg"/>',
      '<a:blip r:embed="rIdImg2"/><a:srcRect l="50000" t="25000"/>',
    ])
    await expect
      .poll(async () => (await pictures(page)).map((p) => p.stored).sort(), { timeout: 30_000 })
      .toEqual(['400x450', '800x600'])
    expect(meanDifference(await shownColours(page, 0), looked[0])).toBeLessThan(4)
    expect(meanDifference(await shownColours(page, 1), looked[1])).toBeLessThan(4)

    // every picture, for the screen: the second loses its cropped area too, and both are resampled
    await clickMenu(app, ['File', 'Reduce File Size…'])
    await dialog.getByLabel(/96 ppi/).check()
    await dialog.getByLabel('Delete cropped areas of pictures').check()
    await dialog.locator('button.primary').click()
    await expect(dialog.locator('.reduce-size-result')).toHaveAttribute('data-status', 'done', {
      timeout: 60_000,
    })
    await expect(dialog.locator('.reduce-size-result')).toContainText(
      '2 of 2 pictures were made smaller',
    )
    await dialog.getByRole('button', { name: 'Close' }).click()
    after = await stored(docPath)
    expect(after.extents).toEqual(before.extents)
    expect(after.media['word/media/image1.png'].slice(0, 2)).toEqual([200, 225])
    expect(after.media['word/media/image2.png'].slice(0, 2)).toEqual([200, 225])
    expect(await crops()).toEqual(['<a:blip r:embed="rIdImg"/>', '<a:blip r:embed="rIdImg2"/>'])
    await page.screenshot({ path: test.info().outputPath('crop-after.png') })

    // reopened: the pictures of the file, shown where and as they were
    await closeAndSaveVideo(launched, 'docs-reduce-file-size')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-reduce-file-size',
      openFile: docPath,
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(page)
    await expect
      .poll(() => pictures(page), { timeout: 30_000 })
      .toEqual([
        { shown: '200x225', stored: '200x225' },
        { shown: '200x225', stored: '200x225' },
      ])
    expect(meanDifference(await shownColours(page, 0), looked[0])).toBeLessThan(4)
    expect(meanDifference(await shownColours(page, 1), looked[1])).toBeLessThan(4)
  } finally {
    await closeAndSaveVideo(launched, 'docs-reduce-file-size')
    rmSync(dir, { recursive: true, force: true })
  }
})
