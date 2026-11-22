/**
 * Insert > Icons with the caret in a header, a footer and a text box, Graphics Fill
 * on an icon that floats beside text, and Size and Position scaling from the
 * size a 300 dpi picture has at its own resolution. Each is checked on the
 * page, in the saved file, and after the file is opened again.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { crc32, deflateSync } from 'node:zlib'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'
import { clickMenu } from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-icons-anywhere-and-picture-dpi'

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`

/** a document with the body given and the media named by relationship id */
async function docxOf(body: string, media: Record<string, { name: string; bytes: Buffer }> = {}) {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Default Extension="svg" ContentType="image/svg+xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      Object.entries(media)
        .map(
          ([rId, { name }]) =>
            `<Relationship Id="${rId}" Type="${IMAGE_REL}" Target="media/${name}"/>`,
        )
        .join('') +
      '</Relationships>',
  )
  for (const { name, bytes } of Object.values(media)) zip.file(`word/media/${name}`, bytes)
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
      '</w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** a plain PNG of w x h; `pixelsPerMetre` adds the pHYs chunk a file states its density with */
function png(w: number, h: number, pixelsPerMetre?: number): Buffer {
  const raw = Buffer.alloc((w * 4 + 1) * h)
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) raw.set([0, 90, 200, 255], y * (w * 4 + 1) + 1 + x * 4)
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
  const phys = Buffer.alloc(9)
  if (pixelsPerMetre) {
    phys.writeUInt32BE(pixelsPerMetre, 0)
    phys.writeUInt32BE(pixelsPerMetre, 4)
    phys[8] = 1
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', head),
    ...(pixelsPerMetre ? [chunk('pHYs', phys)] : []),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

const emu = (px: number) => px * 9525

const SVG_EXT =
  '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
  '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rIdSvg"/>' +
  '</a:ext></a:extLst>'

const picPic = (id: number, w: number, h: number, svg: boolean) =>
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg">${svg ? SVG_EXT : ''}</a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'

/** a paragraph that is one in-line picture, size in px */
const picturePara = (id: number, w: number, h: number) =>
  '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  `${picPic(id, w, h, false)}</wp:inline></w:drawing></w:r></w:p>`

/** a run that is an icon with text wrapped square around it, 96 px, drawn from an SVG part */
const floatingIconRun = (id: number) =>
  '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
  'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${emu(200)}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(10)}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${emu(96)}" cy="${emu(96)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="Graphic ${id}" descr="Square"/><wp:cNvGraphicFramePr/>` +
  `${picPic(id, 96, 96, true)}</wp:anchor></w:drawing></w:r>`

const ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="96" height="96" viewBox="0 0 24 24"><path d="M2 2h20v20H2z"/></svg>'

const ready = async (page: Page, text: string) => {
  await expect(page.locator('.doc-page p', { hasText: text }).first()).toBeVisible({
    timeout: 30_000,
  })
  // the open-time respell kick focuses the body and shields keystrokes for a moment
  await page.waitForTimeout(1500)
}

/** opens Insert > Icons… from the menu bar, picks the named icons and inserts them */
async function insertIcons(app: ElectronApplication, page: Page, names: string[]) {
  const dialog = page.locator('.icons-dialog')
  await clickMenu(app, ['Insert', 'Icons…'])
  await expect(dialog).toBeVisible()
  // a search runs over every category, and what was picked stays picked
  for (const name of names) {
    await dialog.locator('.icons-dialog-search').fill(name)
    await dialog.locator(`.icons-dialog-cell[data-icon="${name}"]`).click()
  }
  await dialog.locator('button.primary').click()
  await expect(dialog).toBeHidden()
}

/** the size in px each matching picture is shown at, zoom taken out */
const shownSizes = (page: Page, selector: string) =>
  page.evaluate(
    (selector) =>
      Array.from(document.querySelectorAll<HTMLImageElement>(selector)).map((img) => {
        const box = img.getBoundingClientRect()
        const zoom = box.width / img.offsetWidth
        return [Math.round(box.width / zoom), Math.round(box.height / zoom)]
      }),
    selector,
  )

/** every drawing of a saved part, with the media its bitmap and its SVG resolve to in the part's own relationships */
async function drawingsOf(docPath: string, part: string, within?: RegExp) {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const xml = (await zip.file(`word/${part}`)?.async('string')) ?? ''
  const rels = (await zip.file(`word/_rels/${part}.rels`)?.async('string')) ?? ''
  // a text box is itself a drawing: its pictures are read from its content alone
  const scope = within ? (within.exec(xml)?.[0] ?? '') : xml
  const target = (rId: string | undefined) =>
    rId ? (new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] ?? '') : ''
  const drawings = await Promise.all(
    [...scope.matchAll(/<w:drawing>.*?<\/w:drawing>/g)].map(async ([drawing]) => {
      const png = target(/<a:blip r:embed="([^"]+)"/.exec(drawing)?.[1])
      const svg = target(/<asvg:svgBlip [^>]*r:embed="([^"]+)"/.exec(drawing)?.[1])
      return {
        drawing,
        alt: /<wp:docPr [^>]*descr="([^"]*)"/.exec(drawing)?.[1],
        extent: /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(drawing)?.slice(1, 3).map(Number),
        png,
        svg,
        pngBytes: png ? await zip.file(`word/${png}`)?.async('uint8array') : undefined,
        svgText: svg ? ((await zip.file(`word/${svg}`)?.async('string')) ?? '') : '',
      }
    }),
  )
  return {
    zip,
    xml,
    drawings,
    types: (await zip.file('[Content_Types].xml')?.async('string')) ?? '',
  }
}

const partOf = async (docPath: string, kind: 'header' | 'footer') => {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  return Object.keys(zip.files)
    .map((name) => new RegExp(`^word/(${kind}\\d+\\.xml)$`).exec(name)?.[1])
    .find(Boolean)
}

/**
 * Every XML part of the saved file that a real parser refuses. One prefix that
 * nothing declares is a parse error, and Word then refuses the whole file; the
 * page's own reader is lenient, so only a real parser shows it.
 */
async function unparseableParts(page: Page, docPath: string) {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const parts = await Promise.all(
    Object.values(zip.files)
      .filter((entry) => !entry.dir && /\.(xml|rels)$/.test(entry.name))
      .map(async (entry) => [entry.name, await entry.async('string')] as const),
  )
  return page.evaluate(
    (parts) =>
      parts.flatMap(([name, xml]) => {
        const error = new DOMParser()
          .parseFromString(xml, 'application/xml')
          .querySelector('parsererror')
        return error ? [`${name}: ${error.textContent}`] : []
      }),
    parts,
  )
}

const TXBX_CONTENT = /<w:txbxContent>.*?<\/w:txbxContent>/
const HEADER_PICTURES = '.page-wrap > .page-hf-header img'
const FOOTER_PICTURES = '.page-wrap > .page-hf-footer img'
const BOX_PICTURES = '.doc-page .doc-textbox img.doc-inline-img'

test('Insert > Icons puts icons into a header, a footer and a text box, saved as SVG with a PNG fallback', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-icons-anywhere-')))
  const docPath = join(dir, 'icons.docx')
  writeFileSync(docPath, await docxOf(['apple', 'banana'].map(para).join('')))
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page, 'banana')

    // the header: the strip opens, and the icons go where its caret is, not into the body
    await clickMenu(app, ['View', 'Header and Footer'])
    const strip = page.locator('.page-hf-editor')
    await expect(strip).toBeVisible()
    await expect(strip).toBeFocused()
    await page.keyboard.type('Logo ')
    await expect(strip).toContainText('Logo')
    await insertIcons(app, page, ['heart', 'rocket'])
    await expect(strip.locator('img.doc-inline-img')).toHaveCount(2)
    await expect(page.locator('.doc-page img.doc-inline-img')).toHaveCount(0)
    await expect
      .poll(() => shownSizes(page, '.page-hf-editor img.doc-inline-img'))
      .toEqual([
        [96, 96],
        [96, 96],
      ])
    // the strip still has the caret: typing goes on after the icons
    await expect(strip).toBeFocused()
    await page.keyboard.type(' end')
    await expect(strip).toContainText('end')
    await page.screenshot({ path: test.info().outputPath('icons-in-header.png') })
    await strip.press('Escape')
    await expect(strip).toHaveCount(0)
    await expect(page.locator(HEADER_PICTURES)).toHaveCount(2)

    // the footer, which the document did not have either
    await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    await page.locator('[data-ribbon-body] .rb-big', { hasText: /^Footer$/ }).click()
    await page.locator('[data-rb-panel] button', { hasText: /^Edit Footer$/ }).click()
    const footerStrip = page.locator('.page-hf-footer .page-hf-editor')
    await expect(footerStrip).toBeFocused()
    await page.keyboard.type('Foot ')
    await expect(footerStrip).toContainText('Foot')
    await insertIcons(app, page, ['heart'])
    await expect(footerStrip.locator('img.doc-inline-img')).toHaveCount(1)
    await footerStrip.press('Escape')
    await expect(footerStrip).toHaveCount(0)
    await expect(page.locator(FOOTER_PICTURES)).toHaveCount(1)

    // a text box: the icon joins the box's text
    await page.locator('.doc-page p', { hasText: 'banana' }).first().click()
    await page.keyboard.press('End')
    await page.locator('.rb-big', { hasText: 'Text Box' }).click()
    await page.locator('button', { hasText: 'Simple Text Box' }).click()
    const box = page.locator('.doc-page .doc-textbox')
    await expect(box).toHaveCount(1)
    await expect
      .poll(() => page.evaluate(() => !!document.activeElement?.closest('.doc-textbox-editor')))
      .toBe(true)
    await page.keyboard.type('Boxed ')
    await insertIcons(app, page, ['star'])
    await expect(page.locator(BOX_PICTURES)).toHaveCount(1)
    await expect(box.locator('.doc-textbox-editor')).toContainText('Boxed')
    // nothing went into the body text beside the box
    await expect(page.locator('.doc-page p', { hasText: 'banana' }).locator('img')).toHaveCount(0)
    await page.screenshot({ path: test.info().outputPath('icon-in-text-box.png') })
    await page.locator('.doc-page p', { hasText: 'apple' }).first().click()

    await saveFromMenu(app)
    await expect.poll(() => partOf(docPath, 'header'), { timeout: 30_000 }).toBeTruthy()
    const headerPart = (await partOf(docPath, 'header'))!
    const header = await drawingsOf(docPath, headerPart)
    expect(header.xml).toContain('Logo ')
    expect(header.xml).not.toContain('aidocs-new-pic')
    expect(header.drawings.map((d) => d.alt)).toEqual(['Heart', 'Rocket'])
    for (const drawing of header.drawings) {
      expect(drawing.extent).toEqual([914400, 914400])
      // the header's own relationships name the media
      expect(drawing.png).toMatch(/^media\/[^/]+\.png$/)
      expect(drawing.svg).toMatch(/^media\/[^/]+\.svg$/)
      expect([...drawing.pngBytes!.slice(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47])
      expect(drawing.svgText).toMatch(/^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg" /)
    }
    expect(new Set(header.drawings.flatMap((d) => [d.png, d.svg])).size).toBe(4)
    expect(header.types).toContain('<Default Extension="svg" ContentType="image/svg+xml"/>')
    const footerPart = (await partOf(docPath, 'footer'))!
    const footer = await drawingsOf(docPath, footerPart)
    expect(footer.xml).toContain('Foot ')
    expect(footer.drawings.map((d) => [d.alt, d.extent])).toEqual([['Heart', [914400, 914400]]])
    expect(footer.drawings[0].svgText).toMatch(/^<svg /)

    await expect
      .poll(async () => (await drawingsOf(docPath, 'document.xml')).xml.includes('Boxed'), {
        timeout: 30_000,
      })
      .toBe(true)
    const body = await drawingsOf(docPath, 'document.xml')
    expect(body.xml).not.toContain('aidocs-new-pic')
    const boxContent = TXBX_CONTENT.exec(body.xml)?.[0] ?? ''
    expect(boxContent).toContain('Boxed')
    expect(boxContent).toContain('descr="Star"')
    const inBox = await drawingsOf(docPath, 'document.xml', TXBX_CONTENT)
    expect(inBox.drawings.map((d) => d.alt)).toEqual(['Star'])
    const [star] = inBox.drawings
    expect(star.extent).toEqual([914400, 914400])
    expect(star.png).toMatch(/^media\/[^/]+\.png$/)
    expect(star.svg).toMatch(/^media\/[^/]+\.svg$/)
    expect(star.svgText).toMatch(/^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg" /)
    expect(star.svgText).not.toBe(header.drawings[0].svgText)
    // the header, the footer and the body with its box are XML another reader accepts
    expect(await unparseableParts(page, docPath)).toEqual([])

    // opened again: the header and the box still show their icons at their size
    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(reopened, 'banana')
    await expect(reopened.locator(HEADER_PICTURES)).toHaveCount(2)
    await expect(reopened.locator('.page-wrap > .page-hf-header')).toContainText('Logo')
    await expect
      .poll(() => shownSizes(reopened, HEADER_PICTURES))
      .toEqual([
        [96, 96],
        [96, 96],
      ])
    await expect(reopened.locator(FOOTER_PICTURES)).toHaveCount(1)
    await expect(reopened.locator(BOX_PICTURES)).toHaveCount(1)
    await expect.poll(() => shownSizes(reopened, BOX_PICTURES)).toEqual([[96, 96]])
    await expect(reopened.locator('.doc-page .doc-textbox')).toContainText('Boxed')
    await reopened.screenshot({ path: test.info().outputPath('icons-reopened.png') })

    // a second save keeps the parts
    await reopened.locator('.doc-page p', { hasText: 'apple' }).first().click()
    await reopened.keyboard.press('End')
    await reopened.keyboard.type(' pie')
    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await drawingsOf(docPath, 'document.xml')).xml.includes('apple pie'), {
        timeout: 30_000,
      })
      .toBe(true)
    const again = await drawingsOf(docPath, headerPart)
    expect(again.drawings.map((d) => [d.alt, d.png, d.svg])).toEqual(
      header.drawings.map((d) => [d.alt, d.png, d.svg]),
    )
    for (const drawing of again.drawings) expect(drawing.svgText).toMatch(/^<svg /)
    const bodyAgain = await drawingsOf(docPath, 'document.xml', TXBX_CONTENT)
    expect(bodyAgain.drawings.map((d) => [d.alt, d.svgText])).toEqual([['Star', star.svgText]])
    expect((await drawingsOf(docPath, footerPart)).drawings.map((d) => d.alt)).toEqual(['Heart'])
    expect(await unparseableParts(reopened, docPath)).toEqual([])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

/** what the page draws the first picture from, and whether it still floats */
const floatingGraphic = (page: Page) =>
  page.evaluate(() => {
    type Node = { type: { name: string }; attrs: Record<string, unknown> }
    const { editor } = (
      window as never as {
        __aidocs: { editor: { state: { doc: { descendants(visit: (node: Node) => void): void } } } }
      }
    ).__aidocs
    const pictures: Node[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docInlineImage') pictures.push(node)
    })
    const attrs = pictures[0]?.attrs ?? {}
    const url = String((attrs.vector as { url?: string } | null)?.url ?? '')
    const data = /^data:image\/svg\+xml;base64,(.*)$/.exec(url)
    return { wrap: attrs.wrap ?? null, svg: data ? atob(data[1]) : url }
  })

test('Graphics Fill recolours an icon that floats beside text and shows the colour it has', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-floating-fill-')))
  const docPath = join(dir, 'floating.docx')
  writeFileSync(
    docPath,
    await docxOf(
      `<w:p><w:r><w:t>Text that wraps around the icon.</w:t></w:r>${floatingIconRun(5)}</w:p>` +
        para('Next paragraph.'),
      {
        rIdImg: { name: 'image1.png', bytes: png(96, 96) },
        rIdSvg: { name: 'image2.svg', bytes: Buffer.from(ICON_SVG) },
      },
    ),
  )
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page, 'Next paragraph.')
    const icon = page.locator('.doc-page img.doc-inline-img').first()
    await expect(icon).toBeVisible()
    expect((await floatingGraphic(page)).wrap).toBeTruthy()
    const placed = await icon.boundingBox()

    await icon.click()
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Picture Format')
    const fill = page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: /^Graphics Fill$/,
    })
    await expect(fill).toBeEnabled()
    const palette = page.locator('.gcp-palette')
    const swatch = (name: string) => palette.locator(`.gcp-swatch[title="${name}"]`).first()

    // the palette opens on the colour the icon has: SVG's black, not No Fill
    await fill.click()
    await expect(palette).toBeVisible()
    await expect(palette.locator('.gcp-swatch.selected').first()).toBeVisible()
    await expect(palette.locator('.gcp-auto')).not.toHaveClass(/selected/)
    await expect(swatch('Red')).not.toHaveClass(/selected/)
    await swatch('Red').click()
    await expect.poll(async () => (await floatingGraphic(page)).svg).toContain('fill="#FF0000"')
    // it stays the floating picture it was, where it was
    expect((await floatingGraphic(page)).wrap).toBeTruthy()
    const after = await icon.boundingBox()
    expect(Math.abs(after!.x - placed!.x)).toBeLessThanOrEqual(1)
    expect(Math.abs(after!.y - placed!.y)).toBeLessThanOrEqual(1)
    // and the palette now opens on red
    await fill.click()
    await expect(swatch('Red')).toHaveClass(/selected/)
    await page.screenshot({ path: test.info().outputPath('floating-graphics-fill.png') })
    await fill.click()
    await expect(palette).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(async () => (await drawingsOf(docPath, 'document.xml')).drawings[0]?.svgText ?? '', {
        timeout: 30_000,
      })
      .toContain('fill="#FF0000"')
    const saved = await drawingsOf(docPath, 'document.xml')
    expect(saved.xml).not.toContain('aidocs-new-pic')
    expect(saved.drawings).toHaveLength(1)
    const [drawing] = saved.drawings
    // the anchored drawing of the file, with its place, wrap and name
    expect(drawing.drawing).toContain('<wp:anchor ')
    expect(drawing.drawing).toContain(`<wp:posOffset>${emu(200)}</wp:posOffset>`)
    expect(drawing.drawing).toContain('<wp:wrapSquare wrapText="bothSides"/>')
    expect(drawing.alt).toBe('Square')
    expect(drawing.extent).toEqual([emu(96), emu(96)])
    // the fallback bitmap was drawn again from the red SVG
    expect(drawing.png).not.toBe('media/image1.png')
    expect([...drawing.pngBytes!.slice(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47])

    // opened again: red on the page, red in the palette
    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(reopened, 'Next paragraph.')
    const again = reopened.locator('.doc-page img.doc-inline-img').first()
    await expect(again).toBeVisible()
    expect((await floatingGraphic(reopened)).wrap).toBeTruthy()
    await again.click()
    await expect(reopened.locator('.ribbon-tab.active')).toHaveText('Picture Format')
    await reopened
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Graphics Fill$/ })
      .click()
    await expect(reopened.locator('.gcp-palette .gcp-swatch[title="Red"]').first()).toHaveClass(
      /selected/,
    )
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const BODY_PICTURES = '.doc-page[contenteditable="true"] img'

test('Size and Position scales from the size a 300 dpi picture has at its own resolution', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-picture-dpi-')))
  const docPath = join(dir, 'dpi.docx')
  // 600 x 300 px at 300 dpi is 2 x 1 inch: 192 x 96 px on the page
  writeFileSync(
    docPath,
    await docxOf(para('Before') + picturePara(7, 192, 96) + para('After'), {
      rIdImg: { name: 'image1.png', bytes: png(600, 300, 11811) },
    }),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page, 'After')
    await expect.poll(() => shownSizes(page, BODY_PICTURES)).toEqual([[192, 96]])

    await page.locator(BODY_PICTURES).first().click({ button: 'right' })
    const menu = page.locator('.ctx-menu')
    await menu.waitFor({ timeout: 15_000 })
    await menu.locator('.ctx-item', { hasText: /^Size and Position/ }).click()
    const size = page.locator('.size-position-dialog')
    const scaleH = size.locator('input[name="scaleHeight"]')
    const scaleW = size.locator('input[name="scaleWidth"]')
    await expect(size.locator('input[name="fromOriginal"]')).toBeChecked()
    // at 96 dpi this picture would read 32%
    await expect(scaleW).toHaveValue('100')
    await expect(scaleH).toHaveValue('100')
    await page.screenshot({ path: test.info().outputPath('size-dialog-300dpi.png') })
    await scaleW.fill('50')
    await expect(scaleH).toHaveValue('50')
    await expect(size.getByLabel('Width', { exact: true })).toHaveValue(/^(1"|2\.54 cm)$/)
    await size.locator('button.primary').click()
    await expect(size).toBeHidden()
    await expect.poll(() => shownSizes(page, BODY_PICTURES)).toEqual([[96, 48]])

    await saveFromMenu(app)
    await expect
      .poll(async () => (await drawingsOf(docPath, 'document.xml')).drawings[0]?.extent, {
        timeout: 30_000,
      })
      .toEqual([emu(96), emu(48)])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
