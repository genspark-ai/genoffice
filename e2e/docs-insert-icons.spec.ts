/**
 * Insert > Icons: the dialog opens from the menu bar and from the Insert tab,
 * lists icons by category and by search, and Insert puts the picked icons
 * into the caret's paragraph as one-inch pictures drawn from their SVG and
 * saved as SVG with a PNG fallback; Graphics Fill recolours one.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

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

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]

/** the pictures of the page, without the placeholder image the editor keeps after an inline one */
const PICTURES = '.doc-page img:not(.ProseMirror-separator)'

/** the size in px each picture of the page is shown at, in document order */
const pictureSizes = (page: Page) =>
  page.evaluate(
    (selector) =>
      Array.from(document.querySelectorAll(selector)).map((img) => {
        const box = img.getBoundingClientRect()
        const zoom = box.width / (img as HTMLImageElement).offsetWidth
        return [Math.round(box.width / zoom), Math.round(box.height / zoom)]
      }),
    PICTURES,
  )

/** what each picture of the page is drawn from: its SVG or a bitmap */
const pictureSources = (page: Page) =>
  page.evaluate(
    (selector) =>
      Array.from(document.querySelectorAll<HTMLImageElement>(selector)).map((img) =>
        /^data:image\/svg\+xml|\.svg$/.test(img.src) ? 'svg' : 'bitmap',
      ),
    PICTURES,
  )

/** the first two pictures of the page: on one line, the second right of the first */
const iconRow = (page: Page) =>
  page.evaluate((selector) => {
    const [a, b] = Array.from(document.querySelectorAll(selector)).map((img) =>
      img.getBoundingClientRect(),
    )
    return { sameLine: Math.abs(a.top - b.top) < 2, sideBySide: b.left >= a.right - 1 }
  }, PICTURES)

const BLACK = [0, 0, 0]
const RED = [255, 0, 0]

/** the colour of the most opaque pixel an image holds, drawn at 96 px */
const INK_OF = `async (src) => {
  const image = new Image()
  image.src = src
  await image.decode()
  const canvas = document.createElement('canvas')
  canvas.width = 96
  canvas.height = 96
  const context = canvas.getContext('2d')
  context.drawImage(image, 0, 0, 96, 96)
  const { data } = context.getImageData(0, 0, 96, 96)
  let best = 0
  for (let at = 0; at < data.length; at += 4) if (data[at + 3] > data[best + 3]) best = at
  return [data[best], data[best + 1], data[best + 2]]
}`

/** the ink colour of each picture as the page draws it */
const centreColours = (page: Page) =>
  page.evaluate(
    async ([selector, inkOf]) => {
      const ink = (0, eval)(`(${inkOf})`) as (src: string) => Promise<number[]>
      const out: number[][] = []
      for (const img of Array.from(document.querySelectorAll<HTMLImageElement>(selector))) {
        // a part of the open document is fetched; a data URL is drawn as it is
        const src = img.src.startsWith('data:')
          ? img.src
          : await fetch(img.src)
              .then((r) => r.text())
              .then((svg) => `data:image/svg+xml;base64,${btoa(svg)}`)
        out.push(await ink(src))
      }
      return out
    },
    [PICTURES, INK_OF],
  )

const bitmapColour = (page: Page, pngBase64: string) =>
  page.evaluate(
    ([base64, inkOf]) =>
      ((0, eval)(`(${inkOf})`) as (src: string) => Promise<number[]>)(
        `data:image/png;base64,${base64}`,
      ),
    [pngBase64, INK_OF],
  )

/** per picture of the saved file, in file order: its SVG root's fill, its alt text, its bitmap */
async function filledParts(docPath: string) {
  const saved = await savedPackage(docPath)
  const svgFills: (string | null)[] = []
  const pngs: { size: number[]; base64: string }[] = []
  for (const picture of saved.pictures) {
    const svg = (await saved.zip.file(`word/${picture.svg}`)?.async('string')) ?? ''
    svgFills.push(/^<svg\b[^>]*\sfill="([^"]*)"/.exec(svg)?.[1] ?? null)
    const png = await saved.zip.file(`word/${picture.png}`)!.async('uint8array')
    const view = new DataView(png.buffer, png.byteOffset)
    pngs.push({
      size: [view.getUint32(16), view.getUint32(20)],
      base64: Buffer.from(png).toString('base64'),
    })
  }
  return { svgFills, pngs, alts: saved.pictures.map((p) => p.alt) }
}

const ICON_PATHS = new Map(
  (
    JSON.parse(
      readFileSync(join(__dirname, '../apps/docs/src/renderer/icon-library/icons.json'), 'utf8'),
    ) as { icons: string[][] }
  ).icons.map((icon) => [icon[0], icon[4]]),
)

/**
 * Each inserted picture's bitmap against its icon drawn afresh from an SVG
 * that is itself `size` px wide: the number of pixels whose alpha differs. A
 * bitmap drawn at the inserted size and stretched would differ along every edge.
 */
const fallbackSharpness = (page: Page, names: string[]) =>
  page.evaluate(
    async (paths) => {
      const alphaOf = async (src: string, size: number) => {
        const image = new Image()
        image.src = src
        await image.decode()
        const canvas = document.createElement('canvas')
        canvas.width = size
        canvas.height = size
        const context = canvas.getContext('2d')!
        context.drawImage(image, 0, 0, size, size)
        const { data } = context.getImageData(0, 0, size, size)
        return data.filter((_, index) => index % 4 === 3)
      }
      // the page draws the SVG; the bitmap is what the picture holds for the save
      const bitmaps: string[] = []
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: {
                doc: {
                  descendants(
                    visit: (node: {
                      type: { name: string }
                      attrs: Record<string, unknown>
                    }) => void,
                  ): void
                }
              }
            }
          }
        }
      ).__aidocs
      editor.state.doc.descendants((node) => {
        if (node.type.name === 'docInlineImage') bitmaps.push(String(node.attrs.dataUrl))
      })
      const sizes: number[] = []
      let differing = 0
      let inked = 0
      for (const [index, src] of bitmaps.entries()) {
        const probe = new Image()
        probe.src = src
        await probe.decode()
        const size = probe.naturalWidth
        sizes.push(size)
        const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 24 24"><path d="${paths[index]}"/></svg>`
        const shown = await alphaOf(src, size)
        const fresh = await alphaOf(`data:image/svg+xml;base64,${btoa(svg)}`, size)
        shown.forEach((alpha, at) => {
          if (alpha !== fresh[at]) differing++
          if (alpha === 255) inked++
        })
      }
      return { sizes, differing, inked: inked > 10_000 }
    },
    names.map((name) => ICON_PATHS.get(name)!),
  )

async function savedPackage(docPath: string) {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const text = (name: string) => zip.file(name)!.async('string')
  const document = await text('word/document.xml')
  const rels = await text('word/_rels/document.xml.rels')
  const targetOf = (rId: string) =>
    new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] ?? ''
  const pictures = [...document.matchAll(/<w:drawing>.*?<\/w:drawing>/g)].map(([drawing]) => ({
    extent: /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(drawing)?.slice(1, 3).map(Number),
    alt: /<wp:docPr [^>]*descr="([^"]*)"/.exec(drawing)?.[1],
    inline: drawing.includes('<wp:inline'),
    png: targetOf(/<a:blip r:embed="(rId\d+)"/.exec(drawing)?.[1] ?? ''),
    svg: targetOf(/<asvg:svgBlip [^>]*r:embed="(rId\d+)"/.exec(drawing)?.[1] ?? ''),
  }))
  return { zip, document, types: await text('[Content_Types].xml'), pictures }
}

test('Insert > Icons lists icons by category and search and inserts the picked ones as SVG pictures with a PNG fallback', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-insert-icons-')))
  const docPath = join(dir, 'icons.docx')
  const table =
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${para('cell')}</w:tc></w:tr></w:tbl>`
  writeFileSync(docPath, await docxWith(table + ['apple', 'banana'].map(para).join('')))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-insert-icons',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const banana = page.locator('.doc-page p', { hasText: 'banana' }).first()
    await expect(banana).toBeVisible({ timeout: 30_000 })
    await banana.click()
    await page.keyboard.press('End')

    const dialog = page.locator('.icons-dialog')
    const cells = dialog.locator('.icons-dialog-cell')
    const cell = (name: string) => dialog.locator(`.icons-dialog-cell[data-icon="${name}"]`)
    const category = (id: string) => dialog.locator(`[role="tab"][data-category="${id}"]`)
    const insert = dialog.locator('button.primary')

    // from the menu bar: the first category is listed, nothing is picked
    await clickMenu(app, ['Insert', 'Icons…'])
    await expect(dialog).toBeVisible()
    await expect(category('people')).toHaveAttribute('aria-selected', 'true')
    await expect(cell('person')).toBeVisible()
    await expect(cell('heart')).toHaveCount(0)
    await expect(insert).toBeDisabled()
    const shownInPeople = await cells.count()
    expect(shownInPeople).toBeGreaterThan(10)
    await dialog.locator('button', { hasText: /^Cancel$/ }).click()
    await expect(dialog).toBeHidden()
    await expect(page.locator(PICTURES)).toHaveCount(0)

    // from the Insert tab
    await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    const ribbonIcons = page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: /^Icons$/,
    })
    await expect(ribbonIcons).toBeEnabled()
    await ribbonIcons.click()
    await expect(dialog).toBeVisible()

    // another category, and one icon picked from it
    await category('signs').click()
    await expect(category('signs')).toHaveAttribute('aria-selected', 'true')
    await expect(cell('person')).toHaveCount(0)
    await cell('heart').click()
    await expect(cell('heart')).toHaveAttribute('aria-selected', 'true')
    await expect(insert).toBeEnabled()

    // a search runs over every category; a second click takes a pick back
    const search = dialog.locator('.icons-dialog-search')
    await search.fill('zzzz')
    await expect(cells).toHaveCount(0)
    await expect(dialog.locator('.icons-dialog-empty')).toBeVisible()
    await search.fill('rocket')
    await expect(cells).toHaveCount(1)
    await expect(category('signs')).toHaveAttribute('aria-selected', 'false')
    await cell('rocket').click()
    await search.fill('pie')
    await cell('data_pie').click()
    await cell('data_pie').click()
    await expect(cell('data_pie')).toHaveAttribute('aria-selected', 'false')
    await expect(dialog.locator('.icons-dialog-count')).toHaveText('2 selected')
    await page.screenshot({ path: test.info().outputPath('icons-dialog.png') })

    await insert.click()
    await expect(dialog).toBeHidden()
    await expect(page.locator(PICTURES)).toHaveCount(2)
    await expect
      .poll(() => pictureSizes(page))
      .toEqual([
        [96, 96],
        [96, 96],
      ])
    // what the page draws is the SVG itself, so an enlarged icon stays sharp
    expect(await pictureSources(page)).toEqual(['svg', 'svg'])
    // the fallback bitmap is the vector drawn at 384 px, not a smaller drawing stretched
    expect(await fallbackSharpness(page, ['heart', 'rocket'])).toEqual({
      sizes: [384, 384],
      differing: 0,
      inked: true,
    })
    // both icons went in as one step: one Undo takes them out, one Redo brings them back
    await page.keyboard.press('ControlOrMeta+z')
    await expect(page.locator(PICTURES)).toHaveCount(0)
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect(page.locator(PICTURES)).toHaveCount(2)
    // both icons sit in the paragraph the caret was in, side by side, and typing goes on after them
    await page.keyboard.type('cherry')
    const lines = () =>
      page.evaluate(
        (selector) =>
          Array.from(document.querySelectorAll('.doc-page p')).map((p) => {
            const icons = p.querySelectorAll(selector.replace('.doc-page ', '')).length
            return `${(p.textContent ?? '').trim()}${icons ? ` +${icons}` : ''}`
          }),
        PICTURES,
      )
    await expect.poll(lines).toEqual(['cell', 'apple', 'bananacherry +2'])
    expect(await iconRow(page)).toEqual({ sameLine: true, sideBySide: true })

    // in a table cell the icon joins the cell's paragraph
    await page.locator('.doc-page td p', { hasText: 'cell' }).first().click()
    await page.keyboard.press('End')
    await clickMenu(app, ['Insert', 'Icons…'])
    await expect(dialog).toBeVisible()
    await category('signs').click()
    await cell('star').click()
    await insert.click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.doc-page td img.doc-inline-img')).toHaveCount(1)
    await expect(page.locator(PICTURES)).toHaveCount(3)
    await page.screenshot({ path: test.info().outputPath('icons-inserted.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => (await savedPackage(docPath)).pictures.length, { timeout: 30_000 })
      .toBe(3)
    const saved = await savedPackage(docPath)
    // the cell's icon comes first in the file; it carries its name like the others
    expect(saved.pictures.map((p) => p.alt)).toEqual(['Star', 'Heart', 'Rocket'])
    expect(saved.document.indexOf('descr="Star"')).toBeLessThan(saved.document.indexOf('</w:tbl>'))
    for (const picture of saved.pictures) {
      expect(picture.inline).toBe(true)
      expect(picture.extent).toEqual([914400, 914400])
      expect(picture.png).toMatch(/^media\/[^/]+\.png$/)
      expect(picture.svg).toMatch(/^media\/[^/]+\.svg$/)
      const png = await saved.zip.file(`word/${picture.png}`)!.async('uint8array')
      expect([...png.slice(0, 8)]).toEqual(PNG_SIGNATURE)
      // IHDR width and height: four times the size the icon is shown at
      const view = new DataView(png.buffer, png.byteOffset)
      expect([view.getUint32(16), view.getUint32(20)]).toEqual([384, 384])
      const svg = await saved.zip.file(`word/${picture.svg}`)!.async('string')
      expect(svg).toMatch(
        /^<svg xmlns="http:\/\/www\.w3\.org\/2000\/svg" width="96" height="96" viewBox="0 0 24 24"><path d="[^"]+"\/><\/svg>$/,
      )
    }
    expect(new Set(saved.pictures.flatMap((p) => [p.png, p.svg])).size).toBe(6)
    expect(saved.types).toContain('<Default Extension="svg" ContentType="image/svg+xml"/>')
    expect(saved.types).toContain('<Default Extension="png" ContentType="image/png"/>')
    expect(saved.document).toContain(
      'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"',
    )

    // reopened: the icons are still there at their size, and a second save keeps the SVG parts
    await closeAndSaveVideo(launched, 'docs-insert-icons')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-insert-icons',
      openFile: docPath,
    })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    const cherry = reopened.locator('.doc-page p', { hasText: 'cherry' }).first()
    await expect(cherry).toBeVisible({ timeout: 30_000 })
    await expect(reopened.locator(PICTURES)).toHaveCount(3)
    await expect
      .poll(() => pictureSizes(reopened))
      .toEqual([
        [96, 96],
        [96, 96],
        [96, 96],
      ])
    // a document that carries the SVG is drawn from it too
    expect(await pictureSources(reopened)).toEqual(['svg', 'svg', 'svg'])
    await cherry.click()
    await reopened.keyboard.press('End')
    await reopened.keyboard.type(' pie')
    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await savedPackage(docPath)).document.includes('cherry pie'), {
        timeout: 30_000,
      })
      .toBe(true)
    const again = await savedPackage(docPath)
    expect(again.pictures.map((p) => [p.png, p.svg])).toEqual(
      saved.pictures.map((p) => [p.png, p.svg]),
    )
    for (const picture of again.pictures) {
      expect(again.zip.file(`word/${picture.svg}`)).toBeTruthy()
      expect(again.zip.file(`word/${picture.png}`)).toBeTruthy()
    }

    // Picture Format > Graphics Fill on the heart, an icon that came from the file
    const heart = reopened.locator(PICTURES).nth(1)
    await heart.click()
    await expect(reopened.locator('.ribbon-tab.active')).toHaveText('Picture Format')
    const fill = reopened.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: /^Graphics Fill$/,
    })
    await expect(fill).toBeEnabled()
    await fill.click()
    await reopened.locator('.gcp-palette .gcp-swatch[title="Red"]').click()
    await expect.poll(() => centreColours(reopened)).toEqual([BLACK, RED, BLACK])
    await reopened.screenshot({ path: test.info().outputPath('icons-graphics-fill.png') })
    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await filledParts(docPath)).svgFills, { timeout: 30_000 })
      .toEqual([null, '#FF0000', null])
    const filled = await filledParts(docPath)
    expect(filled.alts).toEqual(['Star', 'Heart', 'Rocket'])
    // the fallback bitmap was drawn again from the red SVG, at four times the shown size
    expect(filled.pngs[1].size).toEqual([384, 384])
    expect(await bitmapColour(reopened, filled.pngs[1].base64)).toEqual(RED)
    expect(await bitmapColour(reopened, filled.pngs[0].base64)).toEqual(BLACK)

    // read back: the heart is red on the page, from its SVG
    await closeAndSaveVideo(launched, 'docs-insert-icons')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-insert-icons',
      openFile: docPath,
    })
    const third = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(third.locator(PICTURES)).toHaveCount(3, { timeout: 30_000 })
    expect(await pictureSources(third)).toEqual(['svg', 'svg', 'svg'])
    await expect.poll(() => centreColours(third)).toEqual([BLACK, RED, BLACK])
  } finally {
    await closeAndSaveVideo(launched, 'docs-insert-icons')
    rmSync(dir, { recursive: true, force: true })
  }
})
