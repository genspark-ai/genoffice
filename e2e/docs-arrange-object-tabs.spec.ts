/**
 * Picture Format and Shape Format carry the Arrange group of the Layout tab,
 * whatever is selected: a picture in a line of text, a floating picture, a
 * picture frame, a shape, a text box. Each control is on or off by what the
 * object can do, and what it does is checked on the page, in the saved file
 * and after reopening.
 */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { ElectronApplication } from 'playwright'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, pictureRun, saveFromMenu, shapeParagraph } from './docs-objects-fixture'

const emu = (px: number) => px * 9525

const inlinePicture = (id: number, w: number, h: number) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`

/** a paragraph holding one text box in front of the text, px from the column and the paragraph */
const textBoxParagraph = (id: number, x: number, y: number, w: number, h: number, text: string) =>
  `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
  `relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/>` +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${emu(x)}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(y)}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:wrapNone/><wp:docPr id="${id}" name="Text Box ${id}"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>` +
  `<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>` +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
  `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`

/** a rectangle with a word in it: the save writes back what is changed on a shape that holds text */
const labelledShape = (id: number, x: number, y: number, w: number, h: number) =>
  shapeParagraph(id, x, y, w, h).replace(
    '<wps:bodyPr/>',
    '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Shape</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/>',
  )

const documentXml = async (file: string): Promise<string> => {
  const zip = await JSZip.loadAsync(readFileSync(file))
  return (await zip.file('word/document.xml')?.async('string')) ?? ''
}

const ready = (page: Page) =>
  page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )

const PICTURES = '.doc-page[contenteditable="true"] img'
const ARRANGE = [
  'Position',
  'Wrap Text',
  'Bring Forward',
  'Send Backward',
  'Align',
  'Group',
  'Rotate',
]

const control = (page: Page, tip: string) =>
  page.locator(`[data-ribbon-body] button[data-tip="${tip}"]`)

/** which of the Arrange buttons are live, in the group's order, and the Selection Pane after them */
const live = (page: Page): Promise<string[]> =>
  page.evaluate((names) => {
    const body = document.querySelector('[data-ribbon-body]')!
    const on = names.filter((name) => {
      const button = body.querySelector<HTMLButtonElement>(`button[data-tip="${name}"]`)
      return button && !button.disabled
    })
    const pane = [...body.querySelectorAll<HTMLButtonElement>('button')].find((button) =>
      button.textContent?.includes('Selection Pane'),
    )
    return pane && !pane.disabled ? [...on, 'Selection Pane'] : on
  }, ARRANGE)

/** opens one of the group's menu buttons and returns its items: caption, and whether it can be picked */
async function menuOf(page: Page, tip: string): Promise<Record<string, boolean>> {
  await control(page, tip).click()
  const panel = page.locator('[data-ribbon-body] [data-rb-panel]')
  await expect(panel).toBeVisible()
  return panel.evaluate((el) =>
    Object.fromEntries(
      [...el.querySelectorAll<HTMLButtonElement>('button')]
        .filter((button) => button.textContent?.trim())
        .map((button) => [button.textContent!.trim(), !button.disabled]),
    ),
  )
}

async function pick(page: Page, tip: string, item: string) {
  await control(page, tip).click()
  await page.locator('[data-ribbon-body] [data-rb-panel] button', { hasText: item }).click()
  await expect(page.locator('[data-ribbon-body] [data-rb-panel]')).toHaveCount(0)
}

interface Objects {
  runs: { rotDeg: number | null; flipH: boolean; flipV: boolean; wrap: string | null }[]
  blocks: {
    rotDeg: number | null
    flipV: boolean
    wrap: string | null
    posH: string | null
    posV: string | null
  }[]
  boxes: { rotDeg: number | null; behind: boolean }[]
}

/** what the editor holds of every picture and shape, in document order */
const objects = (page: Page): Promise<Objects> =>
  page.evaluate(() => {
    type Attrs = Record<string, unknown>
    type Node = { type: { name: string }; attrs: Attrs }
    const editor = (
      window as unknown as {
        __aidocs: { editor: { state: { doc: { descendants(f: (n: Node) => void): void } } } }
      }
    ).__aidocs.editor
    const out: Objects = { runs: [], blocks: [], boxes: [] }
    editor.state.doc.descendants((node) => {
      const a = node.attrs
      if (node.type.name === 'docInlineImage')
        out.runs.push({
          rotDeg: (a.rotDeg as number | null) ?? null,
          flipH: !!a.flipH,
          flipV: !!a.flipV,
          wrap: (a.wrap as string | null) ?? null,
        })
      if (node.type.name !== 'docProtected') return
      if (a.blockType === 'image')
        out.blocks.push({
          rotDeg: (a.imageRotDeg as number | null) ?? null,
          flipV: !!a.imageFlipV,
          wrap: (a.imageWrap as string | null) ?? null,
          posH: (a.imagePosH as string | null) ?? null,
          posV: (a.imagePosV as string | null) ?? null,
        })
      for (const box of (a.textboxes as Attrs[] | null) ?? [])
        out.boxes.push({ rotDeg: (box.rotDeg as number | null) ?? null, behind: !!box.behind })
    })
    return out
  })

const size = async (page: Page, nth: number): Promise<number[]> => {
  const box = await page.locator(PICTURES).nth(nth).boundingBox()
  return box ? [Math.round(box.width), Math.round(box.height)] : []
}

async function resize(app: ElectronApplication, page: Page, width: number): Promise<void> {
  await app.evaluate(({ BrowserWindow }, w) => {
    for (const win of BrowserWindow.getAllWindows()) win.setContentSize(w, 800)
    return w
  }, width)
  await expect
    .poll(() => page.evaluate(() => window.innerWidth), { timeout: 5_000 })
    .toBe(width)
    .catch(() => {})
  await page.evaluate(
    () => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))),
  )
}

const ALL_BUT_GROUP = [
  'Position',
  'Wrap Text',
  'Bring Forward',
  'Send Backward',
  'Align',
  'Rotate',
  'Selection Pane',
]
const OF_A_SHAPE = [
  'Wrap Text',
  'Bring Forward',
  'Send Backward',
  'Align',
  'Rotate',
  'Selection Pane',
]
const ROTATE = ['Rotate Right 90°', 'Rotate Left 90°', 'Flip Vertical', 'Flip Horizontal']

test('the Arrange group is on Picture Format and Shape Format for every kind of object, and what it does is saved', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-arrange-tabs-')))
  const docPath = join(dir, 'objects.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('One') +
        `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${inlinePicture(1, 80, 60)}<w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>` +
        para('Two') +
        `<w:p>${pictureRun(2, 'left', 0, 120, 90)}</w:p>` +
        para('Three') +
        para('Four') +
        para('Five') +
        para('Six') +
        para('Seven') +
        labelledShape(3, 300, 20, 120, 80) +
        para('Eight') +
        para('Nine') +
        para('Ten') +
        para('Eleven') +
        para('Twelve') +
        textBoxParagraph(4, 300, 20, 140, 60, 'Boxed') +
        para('Last'),
    ),
  )
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-arrange-object-tabs',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const pictures = page.locator(PICTURES)
    await expect(pictures).toHaveCount(2, { timeout: 30_000 })
    const tab = page.locator('.ribbon-tab.active')

    // a picture in a line of text: no rank, but it takes a place, wraps, turns and mirrors
    await pictures.nth(0).click()
    await expect(tab).toHaveText('Picture Format')
    expect(await live(page)).toEqual(['Position', 'Wrap Text', 'Align', 'Rotate', 'Selection Pane'])
    expect(await menuOf(page, 'Rotate')).toEqual(Object.fromEntries(ROTATE.map((k) => [k, true])))
    await page.keyboard.press('Escape')
    await page.screenshot({ path: test.info().outputPath('inline-picture.png') })
    await pick(page, 'Rotate', 'Rotate Right 90°')
    // the line makes room for the turned picture
    await expect.poll(() => size(page, 0)).toEqual([60, 80])
    await pick(page, 'Rotate', 'Flip Horizontal')
    await expect
      .poll(async () => (await objects(page)).runs[0])
      .toEqual({ rotDeg: 90, flipH: true, flipV: false, wrap: null })
    await expect(pictures.nth(0)).toHaveAttribute('style', /rotate\(90deg\).*scaleX\(-1\)/)
    // every wrap is under Wrap Text; choosing one makes the picture float
    expect(Object.keys(await menuOf(page, 'Wrap Text'))).toEqual([
      'In Line with Text',
      'Square · picture left',
      'Square · picture right',
      'Top and Bottom',
      'Behind Text',
      'In Front of Text',
    ])
    await page.keyboard.press('Escape')
    await pick(page, 'Wrap Text', 'Square · picture left')
    await expect.poll(async () => (await objects(page)).runs[0].wrap).toBe('square-left')
    await expect(tab).toHaveText('Picture Format')
    // floating among the text, it has a rank
    expect(await live(page)).toEqual(ALL_BUT_GROUP)

    // a floating picture: everything but Group
    await pictures.nth(1).click()
    await expect(tab).toHaveText('Picture Format')
    await expect.poll(() => live(page)).toEqual(ALL_BUT_GROUP)
    await page.screenshot({ path: test.info().outputPath('floating-picture.png') })
    expect(await menuOf(page, 'Bring Forward')).toEqual({
      'Bring Forward': true,
      'Bring to Front': true,
    })
    await page.keyboard.press('Escape')
    await control(page, 'Position').click()
    await page.locator('[data-rb-panel] .pos-cell.ph-right.pv-bottom').click()
    await pick(page, 'Rotate', 'Rotate Right 90°')
    await pick(page, 'Rotate', 'Flip Vertical')
    await expect
      .poll(async () => (await objects(page)).blocks[0])
      .toEqual({ rotDeg: 90, flipV: true, wrap: 'square-right', posH: 'right', posV: 'bottom' })

    // a shape: no preset place; it turns and mirrors
    await page
      .locator('.doc-page .doc-textbox')
      .nth(0)
      .click({ position: { x: 3, y: 3 } })
    await expect(tab).toHaveText('Shape Format')
    await expect.poll(() => live(page)).toEqual(OF_A_SHAPE)
    expect(await menuOf(page, 'Rotate')).toEqual({
      'Rotate Right 90°': true,
      'Rotate Left 90°': true,
      'Flip Vertical': true,
      'Flip Horizontal': true,
    })
    await page.keyboard.press('Escape')
    await page.screenshot({ path: test.info().outputPath('shape.png') })
    await pick(page, 'Rotate', 'Rotate Right 90°')
    await pick(page, 'Wrap Text', 'Behind Text')
    await expect
      .poll(async () => (await objects(page)).boxes[0])
      .toEqual({ rotDeg: 90, behind: true })

    // a text box: the same group
    await page
      .locator('.doc-page .doc-textbox')
      .nth(1)
      .click({ position: { x: 3, y: 3 } })
    await expect(tab).toHaveText('Shape Format')
    await expect.poll(() => live(page)).toEqual(OF_A_SHAPE)
    await pick(page, 'Rotate', 'Rotate Left 90°')
    await expect
      .poll(async () => (await objects(page)).boxes[1])
      .toEqual({ rotDeg: 270, behind: false })
    await pick(page, 'Send Backward', 'Send to Back')

    // a narrow window folds the group into one button, and the controls still run from it
    await resize(app, page, 620)
    const folded = page.locator('[data-ribbon-body] .ribbon-group.rb-group-menu')
    await expect(folded.locator('.rb-group-menu-name')).toHaveText('Arrange')
    await folded.locator('button.rb-group-menu-button').click()
    await expect(control(page, 'Rotate')).toBeVisible()
    await page.screenshot({ path: test.info().outputPath('shape-format-folded.png') })
    await pick(page, 'Rotate', 'Rotate Left 90°')
    await expect.poll(async () => (await objects(page)).boxes[1].rotDeg).toBe(180)
    await resize(app, page, 1360)
    await expect(folded).toHaveCount(0)
    await pictures.nth(1).click()
    await expect(tab).toHaveText('Picture Format')
    await resize(app, page, 620)
    await expect(folded.locator('.rb-group-menu-name')).toHaveText('Arrange')
    await resize(app, page, 1360)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const drawing = (id: number) =>
            new RegExp(
              `<w:drawing>(?:(?!</w:drawing>)[\\s\\S])*<wp:docPr id="${id}"[\\s\\S]*?</w:drawing>`,
            ).exec(xml)?.[0] ?? ''
          const [one, two, three, four] = [1, 2, 3, 4].map(drawing)
          return [
            // the picture that was in the line: floating, square, turned and mirrored
            /<wp:anchor\b/.test(one) && /<wp:wrapSquare\b/.test(one),
            /<pic:spPr>\s*<a:xfrm rot="5400000" flipH="1">/.test(one),
            /<wp:extent cx="762000" cy="571500"\/><wp:effectExtent l="0" t="95250" r="0" b="95250"\/>/.test(
              one,
            ),
            // the floating picture: at the preset place, turned and mirrored
            /<wp:positionH relativeFrom="margin"><wp:align>right<\/wp:align>/.test(two),
            /<wp:positionV relativeFrom="margin"><wp:align>bottom<\/wp:align>/.test(two),
            /<a:xfrm rot="5400000" flipV="1">/.test(two),
            // the shape: behind the text, turned
            /behindDoc="1"/.test(three) && /<a:xfrm rot="5400000">/.test(three),
            // the text box: turned twice to the left
            /<a:xfrm rot="10800000">/.test(four),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-arrange-object-tabs')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-arrange-object-tabs',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(page)
    await expect(page.locator(PICTURES)).toHaveCount(2, { timeout: 30_000 })
    const read = await objects(page)
    expect(read.runs).toEqual([{ rotDeg: 90, flipH: true, flipV: false, wrap: 'square-left' }])
    expect(read.blocks).toEqual([
      { rotDeg: 90, flipV: true, wrap: 'square-right', posH: 'right', posV: 'bottom' },
    ])
    expect(read.boxes).toEqual([
      { rotDeg: 90, behind: true },
      { rotDeg: 180, behind: false },
    ])
    // the picture that was in the line is still the one among the text, and still the same tab
    const among = await page.evaluate(
      (selector) =>
        [...document.querySelectorAll(selector)].findIndex((el) => !el.closest('.doc-protected')),
      PICTURES,
    )
    await page.locator(PICTURES).nth(among).click()
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Picture Format')
    expect(await live(page)).toEqual(ALL_BUT_GROUP)
    await page.screenshot({ path: test.info().outputPath('reopened.png') })
  } finally {
    await closeAndSaveVideo(launched, 'docs-arrange-object-tabs')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a picture frame gets the group with its own wrap, and Rotate when it is a drawing of its own', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-arrange-object-tabs',
    openFile: join(__dirname, 'assets', 'cover-photo-with-textboxes.docx'),
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    const frame = page.locator('.doc-page .doc-textbox-picture').first()
    await frame.waitFor({ timeout: 30_000 })
    const box = (await frame.boundingBox())!
    await page.mouse.click(box.x + 40, box.y + box.height - 40)
    await expect(page.locator('.ribbon-tab.active')).toHaveText('Picture Format')
    expect(await live(page)).toEqual([
      'Wrap Text',
      'Bring Forward',
      'Send Backward',
      'Align',
      'Rotate',
      'Selection Pane',
    ])
    await control(page, 'Wrap Text').click()
    await expect(page.locator('[data-ribbon-body] [data-rb-panel] button.active')).toHaveText(
      'Behind Text',
    )
    await page.screenshot({ path: test.info().outputPath('picture-frame.png') })
  } finally {
    await closeAndSaveVideo(launched, 'docs-arrange-object-tabs')
  }
})

interface Placed {
  /** the picture's box, from the left and top margin and to the right and bottom margin */
  left: number
  top: number
  right: number
  bottom: number
  /** where the paragraph the picture came out of starts, from the top margin */
  textTop: number
}

/** where each picture block stands inside the page margins, in document order */
const placed = (page: Page): Promise<Placed[]> =>
  page.evaluate(() => {
    const sheet = document.querySelector<HTMLElement>('.doc-page[contenteditable="true"]')!
    const box = sheet.getBoundingClientRect()
    const pad = getComputedStyle(sheet)
    const px = (v: string) => parseFloat(v) || 0
    const inner = {
      left: box.left + px(pad.paddingLeft),
      top: box.top + px(pad.paddingTop),
      right: box.right - px(pad.paddingRight),
      bottom: box.bottom - px(pad.paddingBottom),
    }
    const text = [...sheet.children].find((el) => el.textContent?.startsWith('Before'))
    const textTop = Math.round((text?.getBoundingClientRect().top ?? 0) - inner.top)
    return [...sheet.querySelectorAll('.doc-protected-image img')].map((img) => {
      const r = img.getBoundingClientRect()
      return {
        left: Math.round(r.left - inner.left),
        top: Math.round(r.top - inner.top),
        right: Math.round(inner.right - r.right),
        bottom: Math.round(inner.bottom - r.bottom),
        textTop,
      }
    })
  })

const near = (a: number, b: number) => Math.abs(a - b) <= 2

test('Position takes a picture out of its line of text to the preset place, on the page, in the file and after reopening', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-arrange-position-')))
  const docPath = join(dir, 'position.docx')
  const filler = 'The quick brown fox jumps over the lazy dog. '.repeat(5)
  const among = (id: number, lead: string) =>
    `<w:p><w:r><w:t xml:space="preserve">${lead} </w:t></w:r>${inlinePicture(id, 120, 90)}<w:r><w:t xml:space="preserve"> after. ${filler}</w:t></w:r></w:p>`
  writeFileSync(
    docPath,
    await docxWith(
      para('One') + among(1, 'Before') + para(filler) + among(2, 'Second') + para('Last'),
    ),
  )
  const anchors = async () =>
    [...(await documentXml(docPath)).matchAll(/<wp:anchor\b[\s\S]*?<\/wp:anchor>/g)].map((m) => [
      /<wp:positionH relativeFrom="(\w+)"><wp:align>(\w+)</.exec(m[0])?.slice(1).join(' '),
      /<wp:positionV relativeFrom="(\w+)"><wp:align>(\w+)</.exec(m[0])?.slice(1).join(' '),
      /<wp:wrapSquare\b/.test(m[0]),
      /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(m[0])?.slice(1).join('x'),
    ])
  let before: Placed[] = []
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-arrange-object-tabs',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const pictures = page.locator(PICTURES)
    await expect(pictures).toHaveCount(2, { timeout: 30_000 })
    const tab = page.locator('.ribbon-tab.active')
    const shape = async () => {
      const { runs, blocks } = await objects(page)
      return [runs.length, ...blocks.map((b) => `${b.wrap} ${b.posH} ${b.posV}`)]
    }

    // from Picture Format: the middle of the right margin
    await pictures.nth(0).click()
    await expect(tab).toHaveText('Picture Format')
    await expect(control(page, 'Position')).toBeEnabled()
    await control(page, 'Position').click()
    await page.screenshot({ path: test.info().outputPath('position-menu.png') })
    await page.locator('[data-rb-panel] .pos-cell.ph-right.pv-center').click()
    await expect.poll(shape).toEqual([1, 'square-right right center'])
    // it is a picture block now, still selected, and the gallery marks its place
    await expect(tab).toHaveText('Picture Format')
    await expect.poll(() => live(page)).toEqual(ALL_BUT_GROUP)
    await control(page, 'Position').click()
    await expect(page.locator('[data-rb-panel] .pos-cell.active')).toHaveClass(/ph-right pv-center/)
    await page.keyboard.press('Escape')
    await expect
      .poll(async () => {
        const [p] = await placed(page)
        return p && near(p.right, 0) && near(p.top, p.bottom)
      })
      .toBe(true)
    await page.screenshot({ path: test.info().outputPath('position-placed.png') })

    // from Layout: the bottom left corner
    await page.locator(`${PICTURES}:not(.doc-protected-image img)`).click()
    await expect(tab).toHaveText('Picture Format')
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await expect(tab).toHaveText('Layout')
    await expect(control(page, 'Position')).toBeEnabled()
    await control(page, 'Position').click()
    await page.locator('[data-rb-panel] .pos-cell.ph-left.pv-bottom').click()
    await expect.poll(shape).toEqual([0, 'square-right right center', 'square-left left bottom'])
    await expect
      .poll(async () => {
        const p = (await placed(page))[1]
        return p && near(p.left, 0) && near(p.bottom, 0)
      })
      .toBe(true)
    before = await placed(page)

    await saveFromMenu(app)
    await expect.poll(anchors, { timeout: 30_000 }).toEqual([
      ['margin right', 'margin center', true, '1143000x857250'],
      ['margin left', 'margin bottom', true, '1143000x857250'],
    ])
    if (process.env.KEEP_POSITION_DOCX)
      writeFileSync(process.env.KEEP_POSITION_DOCX, readFileSync(docPath))
    // each anchor stands alone in a paragraph just above the text it came out of
    expect(await documentXml(docPath)).toMatch(
      /<w:p>(?:(?!<w:t[ >])[\s\S])*<wp:anchor\b(?:(?!<w:t[ >])[\s\S])*<\/w:p><w:p><w:r><w:t xml:space="preserve">Before {2}after\./,
    )
  } finally {
    await closeAndSaveVideo(launched, 'docs-arrange-object-tabs')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-arrange-object-tabs',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(page)
    await expect(page.locator(PICTURES)).toHaveCount(2, { timeout: 30_000 })
    const read = await objects(page)
    expect(read.runs).toEqual([])
    expect(read.blocks.map((b) => `${b.wrap} ${b.posH} ${b.posV}`)).toEqual([
      'square-right right center',
      'square-left left bottom',
    ])
    // the file opens to the page that was saved
    await expect
      .poll(async () => {
        const now = await placed(page)
        return (
          now.length === before.length &&
          now.every((p, i) =>
            (Object.keys(p) as (keyof Placed)[]).every((k) => near(p[k], before[i][k])),
          )
        )
      })
      .toBe(true)
    await page.screenshot({ path: test.info().outputPath('position-reopened.png') })
  } finally {
    await closeAndSaveVideo(launched, 'docs-arrange-object-tabs')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Undo puts a picture that Position took out of its line back where it was, and Redo places it again', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-arrange-position-undo-')))
  const docPath = join(dir, 'undo.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('One') +
        `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${inlinePicture(1, 120, 90)}<w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>` +
        para('Last'),
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-arrange-object-tabs',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(page)
    const picture = page.locator(PICTURES)
    await expect(picture).toHaveCount(1, { timeout: 30_000 })
    const kinds = async () => {
      const { runs, blocks } = await objects(page)
      return [runs.length, blocks.length]
    }
    const inLine = await picture.boundingBox()
    await picture.click()
    await control(page, 'Position').click()
    await page.locator('[data-rb-panel] .pos-cell.ph-center.pv-center').click()
    await expect.poll(kinds).toEqual([0, 1])
    const centred = async () => {
      const [p] = await placed(page)
      return p && near(p.left, p.right) && near(p.top, p.bottom)
    }
    await expect.poll(centred).toBe(true)

    await page.locator('.doc-page[contenteditable="true"]').focus()
    await page.keyboard.press('ControlOrMeta+z')
    await expect.poll(kinds).toEqual([1, 0])
    await expect
      .poll(async () => {
        const box = await picture.boundingBox()
        return box && inLine && near(box.x, inLine.x) && near(box.y, inLine.y)
      })
      .toBe(true)
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect.poll(kinds).toEqual([0, 1])
    await expect.poll(centred).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'docs-arrange-object-tabs')
    rmSync(dir, { recursive: true, force: true })
  }
})
