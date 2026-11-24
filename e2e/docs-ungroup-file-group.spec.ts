/**
 * Ungroup on a group that came with the file, through the real press and
 * ribbon path: every child becomes an object of its own where it stood, the
 * text keeps its lines, and the file holds the children as drawings of the
 * group's paragraph. Turned, mirrored and nested groups included.
 */
import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import {
  docxWith,
  groupParagraph,
  groupPicture,
  groupShape,
  innerGroup,
  para,
  rects,
  saveFromMenu,
} from './docs-objects-fixture'

const BOXES = '.doc-page .doc-textbox'
const NODES = '.doc-page .doc-protected:has(.doc-textbox)'
const WORDS = ['Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten']

const withGroup = (group: string): Promise<Buffer> =>
  docxWith(['One', 'Two'].map(para).join('') + group + WORDS.map(para).join(''))

const documentXml = async (path: string): Promise<string> => {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return (await zip.file('word/document.xml')?.async('string')) ?? ''
}

/** every box as [left, top, width, height] measured from the paper's corner: scrolling does not change it */
async function onPaper(page: Page): Promise<number[][]> {
  const [paper] = await rects(page, '.doc-page')
  return (await rects(page, BOXES)).map((r) => [
    r.left - paper.left,
    r.top - paper.top,
    r.width,
    r.height,
  ])
}

/** the top of every text paragraph of the body, from the paper's corner */
const textTops = (page: Page): Promise<number[]> =>
  page.evaluate(() => {
    const paper = document.querySelector('.doc-page')!.getBoundingClientRect().top
    return Array.from(document.querySelectorAll('.doc-page > p'), (p) =>
      p.textContent ? p.getBoundingClientRect().top - paper : null,
    ).filter((top): top is number => top !== null)
  })

const near = (a: number[][], b: number[][], by: number): boolean =>
  a.length === b.length && a.every((box, i) => box.every((v, k) => Math.abs(v - b[i][k]) <= by))

const offsets = (xml: string, axis: 'H' | 'V'): number[] =>
  Array.from(
    xml.matchAll(new RegExp(`<wp:position${axis}[^>]*><wp:posOffset>(-?\\d+)</wp:posOffset>`, 'g')),
  ).map((m) => Math.round(Number(m[1]) / 9525))

async function ungroup(page: Page): Promise<void> {
  await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
  await page.locator('.rb-big', { hasText: /^Group$/ }).click()
  await page.locator('.layout-menu button', { hasText: /^Ungroup$/ }).click()
}

const selected = (page: Page): Promise<[number, number]> =>
  page.evaluate(() => [
    document.querySelectorAll('.doc-page .ProseMirror-selectednode').length,
    document.querySelectorAll('.doc-page .doc-multi-selected').length,
  ])

const undoRedo = (app: Awaited<ReturnType<typeof launchShell>>['app'], label: string) =>
  app.evaluate(({ Menu }, label) => {
    const edit = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')
    edit?.submenu?.items.find((item) => item.label.startsWith(label))?.click()
  }, label)

test('a group from the file ungroups into objects that select, move, format and delete one by one', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ungroup-')))
  const docPath = join(dir, 'group.docx')
  // a text box, an ellipse, a turned triangle and a picture in a 300 x 170 group
  const children =
    groupShape(1, [0, 0, 100, 60], 'C00000', { text: 'Box' }) +
    groupShape(2, [140, 20, 80, 80], '00B050', { prst: 'ellipse' }) +
    groupShape(3, [60, 110, 120, 40], '4472C4', { prst: 'triangle', attrs: ' rot="1800000"' }) +
    groupPicture(4, [240, 100, 60, 60])
  const source =
    ['One', 'Two'].map(para).join('') + groupParagraph(50, [60, 20, 300, 170], children)
  writeFileSync(docPath, await withGroup(groupParagraph(50, [60, 20, 300, 170], children)))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-ungroup-file-group',
    openFile: docPath,
  })
  const shown = { placed: [] as number[][] }
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    const boxes = page.locator(BOXES)
    await expect(boxes).toHaveCount(4, { timeout: 30_000 })
    await expect(page.locator(NODES)).toHaveCount(1)
    const grouped = await onPaper(page)
    const lines = await textTops(page)
    await page.screenshot({ path: test.info().outputPath('grouped.png') })

    await boxes.nth(1).click()
    await ungroup(page)
    // one object per child, all selected, nothing moved: neither the children nor the text
    await expect(page.locator(NODES)).toHaveCount(4)
    expect(await selected(page)).toEqual([1, 3])
    expect(near(await onPaper(page), grouped, 1)).toBe(true)
    expect((await textTops(page)).map(Math.round)).toEqual(lines.map(Math.round))
    await page.screenshot({ path: test.info().outputPath('ungrouped.png') })

    // one undo puts the group back, redo takes it apart again
    await undoRedo(launched.app, 'Undo')
    await expect(page.locator(NODES)).toHaveCount(1)
    expect(near(await onPaper(page), grouped, 1)).toBe(true)
    await undoRedo(launched.app, 'Redo')
    await expect(page.locator(NODES)).toHaveCount(4)
    expect(near(await onPaper(page), grouped, 1)).toBe(true)

    // each child selects alone, and stays selected (a key and a pointer move do not drop it)
    for (const i of [0, 1, 2, 3]) {
      await page.locator('.doc-page > p', { hasText: 'Ten' }).click()
      await boxes.nth(i).click()
      await page.keyboard.press('Shift')
      const box = (await boxes.nth(i).boundingBox())!
      await page.mouse.move(box.x + box.width / 2 + 4, box.y + box.height / 2)
      await expect.poll(() => selected(page)).toEqual([1, 0])
    }

    // the ellipse moves alone
    await page.locator('.doc-page > p', { hasText: 'Ten' }).click()
    const grab = (await boxes.nth(1).boundingBox())!
    await page.mouse.move(grab.x + grab.width / 2, grab.y + grab.height / 2)
    await page.mouse.down()
    await page.mouse.move(grab.x + grab.width / 2 + 50, grab.y + grab.height / 2 + 30, { steps: 6 })
    await page.mouse.up()
    const moved = grouped.map((box, i) =>
      i === 1 ? [box[0] + 50, box[1] + 30, box[2], box[3]] : box,
    )
    await expect.poll(async () => near(await onPaper(page), moved, 1.5)).toBe(true)

    await boxes.nth(0).click()
    // the text box turns yellow from the Shape Format tab
    const fill = page.locator('.rb-big', { hasText: /^Shape Fill$/ })
    if (!(await fill.isVisible()))
      await page.locator('.ribbon-tab', { hasText: /^Shape Format$/ }).click()
    await fill.click()
    await page.locator('.docs-color-pop .gcp-standard-row .gcp-swatch[title="Yellow"]').click()

    // and grows from its corner grip
    const grip = (await page.locator('.ProseMirror-selectednode .box-resize-handle').boundingBox())!
    await page.mouse.move(grip.x + grip.width / 2, grip.y + grip.height / 2)
    await page.mouse.down()
    await page.mouse.move(grip.x + grip.width / 2 + 30, grip.y + grip.height / 2 + 20, { steps: 5 })
    await page.mouse.up()
    await expect
      .poll(async () => {
        const [box] = await onPaper(page)
        return [Math.round(box[2] - grouped[0][2]), Math.round(box[3] - grouped[0][3])]
      })
      .toEqual([30, 20])
    const resized = await onPaper(page)
    // only it changed
    expect(near(resized.slice(1), moved.slice(1), 1.5)).toBe(true)

    // the picture is deleted alone, and comes back
    await boxes.nth(3).click()
    await page.keyboard.press('Delete')
    await expect(boxes).toHaveCount(3)
    await undoRedo(launched.app, 'Undo')
    await expect(boxes).toHaveCount(4)
    expect(near(await onPaper(page), resized, 1.5)).toBe(true)

    // two of them make a group again
    await boxes.nth(1).click()
    await boxes.nth(3).click({ modifiers: ['Shift'] })
    expect(await selected(page)).toEqual([1, 1])
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await page.locator('.rb-big', { hasText: /^Group$/ }).click()
    await expect(page.locator('.layout-menu button', { hasText: /^Ungroup$/ })).toBeDisabled()
    await page.locator('.layout-menu button', { hasText: /^Group$/ }).click()
    await page.locator('.doc-page > p', { hasText: 'Ten' }).click()
    await boxes.nth(3).click()
    await expect.poll(() => selected(page)).toEqual([1, 1])
    shown.placed = await onPaper(page)
    expect(near(shown.placed, resized, 1.5)).toBe(true)
    expect((await textTops(page)).map(Math.round)).toEqual(lines.map(Math.round))
    await page.screenshot({ path: test.info().outputPath('edited.png') })

    await saveFromMenu(launched.app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [xml.match(/<wp:anchor /g)?.length ?? 0, xml.match(/<wpg:wgp[ >]/g)?.length ?? 0]
        },
        { timeout: 30_000 },
      )
      .toEqual([3, 1])
    const xml = await documentXml(docPath)
    // no paragraph was added: the drawings share the paragraph the group was in
    const bodyParagraphs = (text: string) =>
      text.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, '').match(/<w:p[ >]/g)?.length
    expect(bodyParagraphs(xml)).toBe(bodyParagraphs(source) + WORDS.length)
    const holder =
      /<w:p>(?:(?!<\/w:p>)[\s\S])*?<wp:anchor [\s\S]*?<\/w:p>(?=<w:p>|<w:sectPr|<\/w:body>)/.exec(
        xml.replace(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/g, ''),
      )?.[0]
    expect(holder?.match(/<wp:anchor /g)).toHaveLength(3)
    // the text box where it was, 30 x 20 larger and yellow; the new group at the ellipse's new place; the triangle
    expect(offsets(xml, 'H')).toEqual([60, 250, 120])
    expect(offsets(xml, 'V')).toEqual([20, 70, 130])
    expect(xml).toContain(`<wp:extent cx="${130 * 9525}" cy="${80 * 9525}"/>`)
    expect(xml).toContain('<w:t>Box</w:t>')
    expect(xml.match(/<a:srgbClr val="C00000"\/>/g) ?? []).toHaveLength(0)
    expect(xml.match(/<a:srgbClr val="FFFF00"\/>/g)).toHaveLength(1)
    expect(xml.match(/<a:srgbClr val="(?:00B050|4472C4)"\/>/g)).toHaveLength(2)
    expect(xml.match(/<pic:pic[ >]/g)).toHaveLength(1)
    writeFileSync(test.info().outputPath('saved.docx'), readFileSync(docPath))
    // the save reads the file back: the children are still objects of their own
    await expect(page.locator(NODES)).toHaveCount(4)
    await page.locator('.doc-page > p', { hasText: 'Ten' }).click()
    await boxes.nth(0).click()
    await expect.poll(() => selected(page)).toEqual([1, 0])
    expect(near(await onPaper(page), shown.placed, 1.5)).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'docs-ungroup-file-group')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-ungroup-file-group',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator(BOXES)).toHaveCount(4, { timeout: 30_000 })
    // the file shows what the page showed (in the file's order: the new group holds the ellipse and the picture)
    const reopened = await onPaper(page)
    const expected = [shown.placed[0], shown.placed[1], shown.placed[3], shown.placed[2]]
    expect(near(reopened, expected, 1.5)).toBe(true)
    await page.screenshot({ path: test.info().outputPath('reopened.png') })
  } finally {
    await closeAndSaveVideo(launched, 'docs-ungroup-file-group')
    rmSync(dir, { recursive: true, force: true })
  }
})

const KIDS =
  groupShape(1, [0, 0, 100, 60], 'C00000') +
  groupShape(2, [140, 20, 80, 80], '00B050', { prst: 'ellipse' }) +
  groupShape(3, [60, 110, 120, 40], '4472C4', { prst: 'triangle', attrs: ' rot="1800000"' })

const TURNED: Array<[string, string, string[]]> = [
  ['turned 45 degrees', ' rot="2700000"', ['rot="2700000"', 'rot="2700000"', 'rot="4500000"']],
  ['turned a quarter', ' rot="5400000"', ['rot="5400000"', 'rot="5400000"', 'rot="7200000"']],
  ['mirrored', ' flipH="1"', ['flipH="1"', 'flipH="1"', 'rot="19800000" flipH="1"']],
  [
    'turned and mirrored',
    ' rot="2700000" flipV="1"',
    ['rot="2700000" flipV="1"', 'rot="2700000" flipV="1"', 'rot="900000" flipV="1"'],
  ],
]

for (const [name, attrs, turns] of TURNED) {
  test(`a ${name} group ungroups into children that stay where they showed`, async () => {
    test.setTimeout(120_000)
    const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ungroup-')))
    const docPath = join(dir, 'turned.docx')
    writeFileSync(docPath, await withGroup(groupParagraph(50, [80, 40, 220, 160], KIDS, attrs)))
    writeFileSync(test.info().outputPath('grouped.docx'), readFileSync(docPath))
    let launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-ungroup-file-group',
      openFile: docPath,
    })
    let grouped: number[][] = []
    try {
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      const boxes = page.locator(BOXES)
      await expect(boxes).toHaveCount(3, { timeout: 30_000 })
      grouped = await onPaper(page)
      const lines = await textTops(page)
      await page.screenshot({ path: test.info().outputPath('grouped.png') })
      // the group really is drawn turned: its children are not where an unturned group has them
      const flat = [
        [80, 40],
        [220, 60],
      ]
      const body = await page.evaluate(() => {
        const paper = document.querySelector('.doc-page')!
        const style = getComputedStyle(paper)
        const line = document.querySelector('.doc-page .doc-protected:has(.doc-textbox)')!
        return [
          parseFloat(style.paddingLeft),
          line.getBoundingClientRect().top - paper.getBoundingClientRect().top,
        ]
      })
      expect(
        flat.every(
          ([x, y], i) =>
            Math.abs(grouped[i][0] - body[0] - x) > 5 || Math.abs(grouped[i][1] - body[1] - y) > 5,
        ),
      ).toBe(true)

      await boxes.nth(0).click()
      await ungroup(page)
      await expect(page.locator(NODES)).toHaveCount(3)
      // measured: nothing moved by more than 1 px
      expect(near(await onPaper(page), grouped, 1)).toBe(true)
      expect((await textTops(page)).map(Math.round)).toEqual(lines.map(Math.round))
      await page.screenshot({ path: test.info().outputPath('ungrouped.png') })

      await saveFromMenu(launched.app)
      await expect
        .poll(
          async () => {
            const xml = await documentXml(docPath)
            return [xml.match(/<wp:anchor /g)?.length ?? 0, xml.match(/<wpg:wgp[ >]/g)?.length ?? 0]
          },
          { timeout: 30_000 },
        )
        .toEqual([3, 0])
      const xml = await documentXml(docPath)
      // each child carries the group's turn and mirror composed with its own
      expect(
        Array.from(xml.matchAll(/<wps:spPr><a:xfrm([^>]*)>/g)).map((m) => m[1].trim()),
      ).toEqual(turns)
      writeFileSync(test.info().outputPath('ungrouped.docx'), readFileSync(docPath))
    } finally {
      await closeAndSaveVideo(launched, 'docs-ungroup-file-group')
    }

    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-ungroup-file-group',
      openFile: docPath,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(page.locator(BOXES)).toHaveCount(3, { timeout: 30_000 })
      expect(near(await onPaper(page), grouped, 1)).toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'docs-ungroup-file-group')
      rmSync(dir, { recursive: true, force: true })
    }
  })
}

test('a group the text flows above and below stays a group: its band cannot be kept around separate drawings', async () => {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ungroup-')))
  const docPath = join(dir, 'band.docx')
  writeFileSync(
    docPath,
    await withGroup(groupParagraph(50, [80, 10, 220, 160], KIDS, '', '<wp:wrapTopAndBottom/>')),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-ungroup-file-group',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    const boxes = page.locator(BOXES)
    await expect(boxes).toHaveCount(3, { timeout: 30_000 })
    const grouped = await onPaper(page)
    const lines = await textTops(page)
    // the text after the group starts below it
    expect(lines[2]).toBeGreaterThan(grouped[2][1] + grouped[2][3] - 30)
    await boxes.nth(1).click()
    // nothing here can be grouped or taken apart: the button is greyed out
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await expect(page.locator('.rb-big', { hasText: /^Group$/ })).toBeDisabled()
    await expect(page.locator(BOXES)).toHaveCount(3)
    await expect(page.locator(NODES)).toHaveCount(1)
    expect(near(await onPaper(page), grouped, 1)).toBe(true)
    expect((await textTops(page)).map(Math.round)).toEqual(lines.map(Math.round))
    await saveFromMenu(launched.app)
    await page.waitForTimeout(1500)
    expect((await documentXml(docPath)).match(/<wpg:wgp[ >]/g)).toHaveLength(1)
  } finally {
    await closeAndSaveVideo(launched, 'docs-ungroup-file-group')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a nested group ungroups one level at a time', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ungroup-')))
  const docPath = join(dir, 'nested.docx')
  // an ellipse beside an inner group (turned 20 degrees) of two shapes, in a group turned 30 degrees
  const nested = groupParagraph(
    50,
    [80, 40, 320, 160],
    groupShape(1, [0, 0, 100, 60], 'C00000', { prst: 'ellipse' }) +
      innerGroup(
        9,
        [150, 30, 170, 130],
        groupShape(2, [150, 30, 80, 80], '00B050', { prst: 'ellipse' }) +
          groupShape(3, [240, 110, 80, 50], '4472C4', { prst: 'triangle' }),
        ' rot="1200000"',
      ),
    ' rot="1800000"',
  )
  writeFileSync(docPath, await withGroup(nested))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-ungroup-file-group',
    openFile: docPath,
  })
  const before = { grouped: [] as number[][] }
  const anchorsAndGroups = async (): Promise<number[]> => {
    const xml = await documentXml(docPath)
    return [xml.match(/<wp:anchor /g)?.length ?? 0, xml.match(/<wpg:wgp[ >]/g)?.length ?? 0]
  }
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    const boxes = page.locator(BOXES)
    await expect(boxes).toHaveCount(3, { timeout: 30_000 })
    before.grouped = await onPaper(page)
    await page.screenshot({ path: test.info().outputPath('grouped.png') })

    await boxes.nth(0).click()
    await ungroup(page)
    // the ellipse and the inner group, which is still one object
    await expect(page.locator(NODES)).toHaveCount(2)
    expect(near(await onPaper(page), before.grouped, 1)).toBe(true)
    await page.locator('.doc-page > p', { hasText: 'Ten' }).click()
    await boxes.nth(2).click()
    await expect.poll(() => selected(page)).toEqual([1, 0])
    await saveFromMenu(launched.app)
    await expect.poll(anchorsAndGroups, { timeout: 30_000 }).toEqual([2, 1])
    expect(await documentXml(docPath)).not.toContain('wpg:grpSp>')

    // the inner group comes apart in its turn
    await boxes.nth(2).click()
    await ungroup(page)
    await expect(page.locator(NODES)).toHaveCount(3)
    expect(near(await onPaper(page), before.grouped, 1)).toBe(true)
    await page.screenshot({ path: test.info().outputPath('ungrouped.png') })
    await saveFromMenu(launched.app)
    await expect.poll(anchorsAndGroups, { timeout: 30_000 }).toEqual([3, 0])
    const xml = await documentXml(docPath)
    expect(Array.from(xml.matchAll(/<wps:spPr><a:xfrm([^>]*)>/g)).map((m) => m[1].trim())).toEqual([
      'rot="1800000"',
      'rot="3000000"',
      'rot="3000000"',
    ])
  } finally {
    await closeAndSaveVideo(launched, 'docs-ungroup-file-group')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-ungroup-file-group',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator(BOXES)).toHaveCount(3, { timeout: 30_000 })
    expect(near(await onPaper(page), before.grouped, 1)).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'docs-ungroup-file-group')
    rmSync(dir, { recursive: true, force: true })
  }
})
