/**
 * Layout > Group and Ungroup through the real press and ribbon path: grouped
 * objects select and move as one, the file holds them as one Word group, the
 * group reads back in place, and Ungroup gives each child its own drawing.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import {
  docxWith,
  para,
  pictureRun,
  rects,
  saveFromMenu,
  shapeParagraph,
  type Rect,
} from './docs-objects-fixture'

/**
 * Two shapes, a picture alone in its paragraph and a picture beside text, far
 * enough apart that no two are beside each other (side floats cannot overlap
 * on the canvas).
 */
const docxWithObjects = (): Promise<Buffer> =>
  docxWith(
    para('First') +
      shapeParagraph(1, 40, 10, 120, 60) +
      ['Second', 'Third', 'Fourth', 'Fifth'].map(para).join('') +
      shapeParagraph(2, 260, 10, 80, 90) +
      ['Sixth', 'Seventh', 'Eighth', 'Ninth', 'Tenth'].map(para).join('') +
      `<w:p>${pictureRun(3, 30, 10, 90, 70)}</w:p>` +
      ['Eleventh', 'Twelfth', 'Thirteenth', 'Fourteenth', 'Fifteenth'].map(para).join('') +
      `<w:p><w:r><w:t>Text beside a picture that floats in its paragraph. </w:t></w:r>${pictureRun(4, 300, 10, 140, 50)}</w:p>` +
      para('Last'),
  )

const OBJECTS = '.doc-page .doc-textbox, .doc-page .doc-img-wrap, .doc-page img.doc-inline-img'

const documentXml = async (path: string): Promise<string> => {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return (await zip.file('word/document.xml')?.async('string')) ?? ''
}

/** each object's corner measured from the first one's */
const relative = (list: Rect[]): number[][] =>
  list.map((r) => [Math.round(r.left - list[0].left), Math.round(r.top - list[0].top)])

/** each object's corner measured from the paper's */
const onPaper = (list: Rect[], paper: Rect[]): number[][] =>
  list.map((r) => [Math.round(r.left - paper[0].left), Math.round(r.top - paper[0].top)])

const near = (a: number[][], b: number[][], by: number): boolean =>
  a.length === b.length &&
  a.every((pair, i) => Math.abs(pair[0] - b[i][0]) <= by && Math.abs(pair[1] - b[i][1]) <= by)

test('grouped objects select and move as one, save as one Word group, and ungroup again', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-object-group-')))
  const docPath = join(dir, 'objects.docx')
  writeFileSync(docPath, await docxWithObjects())
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-object-group',
    openFile: docPath,
  })
  let placed: number[][] = []
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    const objects = page.locator(OBJECTS)
    await expect(objects).toHaveCount(4, { timeout: 30_000 })
    const start = await rects(page, OBJECTS)
    placed = relative(start)

    await objects.nth(0).click()
    for (const i of [1, 2, 3]) await objects.nth(i).click({ modifiers: ['Shift'] })
    await expect(page.locator('.doc-multi-selected')).toHaveCount(3)
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    const groupButton = page.locator('.rb-big', { hasText: /^Group$/ })
    await groupButton.click()
    // four separate objects: something to group, nothing to ungroup
    const groupItem = page.locator('.layout-menu button', { hasText: /^Group$/ })
    const ungroupItem = page.locator('.layout-menu button', { hasText: /^Ungroup$/ })
    await expect(groupItem).toBeEnabled()
    await expect(ungroupItem).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('group-menu.png') })
    await groupItem.click()

    // the group selects whole: after the selection moved to the text, one member brings the rest
    await page.locator('.doc-page p', { hasText: 'First' }).click()
    await expect(page.locator('.doc-multi-selected')).toHaveCount(0)
    await objects.nth(1).click()
    await expect(page.locator('.doc-multi-selected')).toHaveCount(3)
    await page.screenshot({ path: test.info().outputPath('group-selected.png') })

    // the right-click menu on a member is the group's: the others stay selected, Ungroup is offered
    await objects.nth(1).click({ button: 'right' })
    const groupEntry = page.locator('.ctx-item', { hasText: /^Group/ }).first()
    await expect(groupEntry).toBeEnabled()
    await groupEntry.hover()
    await expect(page.locator('.ctx-submenu .ctx-item', { hasText: /^Ungroup$/ })).toBeEnabled()
    await expect(page.locator('.ctx-submenu .ctx-item', { hasText: /^Group$/ })).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('group-context-menu.png') })
    await page.keyboard.press('Escape')
    await expect(page.locator('.doc-multi-selected')).toHaveCount(3)

    // dragging one member moves the group: every member by the same distance
    const grab = start[2]
    await page.mouse.move(grab.left + grab.width / 2, grab.top + grab.height / 2)
    await page.mouse.down()
    await page.mouse.move(grab.left + grab.width / 2 + 60, grab.top + grab.height / 2 + 40, {
      steps: 6,
    })
    await page.mouse.up()
    await expect
      .poll(async () => {
        const now = await rects(page, OBJECTS)
        const moved = now.map((r, i) => [r.left - start[i].left, r.top - start[i].top])
        return (
          Math.abs(moved[2][0]) > 20 &&
          moved.every((m) => Math.abs(m[0] - moved[2][0]) < 2 && Math.abs(m[1] - moved[2][1]) < 2)
        )
      })
      .toBe(true)
    await expect(page.locator('.doc-multi-selected')).toHaveCount(3)
    await page.screenshot({ path: test.info().outputPath('group-dragged.png') })

    // the drag is one step to undo: every member goes back, and comes forward again
    // Redo is Repeat's wording once there is something to redo: wait for the menu to say so
    const undoRedo = (label: string) =>
      expect
        .poll(() =>
          launched.app.evaluate(({ Menu }, label) => {
            const edit = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')
            const item = edit?.submenu?.items.find((candidate) => candidate.label.startsWith(label))
            item?.click()
            return Boolean(item)
          }, label),
        )
        .toBe(true)
    const dragged = await rects(page, OBJECTS)
    await undoRedo('Undo')
    await expect
      .poll(async () =>
        (await rects(page, OBJECTS)).every(
          (r, i) => Math.abs(r.left - start[i].left) < 2 && Math.abs(r.top - start[i].top) < 2,
        ),
      )
      .toBe(true)
    await undoRedo('Redo')
    await expect
      .poll(async () =>
        (await rects(page, OBJECTS)).every(
          (r, i) => Math.abs(r.left - dragged[i].left) < 2 && Math.abs(r.top - dragged[i].top) < 2,
        ),
      )
      .toBe(true)

    // the file holds one group drawing with the four inside it
    await saveFromMenu(launched.app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const count = (re: RegExp) => xml.match(re)?.length ?? 0
          return [
            count(/<wpg:wgp[ >]/g),
            count(/<wp:anchor /g),
            count(/<wps:wsp>/g),
            count(/<pic:pic[ >]/g),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([1, 1, 2, 2])
  } finally {
    await closeAndSaveVideo(launched, 'docs-object-group')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-object-group',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    // a group is one drawing: its children are its boxes, pictures included
    const boxes = page.locator('.doc-page .doc-protected .doc-textbox')
    await expect(boxes).toHaveCount(4, { timeout: 30_000 })
    await expect(page.locator('.doc-page .doc-protected:has(.doc-textbox)')).toHaveCount(1)
    // the children are where the members were, measured from the first
    await expect
      .poll(async () => near(relative(await rects(page, '.doc-page .doc-textbox')), placed, 2))
      .toBe(true)
    await page.screenshot({ path: test.info().outputPath('group-reopened.png') })

    await boxes.nth(0).click()
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await page.locator('.rb-big', { hasText: /^Group$/ }).click()
    await expect(page.locator('.layout-menu button', { hasText: /^Group$/ })).toBeDisabled()
    await page.locator('.layout-menu button', { hasText: /^Ungroup$/ }).click()
    // nothing moves on the page
    await expect
      .poll(async () => near(relative(await rects(page, '.doc-page .doc-textbox')), placed, 2))
      .toBe(true)

    // and the file holds four drawings of their own again
    await saveFromMenu(launched.app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const count = (re: RegExp) => xml.match(re)?.length ?? 0
          return [count(/<wpg:wgp[ >]/g), count(/<wp:anchor /g)]
        },
        { timeout: 30_000 },
      )
      .toEqual([0, 4])
  } finally {
    await closeAndSaveVideo(launched, 'docs-object-group')
    rmSync(dir, { recursive: true, force: true })
  }
})

/** two shapes with no offset across: placed by alignment, in the column or on the paper */
const ALIGNED = {
  // one centred in its column, one at its right edge
  'in the column': () =>
    para('One') +
    shapeParagraph(1, 'center', 20, 120, 60) +
    ['Two', 'Three', 'Four', 'Five'].map(para).join('') +
    shapeParagraph(2, 'right', 10, 80, 90) +
    para('Last'),
  // the same against the paper, each at a height measured from the paper's top
  'on the paper': () =>
    para('One') +
    shapeParagraph(1, 'center', 260, 120, 60, 'page') +
    ['Two', 'Three', 'Four', 'Five'].map(para).join('') +
    shapeParagraph(2, 'right', 420, 80, 90, 'page') +
    para('Last'),
  // aligned both ways against the paper: no offset on either axis
  'on the paper both ways': () =>
    para('One') +
    shapeParagraph(1, 'center', 'center', 120, 60, 'page') +
    ['Two', 'Three', 'Four', 'Five'].map(para).join('') +
    shapeParagraph(2, 'right', 'bottom', 80, 90, 'page') +
    para('Last'),
  // against the margins, one by offsets and one by alignment
  'against the margins': () =>
    para('One') +
    shapeParagraph(1, 60, 150, 120, 60, 'margin') +
    ['Two', 'Three', 'Four', 'Five'].map(para).join('') +
    shapeParagraph(2, 'right', 'center', 80, 90, 'margin') +
    para('Last'),
}

/** the two objects of those documents before they are a group: a shape is a box, a picture is not yet */
const ALIGNED_OBJECTS = '.doc-page .doc-textbox, .doc-page .doc-img-wrap'

for (const [where, body] of Object.entries(ALIGNED)) {
  test(`shapes placed by alignment ${where} keep their places when they are grouped and saved`, async () => {
    test.setTimeout(180_000)
    const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-object-group-')))
    const docPath = join(dir, 'aligned.docx')
    writeFileSync(docPath, await docxWith(body()))
    let launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-object-group',
      openFile: docPath,
    })
    let placed: number[][] = []
    try {
      const { app } = launched
      const page = await waitForPageWithUrl(app, '://docs/')
      const boxes = page.locator(ALIGNED_OBJECTS)
      await expect(boxes).toHaveCount(2, { timeout: 30_000 })
      const start = await rects(page, ALIGNED_OBJECTS)
      placed = onPaper(start, await rects(page, '.doc-page'))

      await boxes.nth(0).click()
      await boxes.nth(1).click({ modifiers: ['Shift'] })
      await expect(page.locator('.doc-multi-selected')).toHaveCount(1)
      await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
      await page.locator('.rb-big', { hasText: /^Group$/ }).click()
      await page.locator('.layout-menu button', { hasText: /^Group$/ }).click()
      // grouping does not move anything on the paper (the view may scroll to the selection)
      await expect
        .poll(async () =>
          near(
            onPaper(await rects(page, ALIGNED_OBJECTS), await rects(page, '.doc-page')),
            placed,
            2,
          ),
        )
        .toBe(true)

      await saveFromMenu(app)
      await expect
        .poll(
          async () => {
            const xml = await documentXml(docPath)
            // the group is placed by offsets on both axes: alignment could not say where its corner is
            return [
              xml.match(/<wpg:wgp[ >]/g)?.length ?? 0,
              xml.match(/<wp:position[HV][^>]*>\s*<wp:posOffset>/g)?.length ?? 0,
              xml.match(/<wp:position[HV][^>]*>\s*<wp:align>/g)?.length ?? 0,
            ]
          },
          { timeout: 30_000 },
        )
        .toEqual([1, 2, 0])
    } finally {
      await closeAndSaveVideo(launched, 'docs-object-group')
    }

    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-object-group',
      openFile: docPath,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(page.locator('.doc-page .doc-textbox')).toHaveCount(2, { timeout: 30_000 })
      // on the paper where they were, not only where they were against each other
      await expect
        .poll(async () =>
          near(
            onPaper(await rects(page, '.doc-page .doc-textbox'), await rects(page, '.doc-page')),
            placed,
            2,
          ),
        )
        .toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'docs-object-group')
      rmSync(dir, { recursive: true, force: true })
    }
  })
}
