/**
 * Tools > Track Changes > Highlight Changes: "Track moves while tracking
 * changes" is a document setting. Ticked, a whole sentence that is cut and
 * pasted elsewhere is one move on the page, in the Reviewing pane and in the
 * saved file; cleared, the same edit is a deletion and an insertion and the
 * file says so.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-track-moves'
const MOVED = 'This sentence moves. '
const BODY = para('Alpha stays put.') + para(`${MOVED}This one stays.`) + para('Omega ends here.')
// the second paragraph's text starts after the first paragraph (16 characters + 2 boundaries)
const MOVED_FROM = 19
const MOVED_TO = MOVED_FROM + MOVED.length

async function clickMenu(app: ElectronApplication, path: string[]) {
  const click = () =>
    app.evaluate(({ Menu }, path) => {
      let items = Menu.getApplicationMenu()?.items
      let item: Electron.MenuItem | undefined
      for (const label of path) {
        item = items?.find((candidate) => candidate.label.startsWith(label))
        items = item?.submenu?.items
      }
      if (!item?.enabled) return false
      item.click()
      return true
    }, path)
  await expect.poll(click, { timeout: 15_000 }).toBe(true)
}

type Hook = {
  __aidocs: {
    editor: {
      state: { doc: { content: { size: number } } }
      chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } }
    }
  }
}

/** selects document positions; a negative `from` counts back from the end of the document */
async function select(page: Page, from: number, to = from) {
  await page.evaluate(
    ([from, to]) => {
      const { editor } = (window as never as Hook).__aidocs
      const size = editor.state.doc.content.size
      const at = (pos: number) => (pos < 0 ? size + pos : pos)
      editor
        .chain()
        .focus()
        .setTextSelection({ from: at(from), to: at(to) })
        .run()
    },
    [from, to],
  )
  // focus returns to the page a moment after a dialog closes; a key pressed before that is lost
  await expect
    .poll(() =>
      page.evaluate(
        (range) =>
          document.activeElement?.classList.contains('ProseMirror') === true &&
          String(getSelection()).length > 0 === range,
        from !== to,
      ),
    )
    .toBe(true)
}

const clipboard = (page: Page, action: 'cut' | 'paste') =>
  page.keyboard.press(action === 'cut' ? 'ControlOrMeta+x' : 'ControlOrMeta+v')

async function cutAndPasteAtEnd(page: Page, from: number, to: number) {
  await select(page, from, to)
  await clipboard(page, 'cut')
  await expect(page.locator('.doc-page .doc-del').first()).toBeVisible()
  await select(page, -1)
  await clipboard(page, 'paste')
  await expect(page.locator('.doc-page .doc-ins').first()).toBeVisible()
}

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

async function openDoc(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p', { hasText: 'Alpha stays put.' })).toBeVisible({
    timeout: 30_000,
  })
  return { launched, page }
}

async function openReviewingPane(page: Page) {
  await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
  await page
    .locator('[data-ribbon-body]')
    .getByRole('button', { name: 'Reviewing', exact: true })
    .click()
  const pane = page.locator('.review-pane')
  await expect(pane).toBeVisible()
  return pane
}

async function highlightChanges(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Tools', 'Track Changes', 'Highlight Changes'])
  const dialog = page.locator('.hlc-dialog')
  await expect(dialog).toBeVisible()
  return {
    dialog,
    track: dialog.locator('input[name="track-changes"]'),
    moves: dialog.locator('input[name="track-moves"]'),
    ok: async () => {
      await dialog.locator('button.primary').click()
      await expect(dialog).toBeHidden()
    },
  }
}

test('a tracked cut and paste of a sentence is one move, saved as a paired moveFrom / moveTo', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-track-moves-')))
  const docPath = join(dir, 'moves.docx')
  writeFileSync(docPath, await docxWith(BODY))
  let { launched, page } = await openDoc(docPath)
  try {
    const hlc = await highlightChanges(launched.app, page)
    await expect(hlc.moves).toBeEnabled()
    await expect(hlc.moves).toBeChecked()
    await hlc.track.check()
    await hlc.dialog.screenshot({ path: test.info().outputPath('highlight-changes-moves.png') })
    await hlc.ok()

    const from = page.locator('.doc-page .doc-move-from')
    const to = page.locator('.doc-page .doc-move-to')
    await cutAndPasteAtEnd(page, MOVED_FROM, MOVED_TO)
    await expect(from).toHaveText(MOVED)
    await expect(to).toHaveText(MOVED)

    // undo takes the paste and then the cut back, leaving no markup
    await page.keyboard.press('ControlOrMeta+z')
    await expect(to).toHaveCount(0)
    await page.keyboard.press('ControlOrMeta+z')
    await expect(page.locator('.doc-page .doc-del, .doc-page .doc-ins')).toHaveCount(0)
    await expect(page.locator('.doc-page p').nth(1)).toHaveText(`${MOVED}This one stays.`)

    await cutAndPasteAtEnd(page, MOVED_FROM, MOVED_TO)
    await expect(from).toHaveText(MOVED)
    await expect(to).toHaveText(MOVED)

    // less than a sentence stays a deletion and an insertion
    await select(page, 1, 6)
    await clipboard(page, 'cut')
    await expect(page.locator('.doc-page .doc-del:not(.doc-move-from)')).toHaveText('Alpha')
    await select(page, -1)
    await clipboard(page, 'paste')
    await expect(page.locator('.doc-page .doc-ins:not(.doc-move-to)')).toHaveText('Alpha')
    await expect(from).toHaveCount(1)
    await expect(to).toHaveCount(1)

    const pane = await openReviewingPane(page)
    await expect(pane.locator('.review-row-kind', { hasText: 'Moved from' })).toHaveCount(1)
    await expect(pane.locator('.review-row-kind', { hasText: 'Moved to' })).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('tracked-move.png') })

    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).includes('<w:moveTo '), {
        timeout: 20_000,
      })
      .toBe(true)
    const xml = await part(docPath, 'word/document.xml')
    const name = /<w:moveFromRangeStart w:id="(\d+)"[^>]* w:name="([^"]+)"/.exec(xml)
    expect(name).not.toBeNull()
    const [, fromRangeId, moveName] = name!
    const toRange = new RegExp(`<w:moveToRangeStart w:id="(\\d+)"[^>]* w:name="${moveName}"`).exec(
      xml,
    )
    expect(toRange).not.toBeNull()
    expect(xml).toContain(`<w:moveFromRangeEnd w:id="${fromRangeId}"/>`)
    expect(xml).toContain(`<w:moveToRangeEnd w:id="${toRange![1]}"/>`)
    expect(xml).toMatch(
      /<w:moveFromRangeStart [^>]*\/><w:moveFrom [^>]*>(?:(?!<\/w:moveFrom>).)*This sentence moves\. (?:(?!<\/w:moveFrom>).)*<\/w:moveFrom><w:moveFromRangeEnd /,
    )
    expect(xml).toMatch(
      /<w:moveToRangeStart [^>]*\/><w:moveTo [^>]*>(?:(?!<\/w:moveTo>).)*This sentence moves\. (?:(?!<\/w:moveTo>).)*<\/w:moveTo><w:moveToRangeEnd /,
    )
    // every revision and range id in the part is its own
    const ids = [
      ...xml.matchAll(/<w:(?:ins|del|moveFrom|moveTo|move\w+RangeStart) w:id="(\d+)"/g),
    ].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)
    // the word is an ordinary deletion and insertion
    expect(xml).toMatch(/<w:del [^>]*><w:r>(?:(?!<\/w:del>).)*<w:delText[^>]*>Alpha</)
    expect(await part(docPath, 'word/settings.xml')).not.toContain('doNotTrackMoves')

    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, page } = await openDoc(docPath))
    await expect(page.locator('.doc-page .doc-move-from')).toHaveText(MOVED)
    await expect(page.locator('.doc-page .doc-move-to')).toHaveText(MOVED)
    const reopened = await highlightChanges(launched.app, page)
    await expect(reopened.track).toBeChecked()
    await expect(reopened.moves).toBeChecked()
    await reopened.ok()

    // rejecting the destination puts the sentence back where it was
    const reopenedPane = await openReviewingPane(page)
    await expect(reopenedPane.locator('.review-row-kind')).toHaveText([
      'Deleted',
      'Moved from',
      'Moved to',
      'Inserted',
    ])
    const movedTo = reopenedPane.locator('.review-row', { hasText: 'Moved to' })
    await movedTo.click()
    await movedTo.getByRole('button', { name: /Reject/ }).click()
    await expect(page.locator('.doc-page .doc-move-from, .doc-page .doc-move-to')).toHaveCount(0)
    await expect(page.locator('.doc-page p').nth(1)).toHaveText(`${MOVED}This one stays.`)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('with Track moves cleared the same edit is a deletion and an insertion, and the document keeps the setting', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-track-moves-off-')))
  const docPath = join(dir, 'no-moves.docx')
  writeFileSync(docPath, await docxWith(BODY))
  let { launched, page } = await openDoc(docPath)
  try {
    const hlc = await highlightChanges(launched.app, page)
    await hlc.track.check()
    await hlc.moves.uncheck()
    await hlc.ok()

    await cutAndPasteAtEnd(page, MOVED_FROM, MOVED_TO)
    await expect(page.locator('.doc-page .doc-del')).toHaveText(MOVED)
    await expect(page.locator('.doc-page .doc-ins')).toHaveText(MOVED)
    await expect(page.locator('.doc-page .doc-move-from, .doc-page .doc-move-to')).toHaveCount(0)

    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).includes('<w:ins '), {
        timeout: 20_000,
      })
      .toBe(true)
    const xml = await part(docPath, 'word/document.xml')
    expect(xml).toContain('<w:del ')
    expect(xml).not.toMatch(/<w:move(?:From|To)/)
    expect(await part(docPath, 'word/settings.xml')).toMatch(/<w:doNotTrackMoves\s*\/>/)

    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, page } = await openDoc(docPath))
    const reopened = await highlightChanges(launched.app, page)
    await expect(reopened.moves).not.toBeChecked()
    // ticking it again clears the setting from the file
    await reopened.moves.check()
    await reopened.ok()
    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await part(docPath, 'word/settings.xml')).includes('doNotTrackMoves'), {
        timeout: 20_000,
      })
      .toBe(false)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('clearing Track moves leaves a move already in the document drawn as a move', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-track-moves-kept-')))
  const docPath = join(dir, 'kept-move.docx')
  writeFileSync(docPath, await docxWith(BODY))
  let { launched, page } = await openDoc(docPath)
  const lineStyles = (p: Page) =>
    p
      .locator('.doc-page .doc-move-from, .doc-page .doc-move-to')
      .evaluateAll((els) => els.map((el) => getComputedStyle(el).textDecorationStyle))
  const inks = (p: Page) =>
    p
      .locator('.doc-page .doc-move-from, .doc-page .doc-move-to')
      .evaluateAll((els) => els.map((el) => getComputedStyle(el).color))
  try {
    const hlc = await highlightChanges(launched.app, page)
    await hlc.track.check()
    await hlc.ok()
    await cutAndPasteAtEnd(page, MOVED_FROM, MOVED_TO)
    await expect(page.locator('.doc-page .doc-move-from')).toHaveText(MOVED)
    expect(await lineStyles(page)).toEqual(['double', 'double'])
    const moveInks = await inks(page)

    const off = await highlightChanges(launched.app, page)
    await off.moves.uncheck()
    await off.ok()
    await expect(page.locator('.doc-page .doc-move-from')).toHaveText(MOVED)
    await expect(page.locator('.doc-page .doc-move-to')).toHaveText(MOVED)
    await page.screenshot({ path: test.info().outputPath('moves-off-existing-move.png') })
    expect(await lineStyles(page)).toEqual(['double', 'double'])
    expect(await inks(page)).toEqual(moveInks)
    const pane = await openReviewingPane(page)
    await expect(pane.locator('.review-row-kind')).toHaveText(['Moved from', 'Moved to'])

    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await part(docPath, 'word/settings.xml')).includes('doNotTrackMoves'), {
        timeout: 20_000,
      })
      .toBe(true)
    expect(await part(docPath, 'word/document.xml')).toMatch(/<w:moveFrom [^>]*>/)

    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, page } = await openDoc(docPath))
    await expect(page.locator('.doc-page .doc-move-from')).toHaveText(MOVED)
    expect(await lineStyles(page)).toEqual(['double', 'double'])
    expect(await inks(page)).toEqual(moveInks)
    const reopened = await highlightChanges(launched.app, page)
    await expect(reopened.moves).not.toBeChecked()
    await reopened.ok()
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('dragging a selected sentence to another paragraph is one move', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-track-moves-drag-')))
  const docPath = join(dir, 'drag.docx')
  writeFileSync(docPath, await docxWith(BODY))
  const { launched, page } = await openDoc(docPath)
  try {
    const hlc = await highlightChanges(launched.app, page)
    await hlc.track.check()
    await hlc.ok()

    await select(page, MOVED_FROM, MOVED_TO)
    const grab = await page.evaluate(() => {
      const r = getSelection()!.getRangeAt(0).getBoundingClientRect()
      return { x: r.left + r.width / 2, y: r.top + r.height / 2 }
    })
    const last = (await page.locator('.doc-page p').nth(2).boundingBox())!
    const textEnd = await page.evaluate(() => {
      const range = document.createRange()
      range.selectNodeContents(document.querySelectorAll('.doc-page p')[2])
      return range.getBoundingClientRect().right
    })
    // a native text drag cannot be driven from here: the editor gets the drag events themselves
    await page.evaluate(
      ([fromX, fromY, toX, toY]) => {
        const dom = document.querySelector('.doc-page.ProseMirror')!
        const dataTransfer = new DataTransfer()
        const fire = (type: string, clientX: number, clientY: number) =>
          (document.elementFromPoint(clientX, clientY) ?? dom).dispatchEvent(
            new DragEvent(type, {
              dataTransfer,
              clientX,
              clientY,
              bubbles: true,
              cancelable: true,
            }),
          )
        fire('dragstart', fromX, fromY)
        fire('dragover', toX, toY)
        fire('drop', toX, toY)
      },
      [grab.x, grab.y, textEnd - 1, last.y + last.height / 2],
    )

    await expect(page.locator('.doc-page .doc-move-from')).toHaveText(MOVED)
    await expect(page.locator('.doc-page p').nth(2).locator('.doc-move-to')).toHaveText(MOVED)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
