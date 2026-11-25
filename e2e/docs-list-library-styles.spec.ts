/**
 * Format > Bullets and Numbering: Customize on a library card changes that
 * card and Reset puts the built-in one back; the List Styles tab adds a named
 * list style to the document and deletes it again. Both are checked on the
 * page and in the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import {
  copyFileSync,
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

const VIDEO_DIR = 'docs-list-library-styles'
const BODY = para('alpha') + para('beta') + para('gamma') + para('delta')
const ARROW = '→'
const SQUARE = '■'

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

type EditorWindow = {
  __aidocs: {
    editor: {
      state: {
        doc: { descendants(f: (n: { textContent: string }, pos: number) => boolean): void }
      }
      chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } }
    }
  }
}

/** selects from the first paragraph through the paragraph reading `through` */
const selectThrough = (page: Page, through: string) =>
  page.evaluate((through) => {
    const { editor } = (window as never as EditorWindow).__aidocs
    let to = 1
    editor.state.doc.descendants((node, pos) => {
      if (node.textContent === through) to = pos + 1 + through.length
      return false
    })
    editor.chain().focus().setTextSelection({ from: 1, to }).run()
  }, through)

async function openDocs(docPath: string, userDataDir?: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    userDataDir,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page').first()).toContainText('alpha', { timeout: 30_000 })
  return { launched, page }
}

async function openDialog(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Format', 'Bullets and Numbering…'])
  const dialog = page.locator('.bn-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

const markers = (page: Page) =>
  page
    .locator('.doc-page .doc-li')
    .evaluateAll((items) => items.map((el) => el.getAttribute('data-marker')))

test('Customize changes the picked library card, the ribbon gallery follows, and Reset restores the built-in card', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-list-library-')))
  const docPath = join(dir, 'plain.docx')
  writeFileSync(docPath, await docxWith(BODY))
  const first = await openDocs(docPath)
  try {
    const { launched, page } = first
    await selectThrough(page, 'beta')
    let dialog = await openDialog(launched.app, page)
    const cards = dialog.locator('.bn-grid .list-gallery-card')
    const reset = dialog.locator('button[name="reset"]')
    await expect(reset).toBeDisabled()
    await cards.nth(3).click()
    await expect(cards.nth(3)).toHaveText(SQUARE)
    // a built-in card has nothing to reset
    await expect(reset).toBeDisabled()
    await dialog.getByRole('button', { name: 'Customize…' }).click()
    await expect(dialog).toBeHidden()

    const define = page.locator('.modal.list-bullet-dialog')
    await expect(define).toBeVisible()
    // the dialog starts from the card, not from a blank bullet
    await expect(define.locator('.list-dialog-symbols [aria-selected=true]')).toHaveText(SQUARE)
    await define.locator('.list-dialog-symbols [role=option]', { hasText: ARROW }).first().click()
    await define.locator('button.primary').click()
    await expect(define).toBeHidden()
    await expect.poll(() => markers(page)).toEqual([ARROW, ARROW])

    dialog = await openDialog(launched.app, page)
    await expect(cards.nth(3)).toHaveText(ARROW)
    await expect(cards.nth(2)).not.toHaveText(ARROW)
    await expect(reset).toBeDisabled()
    await cards.nth(2).click()
    await expect(reset).toBeDisabled()
    await cards.nth(3).click()
    await expect(reset).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('card-customized.png') })
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    // the ribbon's bullet library shows the same cards
    await page.locator('button.rb-caret[aria-label="Bullets"]').click()
    const gallery = page.locator('.list-gallery-word .list-gallery-glyph')
    await expect(gallery.filter({ hasText: ARROW })).not.toHaveCount(0)
    await expect(gallery.filter({ hasText: SQUARE })).toHaveCount(0)
    await page.keyboard.press('Escape')

    await saveFromMenu(launched.app)
    await expect
      .poll(() => part(docPath, 'word/numbering.xml'), { timeout: 30_000 })
      .toContain(`<w:lvlText w:val="${ARROW}"/>`)
    expect((await part(docPath, 'word/document.xml')).match(/<w:numPr>/g)).toHaveLength(2)
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  // the customized card is the app's, not the document's: it is there in the next session
  const second = await openDocs(docPath, first.launched.userDataDir)
  try {
    const { launched, page } = second
    await expect.poll(() => markers(page)).toEqual([ARROW, ARROW])
    await selectThrough(page, 'alpha')
    const dialog = await openDialog(launched.app, page)
    const cards = dialog.locator('.bn-grid .list-gallery-card')
    const reset = dialog.locator('button[name="reset"]')
    await expect(cards.nth(3)).toHaveText(ARROW)
    await cards.nth(3).click()
    await reset.click()
    await expect(cards.nth(3)).toHaveText(SQUARE)
    await expect(reset).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('card-reset.png') })
    // the card stays picked: OK applies the built-in bullet again
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => markers(page)).toEqual([SQUARE, ARROW])
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('List Styles: Add makes a named list style that is saved and listed on reopen, Delete takes it out and its paragraphs turn plain', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-list-style-add-')))
  const docPath = join(dir, 'plain.docx')
  writeFileSync(docPath, await docxWith(BODY))
  const options = (dialog: ReturnType<Page['locator']>) =>
    dialog.locator('.bn-styles-list [role=option]')
  const first = await openDocs(docPath)
  try {
    const { launched, page } = first
    await selectThrough(page, 'gamma')
    let dialog = await openDialog(launched.app, page)
    await dialog.getByRole('tab', { name: 'List Styles' }).click()
    const add = dialog.getByRole('button', { name: 'Add' })
    const remove = dialog.getByRole('button', { name: 'Remove' })
    await expect(add).toBeEnabled()
    await expect(remove).toBeDisabled()
    // a built-in list style cannot be deleted
    await options(dialog).nth(2).click()
    await expect(remove).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Modify…' })).toBeEnabled()
    await add.click()
    await expect(dialog).toBeHidden()

    const define = page.locator('.modal.list-define-modal')
    await expect(define.locator('h2')).toHaveText('New Style')
    const ok = define.locator('button.primary')
    const name = define.locator('input[name="list-style-name"]')
    await expect(ok).toBeDisabled()
    // a name a list style already has is refused
    await name.fill('no list')
    await expect(ok).toBeDisabled()
    await name.fill('House Steps')
    await expect(ok).toBeEnabled()
    await define.getByLabel('Number format').fill('Step %1:')
    await page.screenshot({ path: test.info().outputPath('new-list-style.png') })
    await ok.click()
    await expect(define).toBeHidden()

    // back on the List Styles tab with the new style picked
    await expect(dialog).toBeVisible()
    await expect(options(dialog)).toHaveText([
      'No List',
      '1 / 1.1 / 1.1.1',
      '1 / a / i',
      'Article / Section',
      'House Steps',
    ])
    await expect(dialog.locator('.bn-styles-list [aria-selected=true]')).toHaveText('House Steps')
    await expect(remove).toBeEnabled()
    await expect(dialog.locator('.bn-styles-preview > span').first()).toContainText('Step 1:')
    await page.screenshot({ path: test.info().outputPath('list-style-added.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => markers(page)).toEqual(['Step 1:', 'Step 2:', 'Step 3:'])

    // a second one, never applied, is saved too
    dialog = await openDialog(launched.app, page)
    await dialog.getByRole('tab', { name: 'List Styles' }).click()
    await add.click()
    await name.fill('Spare')
    await ok.click()
    await expect(dialog.locator('.bn-styles-list [aria-selected=true]')).toHaveText('Spare')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await saveFromMenu(launched.app)
    await expect
      .poll(() => part(docPath, 'word/styles.xml'), { timeout: 30_000 })
      .toContain('w:styleId="HouseSteps"')
    const styles = await part(docPath, 'word/styles.xml')
    expect(styles).toContain(
      '<w:style w:type="numbering" w:customStyle="1" w:styleId="HouseSteps"><w:name w:val="House Steps"/>',
    )
    expect(styles).toContain(
      '<w:style w:type="numbering" w:customStyle="1" w:styleId="Spare"><w:name w:val="Spare"/>',
    )
    const numbering = await part(docPath, 'word/numbering.xml')
    expect(numbering).toMatch(
      /<w:styleLink w:val="HouseSteps"\/><w:lvl w:ilvl="0">.*?<w:lvlText w:val="Step %1:"\/>/,
    )
    expect(numbering.match(/<w:numStyleLink w:val="HouseSteps"\/>/g)).toHaveLength(1)
    expect(numbering.match(/<w:styleLink w:val="Spare"\/>/g)).toHaveLength(1)
    expect(numbering).not.toContain('<w:numStyleLink w:val="Spare"/>')
    expect((await part(docPath, 'word/document.xml')).match(/<w:numPr>/g)).toHaveLength(3)
    copyFileSync(docPath, test.info().outputPath('list-style-added.docx'))
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  const second = await openDocs(docPath)
  try {
    const { launched, page } = second
    await expect.poll(() => markers(page)).toEqual(['Step 1:', 'Step 2:', 'Step 3:'])
    await selectThrough(page, 'alpha')
    const dialog = await openDialog(launched.app, page)
    await dialog.getByRole('tab', { name: 'List Styles' }).click()
    await expect(options(dialog)).toHaveText([
      'No List',
      '1 / 1.1 / 1.1.1',
      '1 / a / i',
      'Article / Section',
      'House Steps',
      'Spare',
    ])
    // the caret's list follows the style, so the tab opens on it
    await expect(dialog.locator('.bn-styles-list [aria-selected=true]')).toHaveText('House Steps')
    page.once('dialog', (confirm) => void confirm.accept())
    await dialog.getByRole('button', { name: 'Remove' }).click()
    await expect(options(dialog)).toHaveText([
      'No List',
      '1 / 1.1 / 1.1.1',
      '1 / a / i',
      'Article / Section',
      'Spare',
    ])
    await expect(dialog.getByRole('button', { name: 'Remove' })).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('list-style-deleted.png') })
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect.poll(() => markers(page)).toEqual([])
    await expect(page.locator('.doc-page')).toContainText('alpha')

    await saveFromMenu(launched.app)
    await expect
      .poll(() => part(docPath, 'word/styles.xml'), { timeout: 30_000 })
      .not.toContain('HouseSteps')
    expect(await part(docPath, 'word/styles.xml')).toContain('w:styleId="Spare"')
    const numbering = await part(docPath, 'word/numbering.xml')
    expect(numbering).not.toContain('w:val="HouseSteps"')
    expect(numbering).toContain('<w:styleLink w:val="Spare"/>')
    expect(await part(docPath, 'word/document.xml')).not.toContain('<w:numPr>')
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
  }

  const third = await openDocs(docPath)
  try {
    const { launched, page } = third
    await expect.poll(() => markers(page)).toEqual([])
    const dialog = await openDialog(launched.app, page)
    await dialog.getByRole('tab', { name: 'List Styles' }).click()
    await expect(options(dialog)).toHaveText([
      'No List',
      '1 / 1.1 / 1.1.1',
      '1 / a / i',
      'Article / Section',
      'Spare',
    ])
  } finally {
    await closeAndSaveVideo(third.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
