/**
 * Dialog controls Word has that were missing: the Document Grid's lines per
 * page and line pitch, the Columns dialog's width and spacing per column with
 * Equal column width off and Start new column, and Bullets and Numbering's
 * Restart numbering / Continue previous list.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-dialog-live-controls-3'
const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/>'

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

async function caretIn(paragraph: Locator) {
  await paragraph.click()
  await paragraph.page().keyboard.press('End')
}

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

const documentXml = (docPath: string) => part(docPath, 'word/document.xml')

async function openDocs(docPath: string, text: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page').first()).toContainText(text, { timeout: 30_000 })
  return { launched, page }
}

/** px on an unzoomed page */
const unzoomed = (page: Page, text: string) =>
  page.evaluate((text) => {
    const paper = document.querySelector('.doc-page')!.getBoundingClientRect()
    const k = 816 / paper.width
    const el = Array.from(document.querySelectorAll('.doc-page p')).find(
      (p) => p.textContent === text,
    )!
    const r = el.getBoundingClientRect()
    return {
      left: Math.round((r.left - paper.left) * k),
      height: Math.round(r.height * k * 10) / 10,
    }
  }, text)

test('the Document Grid shows a line grid in lines per page and pitch, and sets it', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-3-')))
  const docPath = join(dir, 'grid.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('apple') +
        para('banana') +
        `<w:sectPr>${PAGE}<w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>`,
    ),
  )
  const grid = async () => /<w:docGrid\b[^>]*\/>/.exec(await documentXml(docPath))?.[0]
  const first = await openDocs(docPath, 'apple')
  try {
    const { launched, page } = first
    const dialog = page.locator('.doc-dialog')
    const perPage = dialog.locator('input[name="linesPerPage"]')
    const pitch = dialog.locator('input[name="linePitch"]')
    const openGrid = async () => {
      await caretIn(page.locator('.doc-page p', { hasText: 'apple' }).first())
      await clickMenu(launched.app, ['Format', 'Document…'])
      await dialog.getByRole('tab', { name: 'Document Grid' }).click()
    }
    // 15.6 pt lines on 9" of text; the paragraphs' 1.15 spacing scales the pitch
    await expect.poll(async () => (await unzoomed(page, 'apple')).height).toBe(23.9)
    await openGrid()
    await expect(perPage).toHaveValue('41')
    await expect(pitch).toHaveValue('15.6')
    for (const name of ['charsPerLine', 'charPitch', 'defaultCharPitch'])
      await expect(dialog.locator(`input[name="${name}"]`)).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Set Font…' })).toBeEnabled()
    await expect(dialog.getByRole('button', { name: 'Drawing Grid…' })).toHaveCount(0)
    await perPage.fill('20')
    await expect(pitch).toHaveValue('32.4')
    await page.screenshot({ path: test.info().outputPath('document-grid.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await unzoomed(page, 'apple')).height).toBe(49.7)
    await page.screenshot({ path: test.info().outputPath('document-grid-page.png') })
    await saveFromMenu(launched.app)
    await expect
      .poll(grid, { timeout: 30_000 })
      .toBe('<w:docGrid w:type="lines" w:linePitch="648"/>')

    // the pitch typed, in points
    await openGrid()
    await expect(perPage).toHaveValue('20')
    // typed key by key: a decimal on the way stays as typed
    await pitch.fill('')
    await pitch.pressSequentially('24.5')
    await expect(pitch).toHaveValue('24.5')
    await expect(perPage).toHaveValue('26')
    await pitch.press('Backspace')
    await pitch.press('Backspace')
    await expect(perPage).toHaveValue('27')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await unzoomed(page, 'apple')).height).toBe(36.8)
    await saveFromMenu(launched.app)
    await expect
      .poll(grid, { timeout: 30_000 })
      .toBe('<w:docGrid w:type="lines" w:linePitch="480"/>')
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  // the saved grid reads back; without a grid the numbers are shown and not offered
  const second = await openDocs(docPath, 'apple')
  try {
    const { launched, page } = second
    const dialog = page.locator('.doc-dialog')
    await expect.poll(async () => (await unzoomed(page, 'apple')).height).toBe(36.8)
    await caretIn(page.locator('.doc-page p', { hasText: 'apple' }).first())
    await clickMenu(launched.app, ['Format', 'Document…'])
    await dialog.getByRole('tab', { name: 'Document Grid' }).click()
    await expect(dialog.locator('input[name="linesPerPage"]')).toHaveValue('27')
    await expect(dialog.locator('input[name="linePitch"]')).toHaveValue('24')
    await dialog.locator('input[name="gridType"][value="default"]').check()
    await expect(dialog.locator('input[name="linesPerPage"]')).toBeDisabled()
    await expect(dialog.locator('input[name="linePitch"]')).toBeDisabled()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await saveFromMenu(launched.app)
    await expect.poll(grid, { timeout: 30_000 }).toBe('<w:docGrid w:linePitch="480"/>')
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the Columns dialog takes a width and spacing per column, and starts a new column', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-3-')))
  const docPath = join(dir, 'columns.docx')
  const lines = Array.from({ length: 30 }, (_, i) => `line ${String(i + 1).padStart(2, '0')}`)
  writeFileSync(docPath, await docxWith(lines.map(para).join('') + `<w:sectPr>${PAGE}</w:sectPr>`))
  const cols = async () =>
    /<w:cols\b[^>]*(?:\/>|>[\s\S]*?<\/w:cols>)/.exec(await documentXml(docPath))?.[0]
  const lefts = (page: Page) =>
    page.evaluate(() => {
      const paper = document.querySelector('.doc-page')!.getBoundingClientRect()
      const k = 816 / paper.width
      const all = Array.from(document.querySelectorAll('.doc-page p'))
        .filter((el) => el.textContent?.startsWith('line '))
        .map((el) => Math.round((el.getBoundingClientRect().left - paper.left) * k))
      return [...new Set(all)].sort((a, b) => a - b)
    })
  const first = await openDocs(docPath, 'line 01')
  try {
    const { launched, page } = first
    const dialog = page.locator('.columns-dialog')
    const equal = dialog.locator('input[name="equalWidth"]')
    const startNew = dialog.locator('input[name="startNewColumn"]')
    const width = (n: number) => dialog.getByRole('textbox', { name: `Width ${n}`, exact: true })
    const spacing = (n: number) =>
      dialog.getByRole('textbox', { name: `Spacing ${n}`, exact: true })
    const open = async (text: string) => {
      await caretIn(page.locator('.doc-page p', { hasText: text }).first())
      await clickMenu(launched.app, ['Format', 'Columns…'])
    }

    await open('line 01')
    await expect(width(1)).toBeDisabled()
    await expect(width(1)).toHaveValue('6.5"')
    await dialog.locator('input[name="columns"]').nth(1).check()
    // equal columns: the widths follow the one gap
    await expect(width(1)).toHaveValue('3"')
    await expect(width(1)).toBeDisabled()
    await expect(spacing(1)).toBeEnabled()
    await equal.uncheck()
    await expect(width(1)).toBeEnabled()
    await width(1).fill('2in')
    await width(1).blur()
    await expect(width(2)).toHaveValue('4"')
    await spacing(1).fill('0.25in')
    await spacing(1).blur()
    await expect(width(2)).toHaveValue('4.25"')
    await expect(startNew).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('columns-width-spacing.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // 1" margin, a 2" column and a quarter inch between
    await expect.poll(() => lefts(page)).toEqual([96, 312])
    await page.screenshot({ path: test.info().outputPath('columns-width-spacing-page.png') })
    await saveFromMenu(launched.app)
    await expect
      .poll(cols, { timeout: 30_000 })
      .toBe(
        '<w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="2880" w:space="360"/><w:col w:w="6120"/></w:cols>',
      )

    // read back; what follows the caret then starts in the next column
    await open('line 05')
    await expect(equal).not.toBeChecked()
    await expect(width(1)).toHaveValue('2"')
    await expect(spacing(1)).toHaveValue('0.25"')
    await expect(width(2)).toHaveValue('4.25"')
    await dialog.locator('select[name="applyTo"]').selectOption('forward')
    await startNew.check()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // line 06 opens the second column, beside line 01
    await expect
      .poll(async () => {
        const [top, last, next] = await Promise.all(
          ['line 01', 'line 05', 'line 06'].map((text) => unzoomed(page, text)),
        )
        return [last.left, next.left, next.top === top.top]
      })
      .toEqual([96, 312, true])
    await page.screenshot({ path: test.info().outputPath('columns-start-new-page.png') })
    await saveFromMenu(launched.app)
    await expect
      .poll(async () => (await documentXml(docPath)).match(/<w:type w:val="[^"]*"\/>/g), {
        timeout: 30_000,
      })
      .toEqual(['<w:type w:val="nextColumn"/>'])
    expect((await documentXml(docPath)).match(/<w:sectPr\b/g)).toHaveLength(2)
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  // a file that arrives with its own widths and gaps shows them and lays them out
  const second = await openDocs(docPath, 'line 01')
  try {
    const { launched, page } = second
    const dialog = page.locator('.columns-dialog')
    await expect.poll(() => lefts(page)).toEqual([96, 312])
    await expect
      .poll(
        async () => (await unzoomed(page, 'line 06')).top === (await unzoomed(page, 'line 01')).top,
      )
      .toBe(true)
    await caretIn(page.locator('.doc-page p', { hasText: 'line 01' }).first())
    await clickMenu(launched.app, ['Format', 'Columns…'])
    await expect(dialog.locator('input[name="equalWidth"]')).not.toBeChecked()
    await expect(dialog.getByRole('textbox', { name: 'Spacing 1', exact: true })).toHaveValue(
      '0.25"',
    )
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Bullets and Numbering restarts the numbering at the caret and continues the previous list', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-live-controls-3-')))
  const docPath = join(dir, 'list.docx')
  writeFileSync(docPath, await docxWith(['alpha', 'beta', 'gamma', 'delta'].map(para).join('')))
  const markers = (page: Page) =>
    page
      .locator('.doc-page .doc-li')
      .evaluateAll((items) => items.map((el) => el.getAttribute('data-marker')))
  const numIds = async () => [
    ...new Set((await documentXml(docPath)).match(/<w:numId w:val="\d+"\/>/g) ?? []),
  ]
  const first = await openDocs(docPath, 'alpha')
  try {
    const { launched, page } = first
    const dialog = page.locator('.bn-dialog')
    const restart = dialog.locator('input[name="listNumbering"][value="restart"]')
    const onward = dialog.locator('input[name="listNumbering"][value="continue"]')
    const openNumbered = async (text: string) => {
      await caretIn(page.locator('.doc-page').getByText(text, { exact: true }))
      await clickMenu(launched.app, ['Format', 'Bullets and Numbering…'])
      await dialog.getByRole('tab', { name: 'Numbered', exact: true }).click()
    }

    // outside a list there is nothing to restart; no card is ever off its default
    await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: { doc: { content: { size: number } } }
              chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } }
            }
          }
        }
      ).__aidocs
      editor
        .chain()
        .focus()
        .setTextSelection({ from: 1, to: editor.state.doc.content.size - 1 })
        .run()
    })
    await clickMenu(launched.app, ['Format', 'Bullets and Numbering…'])
    await expect(restart).toHaveCount(0)
    await expect(dialog.getByRole('button', { name: 'Reset' })).toBeDisabled()
    await dialog.getByRole('tab', { name: 'Numbered', exact: true }).click()
    await expect(restart).toBeChecked()
    await expect(restart).toBeDisabled()
    await expect(onward).toBeDisabled()
    await dialog.locator('.bn-grid [role=option]').nth(1).click()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => markers(page)).toEqual(['1.', '2.', '3.', '4.'])

    // the first item starts its list and has no earlier list to go on from
    await openNumbered('alpha')
    await expect(restart).toBeChecked()
    await expect(onward).toBeDisabled()
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await openNumbered('gamma')
    await expect(onward).toBeChecked()
    await restart.check()
    await page.screenshot({ path: test.info().outputPath('list-numbering.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => markers(page)).toEqual(['1.', '2.', '1.', '2.'])
    await page.screenshot({ path: test.info().outputPath('list-numbering-page.png') })
    await saveFromMenu(launched.app)
    await expect.poll(async () => (await numIds()).length, { timeout: 30_000 }).toBe(2)
    expect(await part(docPath, 'word/numbering.xml')).toContain('<w:startOverride w:val="1"/>')
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  // the restart reads back, and the list goes on from the one before it again
  const second = await openDocs(docPath, 'alpha')
  try {
    const { launched, page } = second
    const dialog = page.locator('.bn-dialog')
    await expect.poll(() => markers(page)).toEqual(['1.', '2.', '1.', '2.'])
    await caretIn(page.locator('.doc-page').getByText('gamma', { exact: true }))
    await clickMenu(launched.app, ['Format', 'Bullets and Numbering…'])
    await dialog.getByRole('tab', { name: 'Numbered', exact: true }).click()
    await expect(dialog.locator('input[name="listNumbering"][value="restart"]')).toBeChecked()
    await dialog.locator('input[name="listNumbering"][value="continue"]').check()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => markers(page)).toEqual(['1.', '2.', '3.', '4.'])
    await saveFromMenu(launched.app)
    await expect.poll(async () => (await numIds()).length, { timeout: 30_000 }).toBe(1)
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
