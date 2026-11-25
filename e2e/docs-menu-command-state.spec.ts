import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu, withLanguage } from './docs-objects-fixture'
import { selectParagraph } from './docs-word-acceptance-kit'

// The menu bar's check marks and greyed-out items follow the renderer's state through the command-state report.

const viewItem = (app: ElectronApplication, label: string) =>
  app.evaluate(({ Menu }, label) => {
    const view = Menu.getApplicationMenu()?.items.find((item) => item.label === 'View')
    const item = view?.submenu?.items.find((candidate) => candidate.label === label)
    return item ? { checked: item.checked, enabled: item.enabled } : null
  }, label)

/**
 * The saved main part parses as XML. One prefix that no element declares is a
 * parse error, and Word then refuses the whole file; the page's own reader is
 * lenient, so only a real parser shows it.
 */
const savedDocumentParses = async (page: Page, docPath: string) => {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
  return page.evaluate(
    (xml) =>
      new DOMParser().parseFromString(xml, 'application/xml').querySelector('parsererror')
        ?.textContent ?? 'ok',
    xml,
  )
}

const clickViewItem = (app: ElectronApplication, label: string) =>
  app.evaluate(({ Menu }, label) => {
    const view = Menu.getApplicationMenu()?.items.find((item) => item.label === 'View')
    view?.submenu?.items.find((candidate) => candidate.label === label)?.click()
  }, label)

test('View menu check marks follow the document window', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page, app } = launched
  try {
    await page.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const docs = await waitForPageWithUrl(app, '://docs/')
    const canvas = docs.locator('.doc-zoom')
    await expect.poll(() => viewItem(app, 'Dark Mode'), { timeout: 30_000 }).not.toBeNull()
    // the menu is up before the new document is: a view command does nothing until
    // there is one, and a check item ticks itself when clicked, hiding that
    await expect(docs.locator('.status-page')).toContainText(/\d/, { timeout: 30_000 })

    for (const label of ['Dark Mode', 'AI Sidebar', 'Ruler', 'Gridlines']) {
      const before = (await viewItem(app, label))!
      expect(before.enabled).toBe(true)
      await clickViewItem(app, label)
      await expect
        .poll(async () => (await viewItem(app, label))?.checked, { timeout: 15_000 })
        .toBe(!before.checked)
      await clickViewItem(app, label)
      await expect
        .poll(async () => (await viewItem(app, label))?.checked, { timeout: 15_000 })
        .toBe(before.checked)
    }

    // the layouts are one choice: picking Outline moves the tick off Print Layout, picking it again moves it back
    await expect
      .poll(async () => (await viewItem(app, 'Print Layout'))?.checked, { timeout: 15_000 })
      .toBe(true)
    await clickViewItem(app, 'Outline')
    // the document itself changes layout; the tick only follows it
    await expect(canvas).toHaveClass(/\bview-outline\b/)
    await expect
      .poll(
        async () => [
          (await viewItem(app, 'Outline'))?.checked,
          (await viewItem(app, 'Print Layout'))?.checked,
        ],
        { timeout: 15_000 },
      )
      .toEqual([true, false])
    await clickViewItem(app, 'Print Layout')
    await expect(canvas).toHaveClass(/\bview-print\b/)
    await expect
      .poll(async () => (await viewItem(app, 'Print Layout'))?.checked, { timeout: 15_000 })
      .toBe(true)

    // a check item flips its own tick when clicked, so the tick alone proves nothing: the grid must show
    await clickViewItem(app, 'Gridlines')
    await expect(canvas).toHaveClass(/\bshow-grid\b/)
    await clickViewItem(app, 'Gridlines')
    await expect(canvas).not.toHaveClass(/\bshow-grid\b/)

    // Draft is one of the layouts
    await clickViewItem(app, 'Draft')
    await expect(canvas).toHaveClass(/\bview-draft\b/)
    await expect
      .poll(
        async () => [
          (await viewItem(app, 'Draft'))?.checked,
          (await viewItem(app, 'Print Layout'))?.checked,
        ],
        { timeout: 15_000 },
      )
      .toEqual([true, false])
    await clickViewItem(app, 'Print Layout')
    await expect(canvas).toHaveClass(/\bview-print\b/)
    expect((await viewItem(app, 'Reveal Formatting'))?.enabled).toBe(true)
    // and one Word has that is not going to be built is not offered at all
    expect(await viewItem(app, 'Master Document')).toBeNull()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

const clickInsertItem = (app: ElectronApplication, path: string[]) =>
  app.evaluate(({ Menu }, path) => {
    let items = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Insert')?.submenu
      ?.items
    let item: Electron.MenuItem | undefined
    for (const label of path) {
      item = items?.find((candidate) => candidate.label === label)
      items = item?.submenu?.items
    }
    item?.click()
    return item ? { enabled: item.enabled } : null
  }, path)

test('Insert menu items in Word positions run their commands', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('.ProseMirror').first().click()
    await page.keyboard.type('first page')

    // a submenu item that carries its pick: the section break type. An unsaved
    // document keeps the break pending until its first save and says so.
    await clickInsertItem(app, ['Break', 'Section Break (Next Page)'])
    await expect(page.locator('.status-msg')).toContainText('takes effect after saving')

    // a dialog the ribbon and the menu now share
    await clickInsertItem(app, ['Bookmark…'])
    await expect(page.locator('.modal', { hasText: 'Bookmark' }).first()).toBeVisible()

    await page.keyboard.press('Escape')
    await clickInsertItem(app, ['Caption…'])
    await expect(page.locator('.modal', { hasText: 'Caption' }).first()).toBeVisible()

    // Word's chart types: a type we can write opens the chart dialog on it
    await page.keyboard.press('Escape')
    await clickInsertItem(app, ['Chart', 'Pie'])
    await expect(page.locator('.modal-chart .btn-primary', { hasText: 'Pie' })).toBeVisible()
    await page.keyboard.press('Escape')

    await clickInsertItem(app, ['Icons…'])
    await expect(page.locator('.icons-dialog')).toBeVisible()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

/** `lead`: paragraphs put ahead of the rest; `tail`: body XML after the last paragraph */
async function docxWithTable(lead: string[] = [], tail = ''): Promise<Buffer> {
  const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${lead.map((t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`).join('')}<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>${cell('A1')}${cell('B1')}</w:tr><w:tr>${cell('A2')}${cell('B2')}</w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>${tail}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** enabled state of Table menu items, by label path */
const tableItems = (app: ElectronApplication, paths: string[][]) =>
  app.evaluate(({ Menu }, paths) => {
    const table = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Table')
    return paths.map((path) => {
      let items = table?.submenu?.items ?? []
      let found: Electron.MenuItem | undefined
      for (const label of path) {
        found = items.find((candidate) => candidate.label === label)
        items = found?.submenu?.items ?? []
      }
      return found?.enabled ?? null
    })
  }, paths)

test('Table menu items are greyed out until the caret is in a table', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'table.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const paths = [
      ['Insert', 'Rows Below'],
      ['Select', 'Row'],
      ['Merge Cells'],
      ['Table Properties…'],
      ['View Gridlines'],
      ['Insert', 'Table…'],
    ]
    await page.locator('.doc-page p', { hasText: 'Before' }).first().click()
    await expect
      .poll(() => tableItems(app, paths), { timeout: 30_000 })
      .toEqual([false, false, false, false, true, true])

    await page.locator('.doc-page .doc-table').first().locator('td, th').first().click()
    await expect
      .poll(() => tableItems(app, paths), { timeout: 15_000 })
      .toEqual([true, true, false, true, true, true])

    await page.locator('.doc-page p', { hasText: 'After' }).first().click()
    await expect
      .poll(() => tableItems(app, paths), { timeout: 15_000 })
      .toEqual([false, false, false, false, true, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

/** enabled state of Edit menu items, by label */
const editItems = (app: ElectronApplication, labels: string[]) =>
  app.evaluate(({ Menu }, labels) => {
    const edit = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')
    // Repeat is worded by what it would do
    const names = (label: string) =>
      label === 'Repeat' ? ["Can't Repeat", 'Repeat', 'Redo'] : [label]
    return labels.map(
      (label) =>
        edit?.submenu?.items.find((candidate) => names(label).includes(candidate.label))?.enabled ??
        null,
    )
  }, labels)

test('Undo, Cut and Copy are greyed out until there is something to act on', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    const labels = ['Undo', 'Repeat', 'Cut', 'Copy', 'Paste', 'Select All']
    await page.locator('.ProseMirror').first().click()
    await expect
      .poll(() => editItems(app, labels), { timeout: 30_000 })
      .toEqual([false, false, false, false, true, true])

    await page.keyboard.type('hello world')
    await expect
      .poll(() => editItems(app, labels), { timeout: 15_000 })
      .toEqual([true, false, false, false, true, true])

    await page.keyboard.press('ControlOrMeta+A')
    await expect
      .poll(() => editItems(app, labels), { timeout: 15_000 })
      .toEqual([true, false, true, true, true, true])

    await page.keyboard.press('ControlOrMeta+Z')
    await expect.poll(async () => (await editItems(app, labels))[1], { timeout: 15_000 }).toBe(true)

    // in a text field Undo steps aside, so the key undoes the field, not the document
    await page.keyboard.type('kept text')
    await app.evaluate(({ Menu }) => {
      const edit = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')
      const find = edit?.submenu?.items.find((item) => item.label === 'Find')
      find?.submenu?.items.find((item) => item.label.startsWith('Find…'))?.click()
    })
    await expect(page.locator('input:focus')).toBeVisible()
    await expect
      .poll(async () => (await editItems(app, labels)).slice(0, 2), { timeout: 15_000 })
      .toEqual([false, false])
    await page.keyboard.type('abc')
    await page.keyboard.press('ControlOrMeta+Z')
    await expect(page.locator('.ProseMirror').first()).toContainText('kept text')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Tools > Track Changes > Highlight Changes opens the dialog Word has and applies its ticks', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('.ProseMirror').first().click()
    await page.keyboard.type('tracked text')
    const open = () =>
      app.evaluate(({ Menu }) => {
        const tools = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Tools')
        const track = tools?.submenu?.items.find((item) => item.label === 'Track Changes')
        track?.submenu?.items.find((item) => item.label === 'Highlight Changes…')?.click()
      })
    const dialog = page.locator('.hlc-dialog')
    const box = (label: string) => dialog.locator('label', { hasText: label }).locator('input')
    const track = page.locator('.status-track')

    await open()
    await expect(dialog).toBeVisible()
    await expect(box('Track changes while editing')).not.toBeChecked()
    await expect(box('Highlight changes on screen')).toBeChecked()
    // moves are tracked unless the document says otherwise; print follows the screen until it is set here
    await expect(box('Track moves while tracking changes')).toBeEnabled()
    await expect(box('Track moves while tracking changes')).toBeChecked()
    await expect(box('Highlight changes in printed document')).toBeChecked()
    await page.screenshot({ path: test.info().outputPath('highlight-changes.png') })

    await box('Track changes while editing').check()
    await box('Highlight changes on screen').uncheck()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect(track).toHaveClass(/\bon\b/)
    await expect(page.locator('.app')).toHaveClass(/rev-display-none/)

    // the dialog opens on what is now in effect, and Cancel changes nothing
    await open()
    await expect(box('Track changes while editing')).toBeChecked()
    await expect(box('Highlight changes on screen')).not.toBeChecked()
    await expect(box('Highlight changes in printed document')).not.toBeChecked()
    await box('Track changes while editing').uncheck()
    await dialog.locator('button', { hasText: 'Cancel' }).click()
    await expect(track).toHaveClass(/\bon\b/)

    // Options… hands over to the track changes options
    await open()
    await dialog.locator('button', { hasText: 'Options…' }).click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.tco-dialog')).toBeVisible()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Format > Document opens the three-tab dialog Word has and applies margins and header options', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('.ProseMirror').first().click()
    await page.keyboard.type('document settings')
    const open = () =>
      app.evaluate(({ Menu }) => {
        const format = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Format')
        format?.submenu?.items.find((item) => item.label === 'Document…')?.click()
      })
    const dialog = page.locator('.doc-dialog')
    const tab = (name: string) => dialog.locator('.dlg-tabs button', { hasText: name })
    const field = (label: string) => dialog.getByLabel(label, { exact: true })

    await open()
    await expect(dialog).toBeVisible()
    await expect(dialog.locator('.dlg-tabs button')).toHaveText([
      'Margins',
      'Layout',
      'Document Grid',
    ])
    const before = await field('Top').inputValue()
    await page.screenshot({ path: test.info().outputPath('document-margins.png') })
    await field('Top').fill('2"')
    await field('Top').press('Tab')

    await tab('Layout').click()
    await page.screenshot({ path: test.info().outputPath('document-layout.png') })
    await dialog.locator('label', { hasText: 'Different First Page' }).locator('input').check()
    // margins and both distances in one OK: each must survive the others
    await field('Header').fill('0.3"')
    await field('Header').press('Tab')
    await field('Footer').fill('0.8"')
    await field('Footer').press('Tab')
    await tab('Document Grid').click()
    await page.screenshot({ path: test.info().outputPath('document-grid.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    // the dialog opens on what is now in effect
    await open()
    const after = await field('Top').inputValue()
    expect(after).not.toBe(before)
    expect(after).toMatch(/^2(\.0+)?\s*("|in)|^5\.08\s*cm/)
    await tab('Layout').click()
    await expect(
      dialog.locator('label', { hasText: 'Different First Page' }).locator('input'),
    ).toBeChecked()
    await expect(field('Header')).toHaveValue(/^0\.3"?/)
    await expect(field('Footer')).toHaveValue(/^0\.8"?/)
    await dialog.locator('button', { hasText: 'Cancel' }).click()
    await expect(dialog).toBeHidden()

    // Enter in a field applies the dialog, as OK does
    await open()
    await field('Bottom').fill('1.5"')
    await field('Bottom').press('Enter')
    await expect(dialog).toBeHidden()
    await open()
    await expect(field('Bottom')).toHaveValue(/^1\.5"?/)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Change Case, Field, Date and Time, Drop Cap, Watermark and Paste Special open from the menus Word has them in', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    const body = page.locator('.ProseMirror').first()
    await body.click()
    await page.keyboard.type('quiet words')
    await page.keyboard.press('ControlOrMeta+a')
    const click = (menu: string, label: string) =>
      app.evaluate(
        ({ Menu }, [menu, label]) => {
          const top = Menu.getApplicationMenu()?.items.find((item) => item.label === menu)
          top?.submenu?.items.find((item) => item.label === label)?.click()
        },
        [menu, label],
      )

    await click('Format', 'Change Case…')
    const caseDialog = page.locator('.case-dialog')
    await expect(caseDialog.locator('label')).toHaveText([
      'Sentence case',
      'lowercase',
      'UPPERCASE',
      'Capitalize Each Word',
      'tOGGLE cASE',
      'Half-width',
      'Full-width',
    ])
    await page.screenshot({ path: test.info().outputPath('change-case.png') })
    await caseDialog.locator('label', { hasText: 'UPPERCASE' }).locator('input').check()
    await caseDialog.locator('button.primary').click()
    await expect(caseDialog).toBeHidden()
    await expect(body).toContainText('QUIET WORDS')

    await click('Insert', 'Field…')
    const fieldDialog = page.locator('.modal', { hasText: 'Field' }).first()
    await expect(fieldDialog).toBeVisible()
    await fieldDialog.getByLabel('Field code').fill('FILENAME')
    await fieldDialog.locator('.btn-primary').click()
    await expect(fieldDialog).toBeHidden()
    await expect(page.locator('.status-msg')).toContainText('FILENAME')

    // Date and Time: Word's formats, inserted as text unless it should update
    await body.click()
    await page.keyboard.press('ControlOrMeta+End')
    await click('Insert', 'Date and Time…')
    const dateDialog = page.locator('.date-dialog')
    await expect(dateDialog.locator('[role=option]')).toHaveCount(17)
    await page.screenshot({ path: test.info().outputPath('date-time.png') })
    const iso = dateDialog.locator('[role=option]').nth(4)
    const isoText = (await iso.textContent()) ?? ''
    expect(isoText).toMatch(/^\d{4}-\d{2}-\d{2}$/)
    await iso.click()
    await dateDialog.locator('button.primary').click()
    await expect(dateDialog).toBeHidden()
    await expect(body).toContainText(isoText)

    // Drop Cap: the dialog opens on the paragraph and applies position and lines
    await click('Format', 'Drop Cap…')
    const dropDialog = page.locator('.dropcap-dialog')
    const position = (label: string) =>
      dropDialog.locator('label', { hasText: label }).locator('input')
    await expect(position('None')).toBeChecked()
    await position('Dropped').check()
    await dropDialog.getByLabel('Lines to drop').fill('2')
    await page.screenshot({ path: test.info().outputPath('drop-cap.png') })
    await dropDialog.locator('button.primary').click()
    await expect(dropDialog).toBeHidden()
    // the letter is drawn at the lines asked for, not a fixed three
    const cap = page.locator('.has-drop-cap').first()
    await expect(cap).toHaveAttribute('style', /--drop-cap-lines:\s*2/)
    const capSize = await cap.evaluate((el) =>
      parseFloat(getComputedStyle(el, '::first-letter').fontSize),
    )
    const bodySize = await cap.evaluate((el) => parseFloat(getComputedStyle(el).fontSize))
    expect(capSize / bodySize).toBeCloseTo(2, 1)
    await click('Format', 'Drop Cap…')
    await expect(position('Dropped')).toBeChecked()
    await expect(dropDialog.getByLabel('Lines to drop')).toHaveValue('2')
    await dropDialog.locator('button', { hasText: 'Cancel' }).click()

    // Watermark: the dialog the Design tab opens, now also from the Insert menu
    await click('Insert', 'Watermark…')
    const watermarkDialog = page.locator('.watermark-dialog')
    await expect(watermarkDialog).toBeVisible()
    await expect(watermarkDialog.locator('h2')).toHaveText('Insert Watermark')
    await watermarkDialog.locator('.wm-mode', { hasText: 'Text' }).locator('input').check()
    await watermarkDialog.locator('.wm-row input[type=text]').first().fill('DRAFT')
    await page.screenshot({ path: test.info().outputPath('watermark.png') })
    await watermarkDialog.locator('.btn-primary').click()
    await expect(watermarkDialog).toBeHidden()
    // the Design tab's Watermark button lights up once the document has one
    await page.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
    await expect(
      page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Watermark' }),
    ).toHaveClass(/\bactive\b/)

    await body.click()
    await click('Edit', 'Paste Special…')
    await expect(page.locator('.modal', { hasText: 'Paste Special' }).first()).toBeVisible()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Find Next, Find Previous and Columns work from the menus Word has them in', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    const body = page.locator('.ProseMirror').first()
    await body.click()
    await page.keyboard.type('echo one echo two echo three')
    const click = (menu: string, path: string[]) =>
      app.evaluate(
        ({ Menu }, [menu, ...path]) => {
          let items = Menu.getApplicationMenu()?.items.find((item) => item.label === menu)?.submenu
            ?.items
          let item: Electron.MenuItem | undefined
          for (const label of path) {
            item = items?.find((candidate) => candidate.label === label)
            items = item?.submenu?.items
          }
          item?.click()
        },
        [menu, ...path],
      )

    // with no search yet, Find Next opens the panel on the find field
    await click('Edit', ['Find', 'Find Next'])
    const count = page.locator('.find-count')
    await page.locator('.find-input').first().fill('echo')
    await expect(count).toContainText('3')
    const before = await count.textContent()
    await click('Edit', ['Find', 'Find Next'])
    await expect(count).not.toHaveText(before ?? '')
    const afterNext = await count.textContent()
    await click('Edit', ['Find', 'Find Previous'])
    await expect(count).toHaveText(before ?? '')
    expect(afterNext).not.toBe(before)

    // Columns: Word's dialog, on the section as it is
    await click('Format', ['Columns…'])
    const dialog = page.locator('.columns-dialog')
    const preset = (label: string) => dialog.locator('label', { hasText: label }).locator('input')
    await expect(preset('One')).toBeChecked()
    await preset('Two').check()
    await expect(dialog.getByLabel('Number of columns')).toHaveValue('2')
    await page.screenshot({ path: test.info().outputPath('columns.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await click('Format', ['Columns…'])
    await expect(preset('Two')).toBeChecked()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Thumbnail, Advanced Symbol and Next Footnote work from the places Word has them in', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    const body = page.locator('main.editor-scroll .ProseMirror').first()
    await body.click()
    await page.keyboard.type('first second')
    const click = (menu: string, path: string[]) =>
      app.evaluate(
        ({ Menu }, [menu, ...path]) => {
          let items = Menu.getApplicationMenu()?.items.find((item) => item.label === menu)?.submenu
            ?.items
          let item: Electron.MenuItem | undefined
          for (const label of path) {
            item = items?.find((candidate) => candidate.label === label)
            items = item?.submenu?.items
          }
          item?.click()
        },
        [menu, ...path],
      )

    // Thumbnail: the navigation pane on its pages, and off again
    const pane = page.locator('.nav-pane')
    await click('View', ['Sidebar', 'Thumbnail'])
    await expect(pane.locator('.nav-tab.on')).toHaveText('Pages')
    await pane.locator('.nav-tab', { hasText: 'Headings' }).click()
    await click('View', ['Sidebar', 'Thumbnail'])
    await expect(pane.locator('.nav-tab.on')).toHaveText('Pages')
    await click('View', ['Sidebar', 'Thumbnail'])
    await expect(pane).toBeHidden()
    // a Thumbnail pick is spent once: the pane reopens on the tab it was left on
    await click('View', ['Sidebar', 'Thumbnail'])
    await pane.locator('.nav-tab', { hasText: 'Headings' }).click()
    await click('View', ['Sidebar', 'Navigation'])
    await expect(pane).toBeHidden()
    await click('View', ['Sidebar', 'Navigation'])
    await expect(pane.locator('.nav-tab.on')).toHaveText('Headings')
    await click('View', ['Sidebar', 'Navigation'])
    await expect(pane).toBeHidden()

    // Advanced Symbol: Insert leaves the dialog open for the next character
    await body.click()
    await click('Insert', ['Advanced Symbol…'])
    const dialog = page.locator('.symbol-dialog')
    await dialog.getByLabel('Character code').fill('00A9')
    await page.screenshot({ path: test.info().outputPath('symbol.png') })
    await dialog.locator('button.primary').click()
    await expect(body).toContainText('©')
    await dialog.locator('.symbol-cell', { hasText: '±' }).dblclick()
    await expect(body).toContainText('©±')
    await expect(dialog).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()

    // Next Footnote: from the start of the text to the reference
    await click('Insert', ['Footnote…'])
    await page.locator('.footnote-dialog button.primary').click()
    await page.locator('.modal .prompt-textarea').fill('a note')
    await page.keyboard.press('ControlOrMeta+Enter')
    const ref = body.locator('sup.doc-note-ref')
    await expect(ref).toHaveCount(1)
    await body.click()
    // back in front of the reference, by a key every platform moves the caret with
    for (let i = 0; i < 20; i++) await page.keyboard.press('ArrowLeft')
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Next Footnote' })
      .click()
    await expect(ref).toHaveClass(/ProseMirror-selectednode/)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Envelopes and Labels open the mail merge panel from the Tools menu and the Mailings tab', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('main.editor-scroll .ProseMirror').first().click()
    const tools = (label: string) =>
      app.evaluate(({ Menu }, label) => {
        const menu = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Tools')
        menu?.submenu?.items.find((candidate) => candidate.label === label)?.click()
      }, label)
    const panel = page.locator('.modal-mailmerge')
    // the panel takes the focus after it is drawn; Escape pressed before that goes to the page
    const closePanel = async () => {
      await expect(panel.locator(':focus')).toHaveCount(1)
      await page.keyboard.press('Escape')
      await expect(panel).toBeHidden()
    }

    for (const label of ['Envelopes…', 'Labels…']) {
      await tools(label)
      await closePanel()
    }

    // the Mailings tab runs the same commands, and the panel outlives a change of tab
    await page.locator('.ribbon-tab', { hasText: /^Mailings$/ }).click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Labels$/ })
      .click()
    await closePanel()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Start Mail Merge' })
      .click()
    await expect(panel).toBeVisible()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Page Numbers opens its dialog from the Insert menu, and View > Markup hides and shows the markup', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('main.editor-scroll .ProseMirror').first().click()
    await page.keyboard.type('body text')
    const click = (menu: string, label: string) =>
      app.evaluate(
        ({ Menu }, [menu, label]) => {
          const top = Menu.getApplicationMenu()?.items.find((item) => item.label === menu)
          top?.submenu?.items.find((candidate) => candidate.label === label)?.click()
        },
        [menu, label],
      )

    // Page Numbers: Word's defaults, then the top of the page, centred
    await click('Insert', 'Page Numbers…')
    const dialog = page.locator('.page-numbers-dialog')
    const position = dialog.getByLabel('Position')
    const alignment = dialog.getByLabel('Alignment')
    await expect(position).toHaveValue('footer')
    await expect(alignment).toHaveValue('right')
    await position.selectOption('header')
    await alignment.selectOption('center')
    await page.screenshot({ path: test.info().outputPath('page-numbers.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.page-hf-header').first()).toHaveText('1')
    await expect(page.locator('.page-hf-footer', { hasText: '1' })).toHaveCount(0)

    // Markup: the Review tab's display button lights up while the markup is hidden
    await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
    const display = page.getByRole('button', { name: 'Display for Review', exact: true })
    await expect(display).not.toHaveClass(/active/)
    await click('View', 'Markup')
    await expect(display).toHaveClass(/active/)
    await click('View', 'Markup')
    await expect(display).not.toHaveClass(/active/)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Format > Tabs sets, lists and clears the tab stops of the paragraph', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('main.editor-scroll .ProseMirror').first().click()
    await page.keyboard.type('name')
    const openTabs = () =>
      app.evaluate(({ Menu }) => {
        const format = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Format')
        format?.submenu?.items.find((candidate) => candidate.label === 'Tabs…')?.click()
      })
    const dialog = page.locator('.tabs-dialog')
    const listed = dialog.locator('.tabs-dialog-list button')
    const position = dialog.getByLabel('Tab stop position')

    await openTabs()
    await expect(listed).toHaveCount(0)
    await position.fill('1 in')
    await dialog.getByLabel('Right', { exact: true }).check()
    await dialog.getByRole('button', { name: 'Set', exact: true }).click()
    await expect(listed).toHaveCount(1)
    // a second stop left typed: OK sets it too
    await position.fill('3 in')
    await dialog.getByLabel('.......').check()
    await page.screenshot({ path: test.info().outputPath('tabs.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    await openTabs()
    await expect(listed).toHaveCount(2)
    await listed.first().click()
    await expect(dialog.getByLabel('Right', { exact: true })).toBeChecked()
    await listed.last().click()
    await expect(dialog.getByLabel('.......')).toBeChecked()
    await dialog.getByRole('button', { name: 'Clear All' }).click()
    await dialog.locator('button.primary').click()
    await openTabs()
    await expect(listed).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Format > Bullets and Numbering applies a library preset and None takes the list off', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    const body = page.locator('main.editor-scroll .ProseMirror').first()
    await body.click()
    await page.keyboard.type('item')
    const open = () =>
      app.evaluate(({ Menu }) => {
        const format = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Format')
        format?.submenu?.items
          .find((candidate) => candidate.label === 'Bullets and Numbering…')
          ?.click()
      })
    const dialog = page.locator('.bn-dialog')
    const cards = dialog.locator('.bn-grid .list-gallery-card')

    await open()
    await dialog.getByRole('tab', { name: 'Numbered', exact: true }).click()
    await cards.nth(2).click()
    await page.screenshot({ path: test.info().outputPath('bullets-numbering.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect(body.locator('.doc-li-ordered')).toHaveCount(1)

    await open()
    await cards.nth(1).click()
    await dialog.locator('button.primary').click()
    await expect(body.locator('.doc-li')).toHaveCount(1)
    await expect(body.locator('.doc-li-ordered')).toHaveCount(0)

    // OK with nothing picked leaves the paragraph as it is; None takes the list off
    await open()
    await dialog.locator('button.primary').click()
    await expect(body.locator('.doc-li')).toHaveCount(1)
    await open()
    await cards.first().click()
    await dialog.locator('button.primary').click()
    await expect(body.locator('.doc-li')).toHaveCount(0)

    // Customize hands over to the dialog that defines a new bullet
    await open()
    await dialog.getByRole('button', { name: 'Customize…' }).click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.modal', { hasText: 'Define New Bullet' })).toBeVisible()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Distributed, Line Numbers, Clear Contents and Advanced Find work from the places Word has them in', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    const body = page.locator('main.editor-scroll .ProseMirror').first()
    await body.click()
    await page.keyboard.type('a line of text')
    const click = (menu: string, path: string[]) =>
      app.evaluate(
        ({ Menu }, [menu, ...path]) => {
          let items = Menu.getApplicationMenu()?.items.find((item) => item.label === menu)?.submenu
            ?.items
          let item: Electron.MenuItem | undefined
          for (const label of path) {
            item = items?.find((candidate) => candidate.label === label)
            items = item?.submenu?.items
          }
          item?.click()
        },
        [menu, ...path],
      )

    // Distributed: the Home tab's button lights up on the paragraph it set
    const distributed = page.locator('[data-ribbon-body] button.rb-icon[aria-label="Distributed"]')
    await expect(distributed).not.toHaveClass(/active/)
    await distributed.click()
    await expect(distributed).toHaveClass(/active/)

    // Line Numbers: on from the Layout tab, and off again
    await page.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    const lineNumbers = page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: 'Line Numbers',
    })
    const numbers = page.locator('.page-line-numbers')
    await expect(numbers).toHaveCount(0)
    await lineNumbers.click()
    await page.screenshot({ path: test.info().outputPath('line-numbers.png') })
    await page.locator('.layout-menu button', { hasText: /^Continuous$/ }).click()
    await expect(numbers.first()).toContainText('1')
    await expect(lineNumbers).toHaveClass(/active/)
    await lineNumbers.click()
    await page.locator('.layout-menu button', { hasText: /^None$/ }).click()
    await expect(numbers).toHaveCount(0)

    // Advanced Find and Replace: the Find panel
    await click('Edit', ['Find', 'Advanced Find and Replace…'])
    await expect(page.locator('.find-panel')).toBeVisible()

    // Clear > Contents: the selection goes
    await body.click()
    await page.keyboard.press('ControlOrMeta+A')
    await click('Edit', ['Clear', 'Contents'])
    await expect(body).not.toContainText('a line of text')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('the Tools items, Add Text and Finish & Merge that held a place now run', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    const body = page.locator('main.editor-scroll .ProseMirror').first()
    await body.click()
    await page.keyboard.type('a heading to be')
    const tools = (path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Tools')?.submenu
          ?.items
        let item: Electron.MenuItem | undefined
        for (const label of path) {
          item = items?.find((candidate) => candidate.label === label)
          items = item?.submenu?.items
        }
        item?.click()
      }, path)
    const status = page.locator('.status-msg')

    // with nothing selected and the caret in no word, Thesaurus and Translate Selection say what they need
    await page.keyboard.type('  ')
    await tools(['Thesaurus…'])
    await expect(status).toContainText('Select a word or a short phrase first')
    await tools(['Translate', 'Translate Selection'])
    await expect(status).toContainText('Select the text to translate first')
    await page.keyboard.press('Backspace')
    await page.keyboard.press('Backspace')

    // the translation language is kept
    const language = page.locator('.translation-language-dialog')
    await tools(['Translate', 'Set Document Translation Language…'])
    await language.locator('select').selectOption('fr')
    await page.screenshot({ path: test.info().outputPath('translation-language.png') })
    await language.locator('button.primary').click()
    await expect(language).toBeHidden()
    await tools(['Translate', 'Set Document Translation Language…'])
    await expect(language.locator('select')).toHaveValue('fr')
    await language.getByRole('button', { name: 'Cancel' }).click()

    // Customize Keyboard opens its dialog
    const shortcuts = page.locator('.modal-customize-keyboard')
    await tools(['Customize Keyboard…'])
    await expect(shortcuts.locator(':focus')).toHaveCount(1)
    await page.keyboard.press('Escape')
    await expect(shortcuts).toBeHidden()

    // Add Text: a level makes the paragraph a heading, Do Not Show makes it body text again
    await body.click()
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    const addText = page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: 'Add Text',
    })
    await addText.click()
    await page.locator('.layout-menu button', { hasText: /^Level 2$/ }).click()
    await expect(body.locator('h2')).toHaveText('a heading to be')
    await addText.click()
    await page
      .locator('.layout-menu button', { hasText: 'Do Not Show in Table of Contents' })
      .click()
    await expect(body.locator('h2')).toHaveCount(0)

    // Finish & Merge: greyed until a recipient list is attached (as Word has it), then the merge panel
    await page.locator('.ribbon-tab', { hasText: /^Mailings$/ }).click()
    const mailings = (name: string) =>
      page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })
    await expect(mailings('Finish & Merge')).toBeDisabled()
    const listPath = join(mkdtempSync(join(tmpdir(), 'docs-finish-merge-')), 'people.csv')
    writeFileSync(listPath, 'Name\nAda\n')
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, listPath)
    await mailings('Select Recipients').click()
    await mailings('Finish & Merge').click()
    await expect(page.locator('.modal-mailmerge')).toBeVisible()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('Ribbon, Footnotes, Reset Ignored Words, Properties, Always Open Read-Only and the recipient dialogs run', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-batch11-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWithTable())
  const listPath = join(dir, 'people.csv')
  writeFileSync(listPath, 'Name,City\nAda,London\nGrace,New York\nAlan,London\n')
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('.doc-page p', { hasText: 'Before' }).first().click()
    const find = (path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items
        let item: Electron.MenuItem | undefined
        for (const label of path) {
          item = items?.find((candidate) => candidate.label === label)
          items = item?.submenu?.items
        }
        return item ? { checked: item.checked, enabled: item.enabled } : null
      }, path)
    const click = (path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items
        let item: Electron.MenuItem | undefined
        for (const label of path) {
          item = items?.find((candidate) => candidate.label === label)
          items = item?.submenu?.items
        }
        item?.click()
      }, path)
    const status = page.locator('.status-msg')

    // View > Ribbon folds the band and its check mark follows
    const ribbon = page.locator('.ribbon-collapsible')
    await expect
      .poll(() => find(['View', 'Ribbon']), { timeout: 30_000 })
      .toMatchObject({
        checked: true,
        enabled: true,
      })
    await click(['View', 'Ribbon'])
    await expect(ribbon).toHaveClass(/ribbon-collapsed/)
    await expect.poll(() => find(['View', 'Ribbon'])).toMatchObject({ checked: false })
    await click(['View', 'Ribbon'])
    await expect(ribbon).not.toHaveClass(/ribbon-collapsed/)
    await expect.poll(() => find(['View', 'Ribbon'])).toMatchObject({ checked: true })

    await click(['Tools', 'Spelling and Grammar', 'Reset Ignored Words and Grammar'])
    await expect(status).toContainText('Ignored words will be checked again')
    // with no note in the document, Footnotes is greyed out as in Word
    await expect.poll(() => find(['View', 'Footnotes'])).toMatchObject({ enabled: false })

    // File > Properties: the fields wait for the save
    const properties = page.locator('.properties-dialog')
    await click(['File', 'Properties…'])
    await properties.locator('input[name="core.title"]').fill('Quarterly plan')
    await properties.locator('input[name="app.company"]').fill('Analytical Engines')
    await page.screenshot({ path: test.info().outputPath('properties.png') })
    await properties.locator('button.primary').click()
    await expect(properties).toBeHidden()

    // File > Always Open Read-Only: ticked in the menu, pressed on the Review tab
    await click(['File', 'Always Open Read-Only'])
    await expect
      .poll(() => find(['File', 'Always Open Read-Only']))
      .toMatchObject({ checked: true })
    await page.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
    const readOnly = page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: 'Always Open Read-Only',
    })
    await expect(readOnly).toHaveClass(/active/)
    await readOnly.scrollIntoViewIfNeeded()
    await page.screenshot({ path: test.info().outputPath('always-read-only.png') })

    // both reach the file
    await page.locator('.doc-page p', { hasText: 'Before' }).first().click()
    await saveFromMenu(app)
    const part = async (name: string) =>
      (await (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string')) ?? ''
    await expect
      .poll(() => part('docProps/core.xml'), { timeout: 30_000 })
      .toContain('<dc:title>Quarterly plan</dc:title>')
    expect(await part('docProps/app.xml')).toContain('<Company>Analytical Engines</Company>')
    expect(await part('word/settings.xml')).toMatch(
      /<w:writeProtection[^>]*w:recommended="(1|true)"/,
    )
    await click(['File', 'Properties…'])
    await expect(properties.locator('input[name="core.title"]')).toHaveValue('Quarterly plan')
    await properties.getByRole('button', { name: 'Cancel' }).click()

    // Mailings: with no list yet the recipient commands wait, as in Word; Select Recipients attaches one
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, listPath)
    await page.locator('.ribbon-tab', { hasText: /^Mailings$/ }).click()
    const big = (name: string) =>
      page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })
    const record = page.locator('[data-ribbon-body] [data-testid="mailings-record-status"]')
    for (const name of [
      'Edit Recipient List',
      'Filter Recipients',
      'Find Recipient',
      'Rules',
      'Preview Results',
      'Finish & Merge',
    ])
      await expect(big(name), name).toBeDisabled()
    await big('Select Recipients').click()
    await expect(record).toHaveText('Record 1 of 3')
    await expect(big('Edit Recipient List')).toBeEnabled()

    // unticking a row leaves that recipient out
    const recipients = page.locator('.recipients-dialog')
    await big('Edit Recipient List').click()
    await recipients.locator('tbody tr', { hasText: 'Grace' }).locator('input').uncheck()
    await expect(recipients.locator('.recipients-count')).toHaveText('2 of 3 recipients selected')
    await page.screenshot({ path: test.info().outputPath('recipients.png') })
    await recipients.locator('button.primary').click()
    await expect(record).toHaveText('Record 1 of 2')

    // a filter narrows the rest; Clear All lifts it
    const filter = page.locator('.filter-recipients-dialog')
    await big('Filter Recipients').click()
    // OK on the untouched form compares against nothing: no recipient is dropped
    await filter.locator('button.primary').click()
    await expect(record).toHaveText('Record 1 of 2')
    await big('Filter Recipients').click()
    await filter.locator('select[name="field"]').selectOption('Name')
    await filter.locator('select[name="op"]').selectOption('contains')
    await filter.locator('input[name="value"]').fill('AL')
    await page.screenshot({ path: test.info().outputPath('filter-recipients.png') })
    await filter.locator('button.primary').click()
    await expect(record).toHaveText('Record 1 of 1')
    await big('Filter Recipients').click()
    await filter.getByRole('button', { name: 'Clear All' }).click()
    await expect(record).toHaveText('Record 1 of 2')

    // Find Recipient moves the previewed record and says when nothing matches
    const finder = page.locator('.find-recipient-dialog')
    await big('Find Recipient').click()
    await finder.locator('input[name="text"]').fill('alan')
    await finder.locator('button.primary').click()
    await expect(record).toHaveText('Record 2 of 2')
    await finder.locator('input[name="text"]').fill('nobody')
    await finder.locator('button.primary').click()
    await expect(finder.locator('.modal-error')).toContainText('No recipient contains')
    await page.screenshot({ path: test.info().outputPath('find-recipient.png') })
    await finder.getByRole('button', { name: 'Close' }).click()
    await expect(finder).toBeHidden()

    // Preview Results: a merge field shows the current recipient, follows the record, and goes back
    const greeting = page.locator('.doc-page p').first()
    await greeting.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' for ')
    await big('Insert Merge Field').click()
    await page.locator('.rb-field-item', { hasText: /^Name$/ }).click()
    const fieldName = greeting.locator('[data-instr-field]')
    await expect(fieldName).toBeVisible()
    const preview = page.locator('.doc-page .doc-merge-preview')
    await expect(preview).toHaveCount(0)
    const previewButton = big('Preview Results')
    await previewButton.click()
    await expect(previewButton).toHaveClass(/active/)
    await expect(preview).toHaveText('Alan')
    await expect(fieldName).toBeHidden()
    await greeting.scrollIntoViewIfNeeded()
    await page.screenshot({ path: test.info().outputPath('preview-results.png') })
    await page.locator('[data-ribbon-body] button[aria-label="First"]').click()
    await expect(preview).toHaveText('Ada')
    await previewButton.click()
    await expect(previewButton).not.toHaveClass(/active/)
    await expect(preview).toHaveCount(0)
    await expect(fieldName).toBeVisible()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a dialog opened while Reset Ignored Words rechecks the page keeps its focus and its tab', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-respell-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p', { hasText: 'Before' }).first()
    await first.click()
    const text = await first.innerText()
    const enabled = (path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items
        let item: Electron.MenuItem | undefined
        for (const label of path) {
          item = items?.find((candidate) => candidate.label === label)
          items = item?.submenu?.items
        }
        return !!item?.enabled
      }, path)
    await expect.poll(() => enabled(['File', 'Properties…']), { timeout: 30_000 }).toBe(true)
    // both commands reach the page before it has drawn a frame
    await app.evaluate(({ Menu }) => {
      const top = Menu.getApplicationMenu()?.items
      const under = (label: string) => top?.find((item) => item.label === label)?.submenu?.items
      under('Tools')
        ?.find((item) => item.label === 'Spelling and Grammar')
        ?.submenu?.items.find((item) => item.label.startsWith('Reset Ignored Words'))
        ?.click()
      under('File')
        ?.find((item) => item.label === 'Properties…')
        ?.click()
    })
    const properties = page.locator('.properties-dialog')
    await expect(properties).toBeVisible()
    await expect(properties.locator('.bs-tabs button.active')).toHaveText('Summary')
    expect(await page.evaluate(() => !!document.activeElement?.closest('.properties-dialog'))).toBe(
      true,
    )
    // the keyboard is the dialog's: typing fills its field, not the page behind it
    await properties.locator('input[name="core.title"]').focus()
    await page.keyboard.type('Plan')
    await expect(properties.locator('input[name="core.title"]')).toHaveValue('Plan')
    await properties.getByRole('button', { name: 'Cancel' }).click()
    await expect(properties).toBeHidden()
    // the recheck leaves the text as it was
    await expect(first).toHaveText(text)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a dialog opened from the menu bar closes on an Escape that comes as it appears', async () => {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-escape-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWithTable(['cherry', 'apple']))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    await page.locator('.doc-page p', { hasText: 'Before' }).first().click()
    // Escape in the same turn the dialog is put on the page, to whatever has the focus then
    await page.evaluate(() => {
      new MutationObserver((_, observer) => {
        if (!document.querySelector('.sort-dialog')) return
        observer.disconnect()
        document.body.dataset.escapeSent = '1'
        document.activeElement?.dispatchEvent(
          new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
        )
      }).observe(document.body, { childList: true, subtree: true })
    })
    await expect
      .poll(() =>
        app.evaluate(({ Menu }) => {
          const table = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Table')
          const sort = table?.submenu?.items.find((item) => item.label === 'Sort…')
          if (!sort?.enabled) return false
          sort.click()
          return true
        }),
      )
      .toBe(true)
    await expect(page.locator('body[data-escape-sent]')).toHaveCount(1)
    await expect(page.locator('.sort-dialog')).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-dialog-escape')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Focus, Number, Text Direction, Sort, Rename and Move run from the places Word has them in', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-batch12-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWithTable(['cherry', 'apple']))
  const archive = join(dir, 'archive')
  mkdirSync(archive)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const before = page.locator('.doc-page p', { hasText: 'Before' }).first()
    await before.click()
    const find = (path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items
        let item: Electron.MenuItem | undefined
        for (const label of path) {
          item = items?.find((candidate) => candidate.label === label)
          items = item?.submenu?.items
        }
        return item ? { checked: item.checked, enabled: item.enabled } : null
      }, path)
    const click = (path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items
        let item: Electron.MenuItem | undefined
        for (const label of path) {
          item = items?.find((candidate) => candidate.label === label)
          items = item?.submenu?.items
        }
        item?.click()
      }, path)
    const status = page.locator('.status-msg')
    const ok = (dialog: ReturnType<typeof page.locator>) => dialog.locator('button.primary').click()

    // Sort: a table in the range is left alone; the paragraphs the selection touches are sorted
    const sortButton = page.locator('[data-ribbon-body] button[aria-label="Sort"]')
    const sort = page.locator('.sort-dialog')
    const lines = page.locator('main.editor-scroll .ProseMirror').first().locator('p')
    await sortButton.click()
    await ok(sort)
    await expect(status).toContainText('Select two or more paragraphs to sort')
    // not a click and Shift+ArrowDown: the editor redraws its own selection 20 ms after the
    // click gives it the focus, and a selection the key made before that is lost
    await selectParagraph(page, 'cherry', 'apple')
    await sortButton.click()
    await page.screenshot({ path: test.info().outputPath('sort.png') })
    await ok(sort)
    await expect
      .poll(async () => (await lines.allTextContents()).slice(0, 3))
      .toEqual(['apple', 'cherry', 'Before'])

    // View > Focus leaves the page alone on screen, still editable; Esc brings the window back
    const shell = page.locator('.app')
    await expect
      .poll(() => find(['View', 'Focus']), { timeout: 30_000 })
      .toMatchObject({
        checked: false,
        enabled: true,
      })
    await click(['View', 'Focus'])
    await expect(shell).toHaveClass(/focus-mode/)
    await expect(page.locator('.ribbon')).toBeHidden()
    await expect.poll(() => find(['View', 'Focus'])).toMatchObject({ checked: true })
    await page.screenshot({ path: test.info().outputPath('focus.png') })
    await before.click()
    await page.keyboard.type('!')
    await expect(before).toContainText('!')
    await page.keyboard.press('Escape')
    await expect(shell).not.toHaveClass(/focus-mode/)
    await expect(page.locator('.ribbon')).toBeVisible()

    // Insert > Number writes the number in the chosen type at the caret
    await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Number$/ })
      .click()
    const number = page.locator('.number-dialog')
    await number.locator('input[name="number"]').fill('14')
    await number.locator('select[name="type"]').selectOption('upperRoman')
    await page.screenshot({ path: test.info().outputPath('number.png') })
    await ok(number)
    await expect(before).toContainText('!XIV')

    // Format > Text Direction is there in body text too; in a table cell it turns the cell's text
    await expect.poll(() => find(['Format', 'Text Direction…'])).toMatchObject({ enabled: true })
    const cell = page.locator('.doc-page .doc-table').first().locator('td, th').first()
    await cell.click()
    await expect.poll(() => find(['Format', 'Text Direction…'])).toMatchObject({ enabled: true })
    await click(['Format', 'Text Direction…'])
    const direction = page.locator('.text-direction-dialog')
    await direction.locator('input[name="direction"]').nth(1).check()
    await page.screenshot({ path: test.info().outputPath('text-direction.png') })
    await ok(direction)
    await expect.poll(() => cell.evaluate((el) => el.outerHTML.includes('vertical-rl'))).toBe(true)
    // the dialog opens on the direction the cell has, so OK leaves a turned cell turned
    await click(['Format', 'Text Direction…'])
    await expect(direction.locator('input[name="direction"]').nth(1)).toBeChecked()
    await ok(direction)
    await expect(direction).toBeHidden()
    expect(await cell.evaluate((el) => el.outerHTML.includes('vertical-rl'))).toBe(true)

    // File > Rename… keeps the dialog open on a name that cannot be used, then renames the file
    const rename = page.locator('.rename-dialog')
    await click(['File', 'Rename…'])
    await expect(rename.locator('input[name="name"]')).toHaveValue('letter')
    await rename.locator('input[name="name"]').fill('a/b')
    await ok(rename)
    await expect(rename.locator('.modal-error')).toHaveText('That name cannot be used')
    await page.screenshot({ path: test.info().outputPath('rename.png') })
    await rename.locator('input[name="name"]').fill('offer')
    await ok(rename)
    await expect(rename).toBeHidden()
    const renamed = join(dir, 'offer.docx')
    await expect.poll(() => existsSync(renamed) && !existsSync(docPath)).toBe(true)

    // File > Move… takes the file to the chosen folder, and the next save lands there
    await app.evaluate(({ dialog }, folder) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [folder] })
    }, archive)
    await click(['File', 'Move…'])
    const moved = join(archive, 'offer.docx')
    await expect.poll(() => existsSync(moved) && !existsSync(renamed)).toBe(true)
    await saveFromMenu(app)
    await expect
      .poll(
        async () =>
          (await (
            await JSZip.loadAsync(readFileSync(moved))
          )
            .file('word/document.xml')
            ?.async('string')) ?? '',
        { timeout: 30_000 },
      )
      .toContain('cherry')
    expect(existsSync(renamed)).toBe(false)
    expect(existsSync(docPath)).toBe(false)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Bar, Area and Radar charts, Borders and Shading, Font Substitution and Rules run', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-batch13-')))
  const docPath = join(dir, 'letter.docx')
  // a section of its own, as every document Word writes has: the page border is saved on it
  const sectPr =
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
  writeFileSync(docPath, await docxWithTable(['cherry'], sectPr))
  const listPath = join(dir, 'people.csv')
  writeFileSync(listPath, 'Name,City\nAda,London\nAlan,Wilmslow\n')
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const cherry = page.locator('.doc-page p', { hasText: 'cherry' }).first()
    await cherry.click()
    const click = (path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items
        let item: Electron.MenuItem | undefined
        for (const label of path) {
          item = items?.find((candidate) => candidate.label === label)
          items = item?.submenu?.items
        }
        item?.click()
      }, path)
    const saved = async (name: string) => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      const parts = Object.keys(zip.files).filter((file) => file.startsWith(name))
      return (await Promise.all(parts.map((file) => zip.file(file)!.async('string')))).join('\n')
    }

    // Format > Borders and Shading: the sides and the fill of the paragraph, the page border as its tab
    const dialog = page.locator('.borders-shading-dialog')
    const tab = (name: string) => page.locator('.bs-tabs button', { hasText: name })
    await click(['Format', 'Borders and Shading…'])
    await dialog.locator('button[data-side="t"]').click()
    await dialog.locator('button[data-side="b"]').click()
    await page.screenshot({ path: test.info().outputPath('borders.png') })
    await tab('Shading').click()
    await dialog.locator('button[data-fill="FFF2CC"]').click()
    await page.screenshot({ path: test.info().outputPath('shading.png') })
    await tab('Page Border').click()
    const pageBorder = page.locator('.pb-dialog')
    const box = pageBorder.locator('.pb-setting', { hasText: /^Box$/ })
    await expect(pageBorder.locator('.bs-tabs')).toBeVisible()
    await box.click()
    await page.screenshot({ path: test.info().outputPath('page-border-tab.png') })
    // the page border picked here survives a look at another tab, and the one OK takes it too
    await tab('Borders').click()
    await expect(dialog.locator('button[data-side="t"]')).toHaveAttribute('aria-pressed', 'true')
    await tab('Page Border').click()
    await expect(box).toHaveClass(/active/)
    await tab('Borders').click()
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toBeHidden()

    // Format > Font Substitution lists what stands in for the fonts this computer lacks
    const fonts = page.locator('.font-substitution-dialog')
    await click(['Format', 'Font Substitution…'])
    await expect(fonts.locator('table, .font-substitution-none')).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('font-substitution.png') })
    await fonts.getByRole('button', { name: 'Close' }).click()
    await expect(fonts).toBeHidden()

    // Insert > Chart > Bar, Area and Radar open the chart dialog on their type
    const chart = page.locator('.modal-chart')
    for (const [item, type] of [
      ['Bar', 'barH'],
      ['Area', 'area'],
      ['Radar', 'radar'],
    ]) {
      await click(['Insert', 'Chart', item])
      await expect(chart.locator(`button[data-chart-type="${type}"]`)).toHaveClass(/btn-primary/)
      if (type === 'barH')
        await page.screenshot({ path: test.info().outputPath('chart-types.png') })
      await chart.locator('.modal-actions button.btn-primary').click()
      await expect(chart).toBeHidden()
    }
    await page.screenshot({ path: test.info().outputPath('charts.png') })

    // Mailings > Rules: waits for a recipient list; then Next Record goes in at once
    await cherry.click()
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, listPath)
    await page.locator('.ribbon-tab', { hasText: /^Mailings$/ }).click()
    const rules = page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: 'Rules',
    })
    const rule = (name: string) => page.locator('.layout-menu button', { hasText: name })
    await expect(rules).toBeDisabled()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Select Recipients' })
      .click()
    await rules.click()
    await page.locator('.layout-menu button', { hasText: /^Next Record$/ }).click()
    await expect(
      page.locator('[data-ribbon-body] [data-testid="mailings-record-status"]', {
        hasText: 'Record 1 of 2',
      }),
    ).toHaveCount(1)
    const ruleDialog = page.locator('.rule-dialog')
    await rules.click()
    await rule('If…Then…Else…').click()
    await ruleDialog.locator('select[name="field"]').selectOption('City')
    await ruleDialog.locator('input[name="value"]').fill('London')
    await ruleDialog.locator('input[name="then"]').fill('See you in town')
    await ruleDialog.locator('input[name="otherwise"]').fill('Safe travels')
    await page.screenshot({ path: test.info().outputPath('rule.png') })
    await ruleDialog.locator('button.primary').click()
    await expect(ruleDialog).toBeHidden()

    // all of it reaches the file
    await saveFromMenu(app)
    await expect
      .poll(() => saved('word/charts/chart'), { timeout: 30_000 })
      .toContain('c:radarChart')
    const charts = await saved('word/charts/chart')
    expect(charts).toContain('<c:barDir val="bar"/>')
    expect(charts).toContain('c:areaChart')
    const body = await saved('word/document.xml')
    expect(body).toMatch(/<w:pBdr>.*<w:top .*<w:bottom /s)
    expect(body).toMatch(/<w:shd [^>]*w:fill="FFF2CC"/)
    expect(body).toContain('<w:pgBorders')
    expect(body).toMatch(/<w:instrText[^>]*>\s*NEXT\s*</)
    expect(body).toContain('See you in town')
    expect(body).toContain('Safe travels')
    expect(body).toMatch(/MERGEFIELD\s+"?City/)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Character Shading, Index and Tables and Hyphenation run, and Hyphenation is saved with the file', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'hyphenation.docx')
  // Hyphenation is offered for a document that declares its language
  writeFileSync(docPath, await withLanguage(await docxWithTable()))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const click = (menu: string, label: string) =>
      app.evaluate(
        ({ Menu }, [menu, label]) => {
          const top = Menu.getApplicationMenu()?.items.find((item) => item.label === menu)
          top?.submenu?.items.find((candidate) => candidate.label === label)?.click()
        },
        [menu, label],
      )
    const before = page.locator('.doc-page p', { hasText: 'Before' }).first()

    // Character Shading: the selected word turns grey and the button lights up
    await before.dblclick()
    const shading = page.locator(
      '[data-ribbon-body] button.rb-icon[aria-label="Character Shading"]',
    )
    await shading.click()
    await expect(shading).toHaveClass(/active/)
    await expect(before.locator('span[style*="background-color"]')).toHaveCount(1)

    // Index and Tables: Word's four tabs
    await click('Insert', 'Index and Tables…')
    const tables = page.locator('.index-tables-dialog')
    await expect(tables.getByRole('tab', { name: 'Table of Contents' })).toBeEnabled()
    await expect(tables.getByRole('tab', { name: 'Table of Figures' })).toBeEnabled()
    await expect(tables.getByRole('tab', { name: 'Table of Authorities' })).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('index-tables.png') })
    await tables.getByRole('button', { name: 'Cancel' }).click()
    await expect(tables).toBeHidden()

    // Hyphenation: on for the page, remembered by the dialog, written to the file
    const hyphens = () =>
      page.evaluate(() => getComputedStyle(document.querySelector('.doc-page')!).hyphens)
    expect(await hyphens()).not.toBe('auto')
    await click('Tools', 'Hyphenation…')
    const hyphenation = page.locator('.hyphenation-dialog')
    const auto = hyphenation.getByLabel('Automatically hyphenate document')
    await expect(auto).not.toBeChecked()
    await auto.check()
    await page.screenshot({ path: test.info().outputPath('hyphenation.png') })
    await hyphenation.locator('button.primary').click()
    await expect.poll(hyphens).toBe('auto')
    await click('Tools', 'Hyphenation…')
    await expect(auto).toBeChecked()
    await hyphenation.getByRole('button', { name: 'Cancel' }).click()

    await before.click()
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          return (await zip.file('word/settings.xml')?.async('string')) ?? ''
        },
        { timeout: 30_000 },
      )
      .toContain('<w:autoHyphenation/>')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Design > Paragraph Spacing re-spaces the body text and is saved with the file', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'spacing.docx')
  writeFileSync(docPath, await docxWithTable(['One', 'Two']))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const first = page.locator('.doc-page p', { hasText: 'One' }).first()
    const second = page.locator('.doc-page p', { hasText: 'Two' }).first()
    await first.click()
    // the distance from one paragraph's top to the next one's: the line plus the space after
    const pitch = async () => (await second.boundingBox())!.y - (await first.boundingBox())!.y

    await page.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
    const spacing = page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', {
      hasText: 'Paragraph Spacing',
    })
    await spacing.click()
    await page.screenshot({ path: test.info().outputPath('paragraph-spacing.png') })
    await page.locator('.layout-menu button', { hasText: /^No Paragraph Space$/ }).click()
    await page.waitForTimeout(800)
    const tight = await pitch()
    await spacing.click()
    await page.locator('.layout-menu button', { hasText: /^Double$/ }).click()
    await expect.poll(pitch, { timeout: 15_000 }).toBeGreaterThan(tight * 1.5)

    await first.click()
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          return (await zip.file('word/styles.xml')?.async('string')) ?? ''
        },
        { timeout: 30_000 },
      )
      .toMatch(
        /<w:spacing[^>]*w:after="160"[^>]*w:line="480"|<w:spacing[^>]*w:line="480"[^>]*w:after="160"/,
      )
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Home > Character Border boxes the selected text and the box is saved with the file', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'border.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const before = page.locator('.doc-page p', { hasText: 'Before' }).first()
    const border = page.locator('[data-ribbon-body] button.rb-icon[aria-label="Character Border"]')
    const documentXml = async () => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      return (await zip.file('word/document.xml')?.async('string')) ?? ''
    }

    await before.dblclick()
    await expect(border).not.toHaveClass(/active/)
    await border.click()
    await expect(border).toHaveClass(/active/)
    await expect(before.locator('span[style*="border"]')).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('character-border.png') })
    await saveFromMenu(app)
    await expect
      .poll(documentXml, { timeout: 30_000 })
      .toContain('<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>')

    // off again: the box goes, and the file says so in so many words
    await before.dblclick()
    await border.click()
    await expect(border).not.toHaveClass(/active/)
    await expect(before.locator('span[style*="border"]')).toHaveCount(0)
    await saveFromMenu(app)
    await expect.poll(documentXml, { timeout: 30_000 }).not.toContain('w:val="single" w:sz="4"')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a table of figures lists the captions, from Index and Tables and the References tab, and updates', async () => {
  test.setTimeout(120_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-menu-command-state' })
  const { page: shell, app } = launched
  try {
    await shell.evaluate(() =>
      (window as never as { aiOffice: { newDoc(): Promise<void> } }).aiOffice.newDoc(),
    )
    const page = await waitForPageWithUrl(app, '://docs/')
    const body = page.locator('main.editor-scroll .ProseMirror').first()
    await body.click()
    await page.keyboard.type('report')
    const insert = (label: string) =>
      app.evaluate(({ Menu }, label) => {
        const menu = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Insert')
        menu?.submenu?.items.find((candidate) => candidate.label === label)?.click()
      }, label)
    const addCaption = async (text: string) => {
      await insert('Caption…')
      const modal = page.locator('.modal', { hasText: 'Caption' }).first()
      await modal.locator('input[name="caption-text"]').fill(text)
      await modal.locator('button.btn-primary').click()
      await expect(modal).toBeHidden()
    }

    await addCaption('Overview')
    await addCaption('Detail')
    await expect(body).toContainText('Figure 2 Detail')

    // Index and Tables: the Table of Figures tab offers the label in use
    await insert('Index and Tables…')
    const dialog = page.locator('.index-tables-dialog')
    await dialog.getByRole('tab', { name: 'Table of Figures' }).click()
    await expect(dialog.locator('[data-it="label"]')).toHaveValue('Figure')
    await page.screenshot({ path: test.info().outputPath('table-of-figures.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    // each caption now shows twice: where it is, and in the table
    await expect(body.getByText('Figure 1 Overview')).toHaveCount(2)
    await expect(body.getByText('Figure 2 Detail')).toHaveCount(2)

    // a caption added later reaches the table through Update Table: the second of the three
    // buttons of that name, between the table of contents' and the table of authorities'
    await addCaption('Appendix')
    await expect(body.getByText('Figure 3 Appendix')).toHaveCount(1)
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Update Table$/ })
      .nth(1)
      .click()
    await expect(body.getByText('Figure 3 Appendix')).toHaveCount(2)
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
  }
})

test('AutoText keeps and re-inserts a selection, Reveal Formatting describes it, and a marked citation reaches a saved table of authorities', async () => {
  test.setTimeout(150_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'authorities.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const before = page.locator('.doc-page p', { hasText: 'Before' }).first()
    await expect(before).toBeVisible({ timeout: 30_000 })
    const click = (...path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items ?? []
        for (const label of path.slice(0, -1))
          items = items.find((item) => item.label === label)?.submenu?.items ?? []
        items.find((item) => item.label === path[path.length - 1])?.click()
      }, path)
    const documentXml = async () => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      return zip.file('word/document.xml')!.async('string')
    }
    const pageText = page.locator('main.editor-scroll .doc-page').first()

    // AutoText > New: the selection, under the name it suggests
    await before.dblclick()
    await click('Insert', 'AutoText', 'New…')
    const autoText = page.locator('.auto-text-dialog')
    await expect(autoText.locator('input')).toHaveValue('Before')
    await autoText.locator('input').fill('Opening')
    await autoText.locator('button.primary').click()
    await expect(autoText).toBeHidden()

    // AutoText: the entry is listed with what it holds, and goes in at the caret
    // typing first: the caret has to have landed before the menu runs
    const after = page.locator('.doc-page p', { hasText: 'After' }).first()
    await after.click()
    await page.keyboard.type('!')
    await expect(after).toHaveText('After!')
    await click('Insert', 'AutoText', 'AutoText…')
    await expect(autoText.locator('select option')).toHaveText(['Opening'])
    await expect(autoText.locator('.auto-text-preview')).toHaveText('Before')
    await page.screenshot({ path: test.info().outputPath('auto-text.png') })
    await autoText.locator('button.primary').click()
    await expect(autoText).toBeHidden()
    await expect(after).toHaveText('After!Before')

    // deleting the entry empties the list
    await click('Insert', 'AutoText', 'AutoText…')
    await autoText.getByRole('button', { name: 'Delete' }).click()
    await expect(autoText.locator('select')).toHaveCount(0)
    await expect(autoText.locator('button.primary')).toBeDisabled()
    await autoText.getByRole('button', { name: 'Cancel' }).click()
    await expect(autoText).toBeHidden()

    // Reveal Formatting: the font and the paragraph at the caret
    await click('View', 'Reveal Formatting')
    const reveal = page.locator('.reveal-formatting-dialog')
    await expect(reveal.locator('.reveal-formatting-row', { hasText: 'Alignment' })).toBeVisible()
    await expect(
      reveal.locator('.reveal-formatting-row', { hasText: 'Style' }).last(),
    ).toContainText('Normal')
    await expect(reveal.locator('.reveal-formatting-row', { hasText: 'Size' })).toContainText('pt')
    await page.screenshot({ path: test.info().outputPath('reveal-formatting.png') })
    await reveal.locator('button.primary').click()
    await expect(reveal).toBeHidden()

    // Mark Citation: the selected word, as a case
    await before.dblclick()
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Mark Citation$/ })
      .click()
    const citation = page.locator('.mark-citation-dialog')
    await expect(citation.locator('input[name="long"]')).toHaveValue('Before')
    await citation.locator('input[name="long"]').fill('Baldwin v. Alberti')
    await page.screenshot({ path: test.info().outputPath('mark-citation.png') })
    await citation.locator('button.primary').click()
    await expect(citation).toBeHidden()

    // Index and Tables > Table of Authorities: the category's heading, then the citation
    await click('Insert', 'Index and Tables…')
    const tables = page.locator('.index-tables-dialog')
    await tables.getByRole('tab', { name: 'Table of Authorities' }).click()
    // every category unless one is chosen
    await expect(tables.getByLabel('Category')).toHaveValue('0')
    await tables.getByLabel('Category').selectOption('1')
    await page.screenshot({ path: test.info().outputPath('table-of-authorities.png') })
    await tables.locator('button.primary').click()
    await expect(tables).toBeHidden()
    await expect(pageText.getByText('Cases', { exact: true })).toHaveCount(1)
    await expect(pageText.getByText('Baldwin v. Alberti')).toHaveCount(1)

    // Update Table, the third button of that name, rebuilds it in place
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: /^Update Table$/ })
      .nth(2)
      .click()
    await expect(pageText.getByText('Baldwin v. Alberti')).toHaveCount(1)

    await saveFromMenu(app)
    await expect.poll(documentXml, { timeout: 30_000 }).toContain('TOA \\h \\p \\c "1"')
    expect(await documentXml()).toMatch(
      /TA \\l (?:"|&quot;)Baldwin v\. Alberti(?:"|&quot;) \\s (?:"|&quot;)Before(?:"|&quot;) \\c 1/,
    )
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Language marks the selected text and saves it; the Selection Pane selects a text box; Format Object shows its tab; the Citations pane cites a source', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'language.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const before = page.locator('.doc-page p', { hasText: 'Before' }).first()
    await expect(before).toBeVisible({ timeout: 30_000 })
    const click = (...path: string[]) =>
      app.evaluate(({ Menu }, path) => {
        let items = Menu.getApplicationMenu()?.items ?? []
        for (const label of path.slice(0, -1))
          items = items.find((item) => item.label === label)?.submenu?.items ?? []
        items.find((item) => item.label === path[path.length - 1])?.click()
      }, path)
    const documentXml = async () => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      return (await zip.file('word/document.xml')?.async('string')) ?? ''
    }

    // Language: the selected word is marked French, the dialog remembers it, the file says so
    await before.dblclick()
    await click('Tools', 'Language…')
    const language = page.locator('.language-dialog')
    const list = language.locator('select')
    await expect(list).toBeVisible()
    await list.selectOption('fr-FR')
    await page.screenshot({ path: test.info().outputPath('language.png') })
    await language.locator('button.primary').click()
    await expect(language).toBeHidden()
    // the word is still selected, and a double-click on a selected word leaves a caret:
    // the caret goes elsewhere first
    await page.locator('.doc-page p', { hasText: 'After' }).first().click()
    await expect
      .poll(() => page.evaluate(() => document.getSelection()?.anchorNode?.textContent))
      .toBe('After')
    await before.dblclick()
    await click('Tools', 'Language…')
    await expect(list).toHaveValue('fr-FR')
    await language.getByRole('button', { name: 'Cancel' }).click()
    await expect(language).toBeHidden()
    await saveFromMenu(app)
    await expect.poll(documentXml, { timeout: 30_000 }).toContain('<w:lang w:val="fr-FR"/>')

    // Selection Pane: nothing to list yet
    await click('Format', 'Selection Pane…')
    const pane = page.locator('.selection-pane')
    await expect(pane).toContainText('no pictures, text boxes or shapes')

    // a text box drawn on the page shows up as a row
    await click('Insert', 'Text Box', 'Text Box')
    const box = await page.locator('main.editor-scroll .doc-page').first().boundingBox()
    await page.mouse.click(box!.x + box!.width / 2, box!.y + 320)
    const row = pane.locator('.selection-pane-row')
    await expect(row).toHaveCount(1)

    // with the caret back in the text the row is not the selection; a click on it is
    await before.click()
    await expect(row).not.toHaveClass(/active/)
    // Format Object has nothing to format there, and says so
    await click('Format', 'Format Object…')
    await expect(page.getByText('Select a picture, text box or shape first.')).toBeVisible()
    await row.click()
    await expect(row).toHaveClass(/active/)
    await expect(page.locator('.doc-protected.ProseMirror-selectednode')).toHaveCount(1)

    // with the box selected it brings its format tab back from wherever the ribbon is
    const shapeFormat = page.locator('.ribbon-tab', { hasText: /^Shape Format$/ })
    await page.locator('.ribbon-tab', { hasText: /^Home$/ }).click()
    await expect(shapeFormat).not.toHaveClass(/active/)
    await click('Format', 'Format Object…')
    await expect(shapeFormat).toHaveClass(/active/)
    await page.screenshot({ path: test.info().outputPath('selection-pane.png') })

    // the command closes the pane it opened
    await click('Format', 'Selection Pane…')
    await expect(pane).toBeHidden()

    // Citations: the pane lists the document's sources, and a click cites one at the caret
    await page.locator('.ribbon-tab', { hasText: /^References$/ }).click()
    const references = (name: RegExp) =>
      page.locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: name })
    await references(/^Citations$/).click()
    const citations = page.locator('.citations-pane')
    await expect(citations).toContainText('No sources yet')
    await references(/^Insert Citation$/).click()
    await page.locator('[data-rb-panel] button', { hasText: 'Add New Source…' }).click()
    const source = page.locator('.modal', { hasText: 'Author' }).first()
    await source.locator('input').nth(0).fill('Ada Lovelace')
    await source.locator('input').nth(1).fill('Notes on the Analytical Engine')
    await source.locator('input').nth(2).fill('1843')
    await source.locator('button.btn-primary').click()
    await expect(source).toBeHidden()
    const cited = citations.locator('.selection-pane-row')
    await expect(cited).toHaveCount(1)
    await expect(cited).toContainText('Notes on the Analytical Engine')
    const lastLine = page.locator('.doc-page p', { hasText: 'After' }).first()
    await lastLine.click()
    await page.keyboard.type('!')
    await expect(lastLine).toContainText('After!')
    const citedBefore = (await lastLine.innerText()).split('Lovelace').length
    await cited.click()
    await expect
      .poll(async () => (await lastLine.innerText()).split('Lovelace').length)
      .toBe(citedBefore + 1)
    await page.screenshot({ path: test.info().outputPath('citations-pane.png') })
    await references(/^Citations$/).click()
    await expect(citations).toBeHidden()
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Save as Template writes a .dotx and leaves the document alone; New from Template starts an untitled document that saves as a .docx', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'letter.docx')
  const templatePath = join(dir, 'letter.dotx')
  const madePath = join(dir, 'from-template.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const before = page.locator('.doc-page p', { hasText: 'Before' }).first()
    await expect(before).toBeVisible({ timeout: 30_000 })
    const click = (label: string) =>
      app.evaluate(({ Menu }, label) => {
        const file = Menu.getApplicationMenu()?.items.find((item) => item.label === 'File')
        file?.submenu?.items.find((candidate) => candidate.label === label)?.click()
      }, label)
    const saveTo = (file: string) =>
      app.evaluate(({ dialog }, file) => {
        dialog.showSaveDialog = async () => ({ canceled: false, filePath: file })
      }, file)
    const mainType = async (file: string) => {
      if (!existsSync(file)) return ''
      const zip = await JSZip.loadAsync(readFileSync(file))
      const types = (await zip.file('[Content_Types].xml')?.async('string')) ?? ''
      return /wordprocessingml\.(document|template)\.main\+xml/.exec(types)?.[1] ?? ''
    }

    // an edit that is not saved yet goes into the template, and stays unsaved in the document
    await before.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' greeting')
    await expect(before).toHaveText('Before greeting')
    await saveTo(templatePath)
    await click('Save as Template…')
    await expect.poll(() => mainType(templatePath), { timeout: 30_000 }).toBe('template')
    const templateXml = await (
      await JSZip.loadAsync(readFileSync(templatePath))
    )
      .file('word/document.xml')!
      .async('string')
    expect(templateXml).toContain('greeting')
    expect(await mainType(docPath)).toBe('document')
    const onDisk = await (
      await JSZip.loadAsync(readFileSync(docPath))
    )
      .file('word/document.xml')!
      .async('string')
    expect(onDisk).not.toContain('greeting')
    await saveFromMenu(app)
    await expect
      .poll(
        async () =>
          (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')!.async('string'),
        { timeout: 30_000 },
      )
      .toContain('greeting')

    // New from Template: the template's content in a document with no path of its own
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, templatePath)
    await before.click()
    await page.keyboard.press('End')
    await page.keyboard.type('!')
    await expect(before).toHaveText('Before greeting!')
    await saveFromMenu(app)
    // saved first, so replacing the document has nothing to ask about
    await expect
      .poll(
        async () =>
          (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')!.async('string'),
        { timeout: 30_000 },
      )
      .toContain('greeting!')
    await click('New from Template…')
    // the template is not in the templates folder, so it is not in the gallery: the file picker has it
    await page.locator('.template-gallery-dialog button[name="browse"]').click()
    await expect(before).toHaveText('Before greeting')

    // saving it asks where, writes a document, and never touches the template
    const after = page.locator('.doc-page p', { hasText: 'After' }).first()
    await after.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' made')
    await expect(after).toHaveText('After made')
    await saveTo(madePath)
    await click('Save As…')
    await expect.poll(() => mainType(madePath), { timeout: 30_000 }).toBe('document')
    const made = await (
      await JSZip.loadAsync(readFileSync(madePath))
    )
      .file('word/document.xml')!
      .async('string')
    expect(made).toContain('made')
    const templateAfter = await (
      await JSZip.loadAsync(readFileSync(templatePath))
    )
      .file('word/document.xml')!
      .async('string')
    expect(templateAfter).not.toContain('made')
    expect(await mainType(templatePath)).toBe('template')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Phonetic Guide puts a reading on the selected text, changes it, saves it and clears it', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'ruby.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const before = page.locator('.doc-page p', { hasText: 'Before' }).first()
    const open = () =>
      app.evaluate(({ Menu }) => {
        const format = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Format')
        const asian = format?.submenu?.items.find((item) => item.label === 'Asian Layout')
        asian?.submenu?.items.find((item) => item.label === 'Phonetic Guide…')?.click()
      })
    const dialog = page.locator('.phonetic-guide-dialog')
    const reading = dialog.getByLabel('Ruby text', { exact: true })
    const guide = page.locator('.doc-page ruby.doc-ruby')

    // with nothing selected the status bar says what it needs
    await before.click()
    await open()
    await expect(page.locator('.status-msg')).toContainText('Select the text')

    await before.dblclick()
    await open()
    await expect(dialog.getByLabel('Base text')).toHaveValue('Before')
    await reading.fill('bi-for')
    await page.screenshot({ path: test.info().outputPath('phonetic-guide.png') })
    await dialog.locator('button.primary').click()
    await expect(guide.locator('rt')).toHaveText('bi-for')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          return (await zip.file('word/document.xml')?.async('string')) ?? ''
        },
        { timeout: 30_000 },
      )
      .toMatch(/<w:ruby>.*<w:t[^>]*>bi-for<\/w:t>.*<w:t[^>]*>Before<\/w:t>.*<\/w:ruby>/)

    // the guide selected: the dialog opens on its reading, and changes it
    await guide.click()
    await open()
    await expect(reading).toHaveValue('bi-for')
    await reading.fill('before')
    await dialog.locator('button.primary').click()
    await expect(guide.locator('rt')).toHaveText('before')

    // Clear Readings leaves the text as it was; the changed guide is still the selection
    await open()
    await dialog.getByRole('button', { name: 'Clear Readings' }).click()
    await expect(guide).toHaveCount(0)
    await expect(before).toHaveText('Before')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Two Lines in One, Fit Text and the Asian Layout button set the layout of selected text, and it is saved', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'asian.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const before = page.locator('.doc-page p', { hasText: 'Before' }).first()
    const after = page.locator('.doc-page p', { hasText: 'After' }).first()
    const asian = (label: string) =>
      app.evaluate(({ Menu }, label) => {
        const format = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Format')
        const sub = format?.submenu?.items.find((item) => item.label === 'Asian Layout')
        sub?.submenu?.items.find((item) => item.label === label)?.click()
      }, label)
    const documentXml = async () => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      return (await zip.file('word/document.xml')?.async('string')) ?? ''
    }

    // with nothing selected the status bar says what is needed
    await before.click()
    await asian('Two Lines in One…')
    await expect(page.locator('.status-msg')).toContainText('Select the text first')

    // Two Lines in One, from the menu: the word is drawn half size in a box two lines tall
    await before.dblclick()
    await asian('Two Lines in One…')
    const twoLines = page.locator('.two-lines-dialog')
    await twoLines.locator('button.primary').click()
    const combined = before.locator('span[style*="word-break"]')
    await expect(combined).toHaveCount(1)
    await expect(combined).toHaveCSS('display', 'inline-block')

    // Fit Text, from the Home tab's Asian Layout button: the word spreads across the width given
    await after.dblclick()
    await page.locator('[data-ribbon-body] button.rb-icon[aria-label="Asian Layout"]').click()
    await page.locator('.layout-menu button', { hasText: 'Fit Text…' }).click()
    const dialog = page.locator('.fit-text-dialog')
    await dialog.getByLabel('New text width').fill('2 in')
    await page.screenshot({ path: test.info().outputPath('fit-text.png') })
    await dialog.locator('button.primary').click()
    const fitted = after.locator('span[style*="text-align-last"]')
    await expect(fitted).toHaveCSS('width', '192px')

    await saveFromMenu(app)
    await expect.poll(documentXml, { timeout: 30_000 }).toContain('<w:fitText w:val="2880"/>')
    expect(await documentXml()).toMatch(/<w:eastAsianLayout w:id="\d+" w:combine="1"\/>/)

    // the dialog opens on the width in use, and Remove takes the fitting off
    await after.dblclick()
    await asian('Fit Text…')
    await expect(dialog.getByLabel('New text width')).toHaveValue('2"')
    await dialog.getByRole('button', { name: 'Remove' }).click()
    await expect(fitted).toHaveCount(0)
    // the Two Lines in One dialog opens on the layout, and Remove takes it off
    await before.dblclick()
    await asian('Two Lines in One…')
    await twoLines.getByRole('button', { name: 'Remove' }).click()
    await expect(combined).toHaveCount(0)
    await saveFromMenu(app)
    await expect.poll(documentXml, { timeout: 30_000 }).not.toContain('w:fitText')
    expect(await documentXml()).not.toContain('w:eastAsianLayout')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Insert > Chart > Stock and Surface insert their charts, and each is saved as its own type', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'stock.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const dialog = page.locator('.modal-chart')
    const chartParts = async () => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      const names = Object.keys(zip.files).filter((n) => /^word\/charts\/chart\d+\.xml$/.test(n))
      return (await Promise.all(names.map((n) => zip.file(n)!.async('string')))).join('\n')
    }

    // Stock: the dialog starts on the three series a stock chart reads
    await page.locator('.doc-page p', { hasText: 'After' }).first().click()
    await clickInsertItem(app, ['Chart', 'Stock'])
    await expect(dialog.locator('.btn-primary', { hasText: 'Stock' })).toBeVisible()
    const names = dialog.locator('.chart-data-grid tbody th input')
    await expect(names).toHaveCount(3)
    await expect(names.nth(0)).toHaveValue('High')
    await expect(names.nth(2)).toHaveValue('Close')
    await page.screenshot({ path: test.info().outputPath('stock-dialog.png') })
    // a fourth series is the open: it goes in first, and there is no fifth
    const addSeries = dialog.getByRole('button', { name: '+ Series' })
    await addSeries.click()
    await expect(names).toHaveCount(4)
    await expect(names.nth(0)).toHaveValue('Open')
    await expect(names.nth(3)).toHaveValue('Close')
    await expect(addSeries).toBeDisabled()
    for (const [i, v] of ['30', '33', '36'].entries())
      await dialog.locator('.chart-data-grid tbody tr').first().locator('td input').nth(i).fill(v)
    await dialog.locator('.modal-actions .btn-primary').click()
    await expect(dialog).toBeHidden()
    // no line along a series
    const stock = page.locator('.doc-page .doc-chart-canvas').last()
    // per category: the high-low line, the close tick and the open tick
    await expect(stock.locator('svg line[stroke="#404040"]')).toHaveCount(9)
    await expect(stock.locator('polyline, path')).toHaveCount(0)
    await page.screenshot({ path: test.info().outputPath('stock.png') })

    // Surface: a cell per point, keyed by value bands
    await page.locator('.doc-page p', { hasText: 'Before' }).first().click()
    await clickInsertItem(app, ['Chart', 'Surface'])
    await expect(dialog.locator('.btn-primary', { hasText: 'Surface' })).toBeVisible()
    await dialog.locator('.modal-actions .btn-primary').click()
    await expect(dialog).toBeHidden()
    await expect(page.locator('.doc-page rect.doc-chart-surface')).toHaveCount(6)
    await page.locator('.doc-page rect.doc-chart-surface').first().scrollIntoViewIfNeeded()
    await page.screenshot({ path: test.info().outputPath('surface.png') })

    await saveFromMenu(app)
    await expect.poll(chartParts, { timeout: 30_000 }).toContain('<c:stockChart>')
    const saved = await chartParts()
    expect(saved).toContain('<c:hiLowLines/>')
    expect(saved).toContain('<c:surfaceChart>')
    expect(saved).toContain('<c:view3D><c:rotX val="90"/>')
    expect(await savedDocumentParses(page, docPath)).toBe('ok')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Insert > Chart inserts the seven types of the chart extension, each drawn as itself and saved as its own part', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'charts.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const dialog = page.locator('.modal-chart')
    // menu label, the layoutId its series is saved with, and the marks only that drawing has
    const types: Array<[string, string, string, number]> = [
      ['Waterfall', 'waterfall', '.doc-chartex-step', 3],
      ['Funnel', 'funnel', '.doc-chartex-stage', 3],
      ['Treemap', 'treemap', '.doc-chartex-cell', 3],
      ['Box and Whisker', 'boxWhisker', '.doc-chartex-box', 2],
      ['Pareto', 'paretoLine', '.doc-chartex-cumulative', 1],
    ]
    const after = page.locator('.doc-page p', { hasText: 'After' }).first()
    for (const [label, , marks, count] of types) {
      await after.click()
      await clickInsertItem(app, ['Chart', label])
      await expect(dialog).toBeVisible()
      // only the box and whisker reads more than one series: the rest take no second one
      const addSeries = dialog.getByRole('button', { name: '+ Series' })
      if (label === 'Box and Whisker') await expect(addSeries).toBeEnabled()
      else await expect(addSeries).toBeDisabled()
      await dialog.locator('.modal-actions .btn-primary').click()
      await expect(dialog).toBeHidden()
      const drawn = page.locator(`.doc-page ${marks}`)
      await expect(drawn, label).toHaveCount(count)
      await drawn.first().scrollIntoViewIfNeeded()
      await page.screenshot({ path: test.info().outputPath(`chartex-${label}.png`) })
    }
    // a histogram's bins and a sunburst's ring have no mark of their own count to wait on
    for (const label of ['Histogram', 'Sunburst']) {
      await after.click()
      const before = await page.locator('.doc-page .doc-chart-canvas').count()
      await clickInsertItem(app, ['Chart', label])
      await expect(dialog).toBeVisible()
      await dialog.locator('.modal-actions .btn-primary').click()
      await expect(dialog).toBeHidden()
      await expect(page.locator('.doc-page .doc-chart-canvas')).toHaveCount(before + 1)
      // a chart goes in after the paragraph the caret is in: the newest is the first
      await page.locator('.doc-page .doc-chart-canvas').first().scrollIntoViewIfNeeded()
      await page.screenshot({ path: test.info().outputPath(`chartex-${label}.png`) })
    }

    await saveFromMenu(app)
    const parts = async () => {
      const zip = await JSZip.loadAsync(readFileSync(docPath))
      const names = Object.keys(zip.files).filter((n) => /^word\/charts\/chartEx\d+\.xml$/.test(n))
      return (await Promise.all(names.map((n) => zip.file(n)!.async('string')))).join('\n')
    }
    await expect
      .poll(async () => ((await parts()).match(/<cx:chartSpace /g) ?? []).length, {
        timeout: 30_000,
      })
      .toBe(7)
    const saved = await parts()
    for (const [, layoutId] of types) expect(saved, layoutId).toContain(`layoutId="${layoutId}"`)
    expect(saved).toContain('layoutId="sunburst"')
    expect(saved).toContain('<cx:binning intervalClosed="r"/>')
    expect(await savedDocumentParses(page, docPath)).toBe('ok')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Insert > Chart > X Y (Scatter) inserts a scatter chart, Select Data opens its data, and it is saved as one', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-menu-state-')))
  const docPath = join(dir, 'scatter.docx')
  writeFileSync(docPath, await docxWithTable())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-menu-command-state',
    openFile: docPath,
  })
  const { app } = launched
  try {
    const page = await waitForPageWithUrl(app, '://docs/')
    const selectDataItem = () =>
      app.evaluate(({ Menu }) => {
        const edit = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')
        return edit?.submenu?.items.find((candidate) => candidate.label === 'Select Data…')?.enabled
      })
    const selectData = () =>
      app.evaluate(({ Menu }) => {
        const edit = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Edit')
        edit?.submenu?.items.find((candidate) => candidate.label === 'Select Data…')?.click()
      })
    await page.locator('.doc-page p', { hasText: 'After' }).first().click()
    // with no chart in the document Select Data is greyed out
    await expect.poll(selectDataItem).toBe(false)
    await clickInsertItem(app, ['Chart', 'X Y (Scatter)'])
    const dialog = page.locator('.modal-chart')
    // the dialog opens on the type picked
    await expect(dialog.locator('.btn-primary', { hasText: 'Scatter' })).toBeVisible()
    await dialog.locator('.modal-actions .btn-primary').click()
    await expect(dialog).toBeHidden()
    // points, and no line joining them
    const canvas = page.locator('.doc-page .doc-chart-canvas').last()
    // the first series' three points are drawn as its marker shape
    await expect(canvas.locator('polygon')).toHaveCount(3)
    await expect(canvas.locator('polyline, path')).toHaveCount(0)
    await page.screenshot({ path: test.info().outputPath('scatter.png') })

    // Select Data: from the text, it selects the chart, whose data table then shows
    const dataTable = page.locator('.doc-page .doc-chart-data').last()
    await page.locator('.doc-page p', { hasText: 'Before' }).first().click()
    await expect(dataTable).toBeHidden()
    await expect.poll(selectDataItem).toBe(true)
    await selectData()
    await expect(page.locator('.doc-protected.ProseMirror-selectednode')).toHaveCount(1)
    await expect(dataTable).toBeVisible()
    await page.screenshot({ path: test.info().outputPath('select-data.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          const part = Object.keys(zip.files).find((name) =>
            /^word\/charts\/chart\d+\.xml$/.test(name),
          )
          return part ? await zip.file(part)!.async('string') : ''
        },
        { timeout: 30_000 },
      )
      .toContain('<c:scatterChart>')
    expect(await savedDocumentParses(page, docPath)).toBe('ok')
  } finally {
    await closeAndSaveVideo(launched, 'docs-menu-command-state')
    rmSync(dir, { recursive: true, force: true })
  }
})
