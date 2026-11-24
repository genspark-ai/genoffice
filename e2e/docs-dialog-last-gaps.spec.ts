/**
 * Four dialog controls Word has: Insert Cells that shifts cells right or down,
 * the Two Lines in One and Horizontal in Vertical dialogs, Select… in Insert
 * Hyperlink, and Find's "Highlight all items found in". Each is driven
 * through the UI, read off the page, and read back from the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

type EditorWindow = { __aidocs: { editor: import('@tiptap/core').Editor } }

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

/** selects the first occurrence of `text` in the document, or puts the caret at its start */
const selectText = (page: Page, text: string, caretOnly = false) =>
  page.evaluate(
    ([text, caretOnly]) => {
      const { editor } = (window as unknown as EditorWindow).__aidocs
      let from = -1
      editor.state.doc.descendants((node, pos) => {
        if (from < 0 && node.isText && node.text!.includes(text as string))
          from = pos + node.text!.indexOf(text as string)
        return from < 0
      })
      if (from < 0) throw new Error(`no text ${text}`)
      editor
        .chain()
        .focus()
        .setTextSelection(caretOnly ? from : { from, to: from + (text as string).length })
        .run()
    },
    [text, caretOnly] as const,
  )

const partOf = async (docPath: string, name: string): Promise<string> => {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  return (await zip.file(name)?.async('string')) ?? ''
}

async function openDoc(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-dialog-last-gaps',
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page').first()).toBeVisible({ timeout: 30_000 })
  return { launched, app: launched.app, page }
}

const cell = (text: string, tcPr = '') =>
  `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`

const BORDERS =
  '<w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`)
    .join('') +
  '</w:tblBorders>'

const table = (rows: string[], cols: number) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${BORDERS}</w:tblPr><w:tblGrid>` +
  '<w:gridCol w:w="1500"/>'.repeat(cols) +
  `</w:tblGrid>${rows.join('')}</w:tbl>`

/** a table whose second row stops two columns short, as Word writes one after Shift cells right */
const RAGGED = table(
  [
    `<w:tr>${['R1', 'R2', 'R3', 'R4'].map((t) => cell(t)).join('')}</w:tr>`,
    `<w:tr><w:trPr><w:gridAfter w:val="2"/></w:trPr>${['S1', 'S2'].map((t) => cell(t)).join('')}</w:tr>`,
  ],
  4,
)

const GRID = table(
  [1, 2].map((r) => `<w:tr>${['A', 'B', 'C'].map((c) => cell(`${c}${r}`)).join('')}</w:tr>`),
  3,
)

test('Insert Cells shifts cells right and down, and the ragged rows are saved as Word writes them', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-dialog-gaps-')))
  const docPath = join(dir, 'doc.docx')
  writeFileSync(docPath, await docxWith(para('Tables') + GRID + para('Between') + RAGGED))
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    /** the cells of each row of a table, grid placeholders left out */
    const rowsOf = (index: number) =>
      page.evaluate(
        (index) =>
          Array.from(
            document.querySelectorAll('.doc-page table')[index].querySelectorAll('tr'),
            (tr) =>
              Array.from(tr.children)
                .filter((td) => getComputedStyle(td).borderTopStyle !== 'none')
                .map((td) => td.textContent ?? ''),
          ),
        index,
      )
    const dialog = page.locator('.insert-cells-dialog')
    const insert = async (at: string, option: string, shot?: string) => {
      await selectText(page, at, true)
      await clickMenu(app, ['Table', 'Insert', 'Cells…'])
      await dialog.getByLabel(option).check()
      if (shot) await page.screenshot({ path: test.info().outputPath(shot) })
      await dialog.locator('button.primary, button[type=submit]').first().click()
      await expect(dialog).toHaveCount(0)
    }

    // the table from Word shows its short row
    expect(await rowsOf(1)).toEqual([
      ['R1', 'R2', 'R3', 'R4'],
      ['S1', 'S2'],
    ])

    await insert('B1', 'Shift cells right', 'insert-cells-dialog.png')
    expect(await rowsOf(0)).toEqual([
      ['A1', '', 'B1', 'C1'],
      ['A2', 'B2', 'C2'],
    ])
    await insert('A2', 'Shift cells down')
    expect(await rowsOf(0)).toEqual([
      ['A1', '', 'B1', 'C1'],
      ['', 'B2', 'C2'],
      ['A2', '', ''],
    ])
    // the short row of Word's table takes a cell without widening the table
    await insert('S1', 'Shift cells right')
    expect(await rowsOf(1)).toEqual([
      ['R1', 'R2', 'R3', 'R4'],
      ['', 'S1', 'S2'],
    ])
    await page.screenshot({ path: test.info().outputPath('shifted-tables.png') })

    await saveFromMenu(app)
    await expect
      .poll(() => partOf(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('<w:gridAfter w:val="1"/>')
    const xml = await partOf(docPath, 'word/document.xml')
    const tables = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g)!
    const shape = (tbl: string) =>
      tbl.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g)!.map((tr) => ({
        cells: tr.match(/<w:tc>/g)!.length,
        after: /<w:gridAfter w:val="(\d+)"/.exec(tr)?.[1] ?? '0',
      }))
    expect(tables[0].match(/<w:gridCol /g)).toHaveLength(4)
    expect(shape(tables[0])).toEqual([
      { cells: 4, after: '0' },
      { cells: 3, after: '1' },
      { cells: 3, after: '1' },
    ])
    expect(shape(tables[1])).toEqual([
      { cells: 4, after: '0' },
      { cells: 3, after: '1' },
    ])
    await closeAndSaveVideo(session.launched, 'docs-dialog-last-gaps')
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect(page.locator('.doc-page table')).toHaveCount(2)
    expect(await rowsOf(0)).toEqual([
      ['A1', '', 'B1', 'C1'],
      ['', 'B2', 'C2'],
      ['A2', '', ''],
    ])
    expect(await rowsOf(1)).toEqual([
      ['R1', 'R2', 'R3', 'R4'],
      ['', 'S1', 'S2'],
    ])
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-dialog-last-gaps')
    rmSync(dir, { recursive: true, force: true })
  }
})

const ea = (attrs: string, text: string) =>
  `<w:r><w:rPr><w:eastAsianLayout ${attrs}/></w:rPr><w:t>${text}</w:t></w:r>`

const VERTICAL_CELL =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid>' +
  '<w:tr><w:trPr><w:trHeight w:val="2600"/></w:trPr>' +
  '<w:tc><w:tcPr><w:textDirection w:val="tbRl"/></w:tcPr>' +
  '<w:p><w:r><w:t xml:space="preserve">Year 2026 ends</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'

const FROM_WORD =
  '<w:p><w:r><w:t xml:space="preserve">Word wrote </w:t></w:r>' +
  ea('w:id="11" w:combine="1" w:combineBrackets="square"', 'bracketed') +
  '<w:r><w:t xml:space="preserve"> and </w:t></w:r>' +
  ea('w:id="12" w:vert="1" w:vertCompress="1"', '77') +
  '</w:p>'

test('Two Lines in One and Horizontal in Vertical open their dialogs, and brackets and fitting are drawn and saved', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-dialog-gaps-')))
  const docPath = join(dir, 'doc.docx')
  writeFileSync(
    docPath,
    await docxWith(para('Alpha note beta') + FROM_WORD + VERTICAL_CELL + para('Last')),
  )
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    const bracketed = (style: string) => page.locator(`.doc-page span[data-ea-brackets="${style}"]`)
    const before = (el: ReturnType<typeof bracketed>) =>
      el.evaluate((node) => getComputedStyle(node, '::before').content)
    const documentXml = () => partOf(docPath, 'word/document.xml')

    // what Word wrote is drawn: square brackets around the two lines
    await expect(bracketed('square')).toHaveText('bracketed')
    expect(await before(bracketed('square'))).toBe('"["')

    // ---- Two Lines in One ----
    const twoLines = () => page.locator('.two-lines-dialog')
    await selectText(page, 'note')
    await clickMenu(app, ['Format', 'Asian Layout', 'Two Lines in One…'])
    await expect(twoLines().getByRole('button', { name: 'Remove' })).toBeDisabled()
    await expect(twoLines().getByLabel('Bracket style')).toBeDisabled()
    await twoLines().getByLabel('Enclose in brackets').check()
    await expect(twoLines().getByLabel('Bracket style').locator('option')).toHaveText([
      '( )',
      '[ ]',
      '< >',
      '{ }',
    ])
    await twoLines().getByLabel('Bracket style').selectOption('curly')
    await expect(twoLines().locator('.eq-layout-preview [data-ea-brackets="curly"]')).toHaveText(
      'note',
    )
    await page.screenshot({ path: test.info().outputPath('two-lines-dialog.png') })
    await twoLines().locator('button.primary').click()
    await expect(bracketed('curly')).toHaveText('note')
    expect(await before(bracketed('curly'))).toBe('"{"')
    await expect(bracketed('curly')).toHaveCSS('display', 'inline-block')

    // ---- Horizontal in Vertical, in a cell whose text runs down the page ----
    const hiv = () => page.locator('.hiv-dialog')
    const upright = () => page.locator('.doc-page td span[data-ea-vert]')
    await selectText(page, '2026')
    await clickMenu(app, ['Format', 'Asian Layout', 'Horizontal in Vertical…'])
    await expect(hiv().getByLabel('Fit in line')).toBeChecked()
    await page.screenshot({ path: test.info().outputPath('horizontal-in-vertical-dialog.png') })
    await hiv().locator('button.primary').click()
    await expect(upright()).toHaveText('2026')
    const fitted = (await upright().boundingBox())!
    await page.screenshot({ path: test.info().outputPath('layouts-on-page.png') })

    await saveFromMenu(app)
    await expect
      .poll(documentXml, { timeout: 30_000 })
      .toMatch(/<w:eastAsianLayout w:id="\d+" w:combine="1" w:combineBrackets="curly"\/>/)
    expect(await documentXml()).toMatch(
      /<w:eastAsianLayout w:id="\d+" w:vert="1" w:vertCompress="1"\/>/,
    )
    // untouched runs keep what Word wrote
    expect(await documentXml()).toContain(
      '<w:eastAsianLayout w:id="11" w:combine="1" w:combineBrackets="square"/>',
    )
    expect(await documentXml()).toContain(
      '<w:eastAsianLayout w:id="12" w:vert="1" w:vertCompress="1"/>',
    )

    await closeAndSaveVideo(session.launched, 'docs-dialog-last-gaps')
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect(bracketed('curly')).toHaveText('note')
    await expect(upright()).toHaveText('2026')

    // the dialog opens on what the text has
    await selectText(page, '2026')
    await clickMenu(app, ['Format', 'Asian Layout', 'Horizontal in Vertical…'])
    await expect(hiv().getByLabel('Fit in line')).toBeChecked()
    await hiv().getByLabel('Fit in line').uncheck()
    await hiv().locator('button.primary').click()
    await expect(upright()).toHaveText('2026')
    // Word draws the run at its size either way: the fitting is kept for the file only
    const loose = (await upright().boundingBox())!
    expect(Math.abs(loose.width - fitted.width)).toBeLessThan(1)
    expect(loose.width).toBeGreaterThan(loose.height)

    // brackets change in place, then Remove takes each layout off
    await selectText(page, 'note')
    await clickMenu(app, ['Format', 'Asian Layout', 'Two Lines in One…'])
    await expect(twoLines().getByLabel('Enclose in brackets')).toBeChecked()
    await expect(twoLines().getByLabel('Bracket style')).toHaveValue('curly')
    await twoLines().getByLabel('Enclose in brackets').uncheck()
    await twoLines().locator('button.primary').click()
    await expect(bracketed('curly')).toHaveCount(0)
    await saveFromMenu(app)
    await expect
      .poll(documentXml, { timeout: 30_000 })
      .toMatch(/<w:eastAsianLayout w:id="\d+" w:vert="1"\/>/)
    expect(await documentXml()).toMatch(/<w:eastAsianLayout w:id="\d+" w:combine="1"\/>/)
    expect(await documentXml()).not.toContain('curly')

    await selectText(page, '2026')
    await clickMenu(app, ['Format', 'Asian Layout', 'Horizontal in Vertical…'])
    await expect(hiv().getByLabel('Fit in line')).not.toBeChecked()
    await hiv().getByRole('button', { name: 'Remove' }).click()
    await expect(upright()).toHaveCount(0)
    await selectText(page, 'note')
    await clickMenu(app, ['Format', 'Asian Layout', 'Two Lines in One…'])
    await twoLines().getByRole('button', { name: 'Remove' }).click()
    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml()).match(/<w:eastAsianLayout /g)?.length, {
        timeout: 30_000,
      })
      .toBe(2)
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-dialog-last-gaps')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('text selected as a save runs is still selected when the save is over, and its dialog opens on it', async () => {
  test.setTimeout(120_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-dialog-gaps-')))
  const docPath = join(dir, 'doc.docx')
  writeFileSync(docPath, await docxWith(para('Alpha note beta') + VERTICAL_CELL + para('Last')))
  const session = await openDoc(docPath)
  try {
    const { app, page } = session
    const selected = () =>
      page.evaluate(() => {
        const { state } = (window as unknown as EditorWindow).__aidocs.editor
        return state.doc.textBetween(state.selection.from, state.selection.to)
      })
    // a layout the save writes with an id of its own: the page is drawn again from the saved file
    await selectText(page, '2026')
    await clickMenu(app, ['Format', 'Asian Layout', 'Horizontal in Vertical…'])
    await page.locator('.hiv-dialog button.primary').click()
    await expect(page.locator('.doc-page td span[data-ea-vert]')).toHaveText('2026')

    await saveFromMenu(app)
    await selectText(page, 'note')
    await expect(page.locator('.status-msg')).toContainText('Saved')
    expect(await selected()).toBe('note')
    await clickMenu(app, ['Format', 'Asian Layout', 'Two Lines in One…'])
    await expect(
      page.locator('.two-lines-dialog .eq-layout-preview .ea-two-lines-sample'),
    ).toHaveText('note')
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-dialog-last-gaps-selection')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Select… in Insert Hyperlink links to a local file, and Find keeps every hit highlighted when asked', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-dialog-gaps-')))
  const docPath = join(dir, 'doc.docx')
  const target = join(dir, 'Annual report 2026.pdf')
  writeFileSync(target, '%PDF-1.4\n')
  writeFileSync(
    docPath,
    await docxWith(
      para('See the report here') + para('alpha one alpha') + para('alpha two') + para('Last'),
    ),
  )
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    const url = pathToFileURL(target).href

    // ---- Insert Hyperlink > Select… ----
    await app.evaluate(({ dialog }, file) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
    }, target)
    await selectText(page, 'here')
    await clickMenu(app, ['Insert', 'Hyperlink…'])
    const link = page.locator('.link-dialog')
    await link.getByRole('button', { name: 'Select…' }).click()
    await expect(link.locator('input[name="linkAddress"]')).toHaveValue(url)
    await page.screenshot({ path: test.info().outputPath('hyperlink-select.png') })
    await link.locator('button.btn-primary').click()
    const anchor = page.locator('.doc-page a.doc-link')
    await expect(anchor).toHaveText('here')
    await expect(anchor).toHaveAttribute('href', url)

    // with no text selected the link shows the name of the file
    await selectText(page, 'Last', true)
    await clickMenu(app, ['Insert', 'Hyperlink…'])
    await link.getByRole('button', { name: 'Select…' }).click()
    await link.locator('button.btn-primary').click()
    await expect(anchor).toHaveText(['here', 'Annual report 2026.pdf'])

    await saveFromMenu(app)
    const rels = () => partOf(docPath, 'word/_rels/document.xml.rels')
    await expect.poll(rels, { timeout: 30_000 }).toContain(url)
    expect(await rels()).toMatch(
      new RegExp(
        `Type="[^"]*/hyperlink" Target="${url.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}" TargetMode="External"`,
      ),
    )
    expect((await partOf(docPath, 'word/document.xml')).match(/<w:hyperlink /g)).toHaveLength(2)

    await closeAndSaveVideo(session.launched, 'docs-dialog-last-gaps')
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect(page.locator('.doc-page a.doc-link').first()).toHaveAttribute('href', url)

    // ---- Find: Highlight all items found in ----
    const hits = page.locator('.doc-page .search-hit')
    const panel = page.locator('.find-panel')
    const keep = panel.locator('input[name="find-highlight-all"]')
    const find = async (text: string) => {
      await clickMenu(app, ['Edit', 'Find', 'Find…'])
      await panel.locator('.find-input').first().fill(text)
    }
    await selectText(page, 'Last', true)
    await find('alpha')
    await expect(panel.locator('.find-count')).toHaveText('1/3')
    await expect(hits).toHaveCount(3)
    await expect(keep).not.toBeChecked()
    await expect(panel.getByText('Highlight all items found in')).toBeVisible()
    // unticked, the hits go with the panel
    await page.keyboard.press('Escape')
    await expect(panel).toHaveCount(0)
    await expect(hits).toHaveCount(0)

    await find('alpha')
    await keep.check()
    await page.screenshot({ path: test.info().outputPath('find-highlight-all.png') })
    await page.keyboard.press('Escape')
    await expect(panel).toHaveCount(0)
    await expect(hits).toHaveCount(3)
    await expect(page.locator('.doc-page .search-hit-active')).toHaveCount(0)
    await page.screenshot({ path: test.info().outputPath('find-hits-kept.png') })
    // a click in the document lets them go
    await page.locator('.doc-page p', { hasText: 'See the report' }).click()
    await expect(hits).toHaveCount(0)

    // kept to the selection: only the hits inside it stay
    await selectText(page, 'alpha one alpha')
    await find('alpha')
    await panel.locator('select[name="find-scope"]').selectOption('selection')
    await expect(panel.locator('.find-count')).toHaveText('1/2')
    await keep.check()
    await page.keyboard.press('Escape')
    await expect(hits).toHaveCount(2)
    // a key in the document lets them go too
    await page.evaluate(() => (window as unknown as EditorWindow).__aidocs.editor.commands.focus())
    await expect
      .poll(() => page.evaluate(() => !!document.activeElement?.closest('.ProseMirror')))
      .toBe(true)
    await page.keyboard.press('ArrowRight')
    await expect(hits).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-dialog-last-gaps')
    rmSync(dir, { recursive: true, force: true })
  }
})
