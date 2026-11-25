/**
 * Format > Bullets and Numbering, List Styles tab: the built-in list styles and
 * the document's own are listed with a preview, OK numbers the selected
 * paragraphs, the save links the list to the style, and the file reopens the same.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-list-styles'
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const BODY = para('alpha') + para('beta') + para('gamma') + para('delta')

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
      state: { doc: { content: { size: number } } }
      chain(): {
        focus(): {
          setTextSelection(r: object | number): { run(): boolean }
          updateAttributes(type: string, attrs: object): { run(): boolean }
        }
      }
    }
  }
}

/** selects from the first paragraph through the paragraph reading `through`, or puts the caret at that paragraph's end */
const selectThrough = (page: Page, through: string, caretOnly = false) =>
  page.evaluate(
    ([through, caretOnly]) => {
      const { editor } = (window as never as EditorWindow).__aidocs
      let to = 1
      ;(
        editor.state.doc as never as {
          descendants(f: (n: { textContent: string }, pos: number) => boolean): void
        }
      ).descendants((node, pos) => {
        if (node.textContent === through) to = pos + 1 + through.length
        return false
      })
      editor
        .chain()
        .focus()
        .setTextSelection({ from: caretOnly ? to : 1, to })
        .run()
    },
    [through, caretOnly] as const,
  )

async function openDocs(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page').first()).toContainText('alpha', { timeout: 30_000 })
  return { launched, page }
}

async function openListStyles(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Format', 'Bullets and Numbering…'])
  const dialog = page.locator('.bn-dialog')
  await dialog.getByRole('tab', { name: 'List Styles' }).click()
  return dialog
}

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

const markers = (page: Page) =>
  page
    .locator('.doc-page .doc-li')
    .evaluateAll((items) => items.map((el) => el.getAttribute('data-marker')))

test('a built-in list style numbers the selection, is written as a style-linked list, and reopens the same', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-list-styles-')))
  const docPath = join(dir, 'plain.docx')
  writeFileSync(docPath, await docxWith(BODY))
  let numId = ''
  const first = await openDocs(docPath)
  try {
    const { launched, page } = first
    await selectThrough(page, 'gamma')
    const dialog = await openListStyles(launched.app, page)
    const options = dialog.locator('.bn-styles-list [role=option]')
    await expect(options).toHaveText([
      'No List',
      '1 / 1.1 / 1.1.1',
      '1 / a / i',
      'Article / Section',
    ])
    await expect(dialog.getByRole('button', { name: 'Add' })).toBeEnabled()
    await expect(dialog.getByRole('button', { name: 'Remove' })).toBeDisabled()
    // nothing is picked yet
    await expect(dialog.getByRole('button', { name: 'Modify…' })).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Customize…' })).toHaveCount(0)
    const preview = dialog.locator('.bn-styles-preview > span')
    await expect(preview).toHaveCount(0)
    await options.nth(3).click()
    await expect(preview.nth(1)).toContainText('Section 1.01')
    await options.nth(2).click()
    await expect(preview).toHaveCount(9)
    await expect(preview.nth(0)).toContainText('1)')
    await expect(preview.nth(1)).toContainText('a)')
    await expect(preview.nth(2)).toContainText('i)')
    await page.screenshot({ path: test.info().outputPath('list-styles-tab.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()

    await expect.poll(() => markers(page)).toEqual(['1)', '2)', '3)'])
    await expect(page.locator('.doc-page p', { hasText: 'delta' })).not.toHaveClass(/doc-li/)
    // gamma one level down takes the style's second level
    await selectThrough(page, 'gamma', true)
    await page.evaluate(() => {
      const { editor } = (window as never as EditorWindow).__aidocs
      editor.chain().focus().updateAttributes('docListItem', { ilvl: 1 }).run()
    })
    await expect.poll(() => markers(page)).toEqual(['1)', '2)', 'a)'])
    await page.screenshot({ path: test.info().outputPath('list-style-applied.png') })

    await saveFromMenu(launched.app)
    await expect
      .poll(() => part(docPath, 'word/numbering.xml'), { timeout: 30_000 })
      .toContain('<w:numStyleLink w:val="1ai"/>')
    const numbering = await part(docPath, 'word/numbering.xml')
    const linking =
      /<w:abstractNum w:abstractNumId="(\d+)"><w:multiLevelType w:val="multilevel"\/><w:numStyleLink w:val="1ai"\/><\/w:abstractNum>/.exec(
        numbering,
      )
    expect(linking).not.toBeNull()
    const num = new RegExp(
      `<w:num w:numId="(\\d+)"><w:abstractNumId w:val="${linking![1]}"/></w:num>`,
    ).exec(numbering)
    expect(num).not.toBeNull()
    expect(numbering).toMatch(/<w:styleLink w:val="1ai"\/><w:lvl w:ilvl="0">/)
    const styles = await part(docPath, 'word/styles.xml')
    const style =
      /<w:style w:type="numbering" w:styleId="1ai"><w:name w:val="Outline List 1"\/>.*?<w:numId w:val="(\d+)"\/>/.exec(
        styles,
      )
    expect(style).not.toBeNull()
    expect(numbering).toMatch(
      new RegExp(`<w:num w:numId="${style![1]}"><w:abstractNumId w:val="\\d+"/></w:num>`),
    )
    const body = await part(docPath, 'word/document.xml')
    expect(body.match(new RegExp(`<w:numId w:val="${num![1]}"/>`, 'g'))).toHaveLength(3)
    expect(await part(docPath, 'word/_rels/document.xml.rels')).toContain('Target="styles.xml"')
    expect(await part(docPath, '[Content_Types].xml')).toContain('PartName="/word/styles.xml"')

    numId = num![1]
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }
  const second = await openDocs(docPath)
  try {
    const { launched, page } = second
    await expect.poll(() => markers(page)).toEqual(['1)', '2)', 'a)'])
    // the caret's list style is the selected one, and the only one in use besides No List
    await selectThrough(page, 'alpha', true)
    const dialog = await openListStyles(launched.app, page)
    await expect(dialog.locator('.bn-styles-list [aria-selected=true]')).toHaveText('1 / a / i')
    await dialog.locator('.bn-styles-show select').selectOption('inUse')
    await expect(dialog.locator('.bn-styles-list [role=option]')).toHaveText([
      'No List',
      '1 / a / i',
    ])
    // No List takes the numbering off the caret's paragraph
    await dialog.locator('.bn-styles-list [role=option]').first().click()
    await expect(dialog.locator('.bn-styles-preview > span')).toHaveCount(0)
    await dialog.locator('button.primary').click()
    await expect.poll(() => markers(page)).toEqual(['1)', 'a)'])
    await saveFromMenu(launched.app)
    await expect
      .poll(
        async () =>
          (await part(docPath, 'word/document.xml')).match(
            new RegExp(`<w:numId w:val="${numId}"/>`, 'g'),
          )?.length,
        { timeout: 30_000 },
      )
      .toBe(2)
    expect((await part(docPath, 'word/styles.xml')).match(/w:styleId="1ai"/g)).toHaveLength(1)
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test("the document's own list style is listed after the built-ins and applies its levels", async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-list-styles-')))
  const docPath = join(dir, 'clauses.docx')
  const zip = await JSZip.loadAsync(await docxWith(BODY))
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W_NS}>` +
      '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/></w:style>' +
      '<w:style w:type="numbering" w:styleId="Clauses" w:customStyle="1"><w:name w:val="Clauses"/>' +
      '<w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style></w:styles>',
  )
  zip.file(
    'word/numbering.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:numbering ${W_NS}>` +
      '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="Clauses"/>' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="Clause %1:"/><w:pPr><w:ind w:left="1080" w:hanging="1080"/></w:pPr></w:lvl>' +
      '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:pPr><w:ind w:left="1800" w:hanging="720"/></w:pPr></w:lvl>' +
      '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
  )
  const rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  zip.file(
    'word/_rels/document.xml.rels',
    (await zip.file('word/_rels/document.xml.rels')!.async('string')).replace(
      '</Relationships>',
      `<Relationship Id="rIdStyles" Type="${rels}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rIdNumbering" Type="${rels}/numbering" Target="numbering.xml"/></Relationships>`,
    ),
  )
  const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml'
  zip.file(
    '[Content_Types].xml',
    (await zip.file('[Content_Types].xml')!.async('string')).replace(
      '</Types>',
      `<Override PartName="/word/styles.xml" ContentType="${wml}.styles+xml"/>` +
        `<Override PartName="/word/numbering.xml" ContentType="${wml}.numbering+xml"/></Types>`,
    ),
  )
  writeFileSync(docPath, await zip.generateAsync({ type: 'nodebuffer' }))
  const { launched, page } = await openDocs(docPath)
  try {
    await selectThrough(page, 'beta')
    const dialog = await openListStyles(launched.app, page)
    const options = dialog.locator('.bn-styles-list [role=option]')
    await expect(options).toHaveText([
      'No List',
      '1 / 1.1 / 1.1.1',
      '1 / a / i',
      'Article / Section',
      'Clauses',
    ])
    await options.nth(4).click()
    const preview = dialog.locator('.bn-styles-preview > span')
    await expect(preview).toHaveCount(2)
    await expect(preview.nth(0)).toContainText('Clause A:')
    await expect(preview.nth(1)).toContainText('A.1')
    await dialog.locator('button.primary').click()
    await expect.poll(() => markers(page)).toEqual(['Clause A:', 'Clause B:'])

    await saveFromMenu(launched.app)
    await expect
      .poll(() => part(docPath, 'word/numbering.xml'), { timeout: 30_000 })
      .toContain('<w:numStyleLink w:val="Clauses"/>')
    const numbering = await part(docPath, 'word/numbering.xml')
    // the style's own definition is untouched and not written a second time
    expect(numbering.match(/<w:styleLink w:val="Clauses"\/>/g)).toHaveLength(1)
    expect(numbering).toContain('<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>')
    const num = /<w:num w:numId="(\d+)"><w:abstractNumId w:val="1"\/><\/w:num>/.exec(numbering)
    expect(num).not.toBeNull()
    expect(
      (await part(docPath, 'word/document.xml')).match(
        new RegExp(`<w:numId w:val="${num![1]}"/>`, 'g'),
      ),
    ).toHaveLength(2)
    expect((await part(docPath, 'word/styles.xml')).match(/w:styleId="Clauses"/g)).toHaveLength(1)
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
