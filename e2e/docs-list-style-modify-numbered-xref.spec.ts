/**
 * Three dialog entries that used to be hidden:
 *  - Bullets and Numbering > List Styles > Modify… rewrites the levels of a list
 *    style where it is defined, and renames a user-made one;
 *  - Cross-reference > Numbered item inserts REF \r / \n / \w fields whose
 *    result is the paragraph's number and follows the numbering;
 *  - the references a file already holds show their numbers and follow too.
 * Each is checked on the page, in the saved parts and after a reopen.
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

const VIDEO_DIR = 'docs-list-style-modify-numbered-xref'

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
        doc: {
          descendants(
            f: (
              n: { textContent: string; isTextblock: boolean; nodeSize: number },
              pos: number,
            ) => boolean,
          ): void
        }
      }
      chain(): { focus(): { setTextSelection(r: object | number): { run(): boolean } } }
    }
  }
}

/** selects the paragraphs from the one starting with `first` through `last`, or without `last` puts the caret at that end of the first */
const select = (page: Page, first: string, last: string | null, edge: 'start' | 'end' = 'end') =>
  page.evaluate(
    ([first, last, edge]) => {
      const { editor } = (window as never as EditorWindow).__aidocs
      let from = -1
      let to = -1
      editor.state.doc.descendants((node, pos) => {
        if (!node.isTextblock) return true
        if (node.textContent.startsWith(first)) {
          from = edge === 'start' || last !== null ? pos + 1 : pos + node.nodeSize - 1
          if (last === null) to = from
        }
        if (last !== null && node.textContent === last) to = pos + 1 + last.length
        return false
      })
      if (from < 0 || to < 0) throw new Error(`no paragraph reads ${first}`)
      editor.chain().focus().setTextSelection({ from, to }).run()
    },
    [first, last, edge] as const,
  )

/** the caret at the end of the paragraph starting with `first`, once the keys go to the document */
async function typeAt(page: Page, first: string) {
  await select(page, first, null)
  await expect
    .poll(() => page.evaluate(() => !!document.activeElement?.closest('.ProseMirror')))
    .toBe(true)
}

async function openDocs(docPath: string, shows: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page').first()).toContainText(shows, { timeout: 30_000 })
  return { launched, page }
}

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

const markers = (page: Page) =>
  page
    .locator('.doc-page .doc-li')
    .evaluateAll((items) => items.map((el) => el.getAttribute('data-marker')))

async function openListStyles(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Format', 'Bullets and Numbering…'])
  const dialog = page.locator('.bn-dialog')
  await dialog.getByRole('tab', { name: 'List Styles' }).click()
  return dialog
}

test('List Styles > Modify… changes the levels of the style for every list that follows it, and renames a style of the user', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-list-style-modify-')))
  const docPath = join(dir, 'plain.docx')
  writeFileSync(docPath, await docxWith(['alpha', 'beta', 'gamma', 'delta'].map(para).join('')))
  const options = (dialog: ReturnType<Page['locator']>) =>
    dialog.locator('.bn-styles-list [role=option]')
  const abstractOf = (numbering: string, styleId: string) =>
    numbering
      .split('<w:abstractNum ')
      .find((chunk) => chunk.includes(`<w:styleLink w:val="${styleId}"/>`)) ?? ''

  const first = await openDocs(docPath, 'alpha')
  try {
    const { launched, page } = first
    await select(page, 'alpha', 'gamma')
    let dialog = await openListStyles(launched.app, page)
    const modify = dialog.locator('button[name="list-style-modify"]')
    await expect(modify).toBeDisabled()
    await options(dialog).nth(0).click()
    // No List has no levels
    await expect(modify).toBeDisabled()
    await options(dialog).nth(2).click()
    await expect(modify).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('list-styles-modify-button.png') })
    await dialog.locator('button.primary').click()
    await expect.poll(() => markers(page)).toEqual(['1)', '2)', '3)'])
    // the style and its list are in the file before it is modified
    await saveFromMenu(launched.app)
    await expect
      .poll(() => part(docPath, 'word/numbering.xml'), { timeout: 30_000 })
      .toContain('<w:styleLink w:val="1ai"/>')

    await select(page, 'beta', null)
    dialog = await openListStyles(launched.app, page)
    await expect(dialog.locator('.bn-styles-list [aria-selected=true]')).toHaveText('1 / a / i')
    await modify.click()
    await expect(dialog).toBeHidden()
    const define = page.locator('.modal.list-define-modal')
    await expect(define.locator('h2')).toHaveText('Modify Style')
    const name = define.locator('input[name="list-style-name"]')
    // one of Word's own list styles keeps its name
    await expect(name).toHaveValue('1 / a / i')
    await expect(name).toBeDisabled()
    // the editor starts from the definition of the style
    await expect(define.getByLabel('Number format')).toHaveValue('%1)')
    await define.getByLabel('Number format').fill('Item %1:')
    await define.getByLabel('Start at').fill('4')
    await page.screenshot({ path: test.info().outputPath('modify-list-style.png') })
    await define.locator('button.primary').click()
    await expect(define).toBeHidden()

    // back on the List Styles tab, the style picked and previewed as modified
    await expect(dialog.locator('.bn-styles-list [aria-selected=true]')).toHaveText('1 / a / i')
    await expect(dialog.locator('.bn-styles-preview > span').first()).toContainText('Item 4:')
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect.poll(() => markers(page)).toEqual(['Item 4:', 'Item 5:', 'Item 6:'])
    await page.screenshot({ path: test.info().outputPath('list-style-modified-page.png') })

    // a style of the user's: Modify renames it
    dialog = await openListStyles(launched.app, page)
    await dialog.getByRole('button', { name: 'Add' }).click()
    await name.fill('House Steps')
    await define.locator('button.primary').click()
    await expect(dialog.locator('.bn-styles-list [aria-selected=true]')).toHaveText('House Steps')
    await modify.click()
    await expect(name).toBeEnabled()
    await expect(name).toHaveValue('House Steps')
    // a name another style has is refused
    await name.fill('no list')
    await expect(define.locator('button.primary')).toBeDisabled()
    await name.fill('Site Steps')
    await define.getByLabel('Number format').fill('Step %1.')
    await define.locator('button.primary').click()
    await expect(options(dialog)).toHaveText([
      'No List',
      '1 / 1.1 / 1.1.1',
      '1 / a / i',
      'Article / Section',
      'Site Steps',
    ])
    await expect(dialog.locator('.bn-styles-preview > span').first()).toContainText('Step 1.')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await saveFromMenu(launched.app)
    await expect
      .poll(() => part(docPath, 'word/numbering.xml'), { timeout: 30_000 })
      .toContain('<w:lvlText w:val="Item %1:"/>')
    const numbering = await part(docPath, 'word/numbering.xml')
    const own = abstractOf(numbering, '1ai')
    expect(own).toMatch(/<w:lvl w:ilvl="0"[^>]*><w:start w:val="4"\/>/)
    expect(own).toContain('<w:lvlText w:val="Item %1:"/>')
    // the other levels are the style's as they were
    expect(own).toContain('<w:lvlText w:val="%2)"/>')
    expect(numbering.match(/<w:styleLink w:val="1ai"\/>/g)).toHaveLength(1)
    expect(numbering.match(/<w:numStyleLink w:val="1ai"\/>/g)).toHaveLength(1)
    expect(abstractOf(numbering, 'HouseSteps')).toContain('<w:lvlText w:val="Step %1."/>')
    expect(await part(docPath, 'word/styles.xml')).toContain(
      '<w:style w:type="numbering" w:customStyle="1" w:styleId="HouseSteps"><w:name w:val="Site Steps"/>',
    )
    copyFileSync(docPath, test.info().outputPath('list-style-modified.docx'))
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  const second = await openDocs(docPath, 'alpha')
  try {
    const { launched, page } = second
    await expect.poll(() => markers(page)).toEqual(['Item 4:', 'Item 5:', 'Item 6:'])
    await select(page, 'beta', null)
    const dialog = await openListStyles(launched.app, page)
    await expect(options(dialog)).toHaveText([
      'No List',
      '1 / 1.1 / 1.1.1',
      '1 / a / i',
      'Article / Section',
      'Site Steps',
    ])
    await expect(dialog.locator('.bn-styles-preview > span').first()).toContainText('Item 4:')
    // modified again from the file's definition: the level set before is what the editor shows
    await dialog.locator('button[name="list-style-modify"]').click()
    const define = page.locator('.modal.list-define-modal')
    await expect(define.getByLabel('Number format')).toHaveValue('Item %1:')
    await expect(define.getByLabel('Start at')).toHaveValue('4')
    await define.getByRole('button', { name: 'Cancel' }).click()
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

const NUMBERING =
  `<?xml version="1.0" encoding="UTF-8"?><w:numbering ${NS}>` +
  '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'

const item = (text: string, ilvl: number, bookmark?: [id: number, name: string]) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
  (bookmark ? `<w:bookmarkStart w:id="${bookmark[0]}" w:name="${bookmark[1]}"/>` : '') +
  `<w:r><w:t>${text}</w:t></w:r>` +
  (bookmark ? `<w:bookmarkEnd w:id="${bookmark[0]}"/>` : '') +
  '</w:p>'

const text = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`

const ref = (instr: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:t>${result}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'

/** a contract as Word saves one: numbered clauses, and references to their numbers with the results Word computed */
async function contractDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
  )
  zip.file('word/numbering.xml', NUMBERING)
  const body =
    item('Scope', 0) +
    item('Terms', 1, [1, '_Ref100']) +
    item('Fees', 1) +
    item('Notices', 0, [2, '_Ref200']) +
    `<w:p>${text('See clause ')}${ref('REF _Ref100 \\w \\h', '1.(a)')}${text(' and ')}` +
    `${ref('REF _Ref200 \\r \\h', '2')}${text('; formerly ')}${ref('REF _Ref200 \\n \\h', '9')}</w:p>` +
    `<w:p>${text('Also see ')}</w:p>`
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** each REF field of the saved body as "instruction => result" */
const fieldsOf = (xml: string): string[] =>
  [
    ...xml.matchAll(
      /<w:instrText[^>]*>([^<]*)<\/w:instrText>[\s\S]*?w:fldCharType="separate"([\s\S]*?)w:fldCharType="end"/g,
    ),
  ].map(
    (m) =>
      `${m[1]
        .trim()
        .replace(/\s+/g, ' ')
        .replace(/_Ref\d+/, '_Ref')} => ` +
      [...m[2].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(''),
  )

test('Cross-reference > Numbered item: the references show the paragraph number, follow the numbering, and are saved with their result', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-numbered-xref-')))
  const docPath = join(dir, 'contract.docx')
  writeFileSync(docPath, await contractDocx())
  const refs = (page: Page) => page.locator('.doc-page .doc-ref-field')

  const first = await openDocs(docPath, 'Notices')
  try {
    const { launched, page } = first
    // the file's own references show the numbers the file holds
    await expect(refs(page)).toHaveText(['1.(a)', '2', '9'])
    await expect.poll(() => markers(page)).toEqual(['1.', '(a)', '(b)', '2.'])

    const crossRef = page.locator('.crossref-dialog')
    const openCrossRef = async () => {
      await select(page, 'Also see ', null)
      await clickMenu(launched.app, ['Insert', 'Cross-reference…'])
      await expect(crossRef).toBeVisible()
      await crossRef.locator('select[name="ref-type"]').selectOption('numbered')
    }
    const separate = crossRef.locator('input[name="ref-separator"]')
    const separator = crossRef.locator('input[name="ref-separator-text"]')

    await openCrossRef()
    await expect(crossRef.locator('select[name="ref-type"] option').first()).toHaveText(
      'Numbered item',
    )
    await expect(crossRef.locator('select[name="ref-content"] option')).toHaveText([
      'Page Number',
      'Paragraph number',
      'Paragraph number (no context)',
      'Paragraph number (full context)',
      'Paragraph text',
      'Above/below',
    ])
    await expect(crossRef.locator('select[name="ref-content"]')).toHaveValue('paraNumber')
    await expect(crossRef.locator('.bookmark-row')).toHaveText([
      '1. Scope',
      '(a) Terms',
      '(b) Fees',
      '2. Notices',
    ])
    await expect(crossRef.locator('input[name="ref-above-below"]')).toBeEnabled()
    // the separator belongs to the full-context number
    await expect(separate).toBeDisabled()
    await expect(separator).toBeDisabled()
    await crossRef.locator('.bookmark-row', { hasText: 'Fees' }).click()
    await page.screenshot({ path: test.info().outputPath('crossref-numbered-item.png') })
    await crossRef.locator('button[name="insert"]').click()
    await expect(refs(page)).toHaveText(['1.(a)', '2', '9', '1.(b)'])

    await openCrossRef()
    await crossRef.locator('select[name="ref-content"]').selectOption('paraNumberFullContext')
    await expect(separate).toBeEnabled()
    await expect(separator).toBeDisabled()
    await separate.check()
    await separator.fill('-')
    await crossRef.locator('input[name="ref-above-below"]').check()
    await crossRef.locator('.bookmark-row', { hasText: 'Fees' }).click()
    await page.screenshot({ path: test.info().outputPath('crossref-separator.png') })
    await crossRef.locator('button[name="insert"]').click()
    await expect(refs(page)).toHaveText(['1.(a)', '2', '9', '1.(b)', '1-(b) above'])

    await openCrossRef()
    await crossRef.locator('select[name="ref-content"]').selectOption('paraNumberNoContext')
    await crossRef.locator('.bookmark-row', { hasText: 'Notices' }).click()
    await crossRef.locator('button[name="insert"]').click()
    await expect(refs(page)).toHaveText(['1.(a)', '2', '9', '1.(b)', '1-(b) above', '2'])

    // a clause added before them renumbers the list, and the references with it;
    // the one that did not agree with the numbering in the file stays as the file has it
    await expect(crossRef).toBeHidden()
    await typeAt(page, 'Scope')
    await page.keyboard.press('Enter')
    await page.keyboard.type('Parties')
    await expect.poll(() => markers(page)).toEqual(['1.', '2.', '(a)', '(b)', '3.'])
    await expect(refs(page)).toHaveText(['2.(a)', '3', '9', '2.(b)', '2-(b) above', '3'])
    await page.screenshot({ path: test.info().outputPath('crossref-renumbered.png') })

    await saveFromMenu(launched.app)
    await expect
      .poll(async () => fieldsOf(await part(docPath, 'word/document.xml')).length, {
        timeout: 30_000,
      })
      .toBe(6)
    const xml = await part(docPath, 'word/document.xml')
    expect(fieldsOf(xml)).toEqual([
      'REF _Ref \\w \\h => 2.(a)',
      'REF _Ref \\r \\h => 3',
      'REF _Ref \\n \\h => 9',
      'REF _Ref \\r \\h => 2.(b)',
      'REF _Ref \\w \\p \\d "-" \\h => 2-(b) above',
      'REF _Ref \\n \\h => 3',
    ])
    expect(xml).not.toMatch(/w:dirty/)
    // the paragraph a new reference points at got one hidden bookmark, used by both
    const fees = /<w:p>(?:(?!<\/w:p>)[\s\S])*Fees[\s\S]*?<\/w:p>/.exec(xml)![0]
    const made = /<w:bookmarkStart w:id="\d+" w:name="(_Ref\d+)"\/>/.exec(fees)![1]
    expect(xml.split(`REF ${made} `)).toHaveLength(3)
    expect(xml.match(/<w:bookmarkStart /g)).toHaveLength(3)
    copyFileSync(docPath, test.info().outputPath('numbered-xref.docx'))
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  const second = await openDocs(docPath, 'Notices')
  try {
    const { launched, page } = second
    await expect(refs(page)).toHaveText(['2.(a)', '3', '9', '2.(b)', '2-(b) above', '3'])
    // the added clause taken out again: the references of the reopened file follow
    await typeAt(page, 'Parties')
    // its text, then the emptied item, then the paragraph it left
    for (let i = 0; i < 'Parties'.length + 2; i++) await page.keyboard.press('Backspace')
    await expect.poll(() => markers(page)).toEqual(['1.', '(a)', '(b)', '2.'])
    await expect(refs(page)).toHaveText(['1.(a)', '2', '9', '1.(b)', '1-(b) above', '2'])
    await saveFromMenu(launched.app)
    await expect
      .poll(async () => fieldsOf(await part(docPath, 'word/document.xml'))[0], { timeout: 30_000 })
      .toBe('REF _Ref \\w \\h => 1.(a)')
    const xml = await part(docPath, 'word/document.xml')
    expect(fieldsOf(xml).map((f) => f.split(' => ')[1])).toEqual([
      '1.(a)',
      '2',
      '9',
      '1.(b)',
      '1-(b) above',
      '2',
    ])
    expect(xml).not.toMatch(/w:dirty/)
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
