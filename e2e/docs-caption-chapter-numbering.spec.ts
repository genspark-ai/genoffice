/**
 * Caption > Format… (Word's Caption Numbering): a number format, and the chapter number in
 * front of the caption's own, which then starts again in each chapter. Driven through the
 * Caption dialog on a document with numbered Heading 1 chapters; read on the page, in the
 * saved file, and on the page again after the file is reopened.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-caption-chapter-numbering'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'

const para = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
const chapter = (text: string) =>
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
  `<w:r><w:t>${text}</w:t></w:r></w:p>`
const simple = (instr: string, result: string) =>
  `<w:fldSimple w:instr=" ${instr} "><w:r><w:rPr><w:noProof/></w:rPr><w:t>${result}</w:t></w:r></w:fldSimple>`
/** a caption as Word's Caption dialog writes it with "Include chapter number" */
const wordCaption = (no: string, n: string, text: string) =>
  '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">Figure </w:t></w:r>' +
  `${simple('STYLEREF 1 \\s', no)}<w:r><w:noBreakHyphen/></w:r>${simple('SEQ Figure \\* ARABIC \\s 1', n)}` +
  `<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p>`

async function docx(body: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/></w:style>' +
      '</w:styles>',
  )
  zip.file(
    'word/numbering.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:numbering ${W}>` +
      '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>' +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/>' +
      '<w:lvlText w:val="Chapter %1"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}${SECT_PR}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

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

type EditorWindow = { __aidocs: { editor: import('@tiptap/core').Editor } }

/** puts the caret at the end (or the start) of the paragraph starting with `text` */
const caretIn = (page: Page, text: string, atStart = false) =>
  page.evaluate(
    ([text, atStart]) => {
      const { editor } = (window as unknown as EditorWindow).__aidocs
      let range: { from: number; to: number } | null = null
      editor.state.doc.descendants((node, pos) => {
        if (!range && node.isTextblock && node.textContent.startsWith(text as string))
          range = { from: pos + 1, to: pos + node.nodeSize - 1 }
        return !range
      })
      if (!range) throw new Error(`no paragraph starts with ${text}`)
      const { from, to } = range as { from: number; to: number }
      editor
        .chain()
        .focus()
        .setTextSelection(atStart ? from : to)
        .run()
    },
    [text, atStart] as const,
  )

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

/** each caption of the saved body: the instructions of its fields, and the text it shows */
function savedCaptions(xml: string): Array<{ instr: string[]; text: string }> {
  return [...xml.matchAll(/<w:p[ >](?:(?!<\/w:p>)[\s\S])*<\/w:p>/g)]
    .filter((m) => / SEQ /.test(m[0]))
    .map((m) => ({
      instr: [...m[0].matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>|w:instr="([^"]*)"/g)].map(
        (i) => (i[1] ?? i[2]).trim(),
      ),
      text: [...m[0].matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:(noBreakHyphen)\/>/g)]
        .map((t) => (t[2] ? '\u2011' : t[1]))
        .join(''),
    }))
}

/** an open ends by putting the caret at the start: a caret a test places before that is moved away */
const opened = (page: Page) =>
  page.waitForFunction(
    () => (window as unknown as Partial<EditorWindow>).__aidocs?.editor?.view.hasFocus() === true,
    undefined,
    { timeout: 30_000 },
  )

const pagesSettled = (page: Page) =>
  page.waitForFunction(
    () =>
      (window as unknown as { __aidocs: { pagesSettled: () => boolean } }).__aidocs.pagesSettled(),
    undefined,
    { timeout: 30_000 },
  )

async function save(app: ElectronApplication, page: Page, docPath: string, marker: string) {
  await pagesSettled(page)
  await saveFromMenu(app)
  await expect.poll(() => documentXml(docPath), { timeout: 20_000 }).toContain(marker)
  return documentXml(docPath)
}

const captionsOnPage = (page: Page) => page.locator('.doc-page .doc-protected-field-text')

test('captions take the chapter number through Caption Numbering, restart in each chapter, and follow a new figure and a new chapter', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-caption-chapter-')))
  const docPath = join(dir, 'chapters.docx')
  writeFileSync(
    docPath,
    await docx(
      chapter('Alpha') +
        para('Alpha body') +
        para('Alpha more') +
        chapter('Beta') +
        para('Beta body') +
        para('See also:'),
    ),
  )
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Beta body' }).first()).toBeVisible({
      timeout: 30_000,
    })
    await opened(page)
    let caption = page.locator('.caption-dialog')
    const numbering = page.locator('.caption-dialog[aria-label="Caption Numbering"]')
    const openCaption = async (at: string) => {
      await caretIn(page, at)
      await clickMenu(app, ['Insert', 'Caption…'])
      await expect(caption.locator('input[name="caption-text"]')).toBeVisible()
    }

    // ---- Format…: the chapter number goes in front, Heading 1 and a hyphen by default ----
    await openCaption('Alpha body')
    await expect(caption.locator('input[name="caption-text"]')).toHaveAttribute(
      'placeholder',
      /Figure 1 /,
    )
    await caption.locator('button[name="caption-numbering"]').click()
    await expect(numbering).toBeVisible()
    await expect(numbering.locator('select[name="caption-numbering-format"]')).toHaveValue('ARABIC')
    await expect(numbering.locator('input[name="caption-chapter"]')).not.toBeChecked()
    await expect(numbering.locator('select[name="caption-chapter-style"]')).toBeDisabled()
    await expect(numbering.locator('select[name="caption-chapter-sep"]')).toBeDisabled()
    await numbering.locator('input[name="caption-chapter"]').check()
    await expect(numbering.locator('select[name="caption-chapter-style"]')).toHaveValue('1')
    await expect(numbering.locator('select[name="caption-chapter-style"] option')).toHaveCount(9)
    await expect(numbering.locator('select[name="caption-chapter-sep"]')).toHaveValue('hyphen')
    await expect(numbering.locator('select[name="caption-chapter-sep"] option')).toHaveText([
      '-',
      '.',
      ':',
      '\u2014',
      '\u2013',
    ])
    await page.screenshot({ path: test.info().outputPath('caption-numbering.png') })
    await numbering.locator('button[name="ok"]').click()
    await expect(numbering).toBeHidden()
    await expect(caption.locator('input[name="caption-text"]')).toHaveAttribute(
      'placeholder',
      /Figure 1\u20111 /,
    )
    await caption.locator('input[name="caption-text"]').fill('First')
    await page.screenshot({ path: test.info().outputPath('caption-dialog.png') })
    await caption.locator('button[name="ok"]').click()
    await expect(captionsOnPage(page)).toHaveText(['Figure 1\u20111 First'])

    // the label's next caption is numbered the same way without asking again
    await openCaption('Alpha more')
    await caption.locator('button[name="caption-numbering"]').click()
    await expect(numbering.locator('input[name="caption-chapter"]')).toBeChecked()
    await page.keyboard.press('Escape')
    await expect(numbering).toBeHidden()
    await caption.locator('input[name="caption-text"]').fill('Second')
    await caption.locator('button[name="ok"]').click()
    await openCaption('Beta body')
    await expect(caption.locator('input[name="caption-text"]')).toHaveAttribute(
      'placeholder',
      /Figure 2\u20111 /,
    )
    await caption.locator('input[name="caption-text"]').fill('Third')
    await caption.locator('button[name="ok"]').click()
    await expect(captionsOnPage(page)).toHaveText([
      'Figure 1\u20111 First',
      'Figure 1\u20112 Second',
      'Figure 2\u20111 Third',
    ])

    // ---- a figure put between two others moves the ones after it in its chapter only ----
    await openCaption('Alpha body')
    await caption.locator('input[name="caption-text"]').fill('Between')
    await caption.locator('button[name="ok"]').click()
    await expect(captionsOnPage(page)).toHaveText([
      'Figure 1\u20111 Between',
      'Figure 1\u20112 First',
      'Figure 1\u20113 Second',
      'Figure 2\u20111 Third',
    ])

    // ---- a reference to a caption quotes it with its chapter number ----
    const crossRef = page.locator('.crossref-dialog')
    await caretIn(page, 'See also:')
    await clickMenu(app, ['Insert', 'Cross-reference…'])
    await crossRef.locator('select[name="ref-type"]').selectOption('caption:Figure')
    await expect(crossRef.locator('.bookmark-row')).toHaveText([
      'Figure 1\u20111 Between',
      'Figure 1\u20112 First',
      'Figure 1\u20113 Second',
      'Figure 2\u20111 Third',
    ])
    await crossRef.locator('.bookmark-row', { hasText: 'Third' }).click()
    await crossRef.locator('button[name="insert"]').click()
    await expect(page.locator('.doc-page .doc-ref-field')).toHaveText(['Figure 2\u20111 Third'])

    let xml = await save(app, page, docPath, 'STYLEREF')
    expect(xml).not.toContain('w:dirty')
    expect(savedCaptions(xml)).toEqual(
      ['1\u20111 Between', '1\u20112 First', '1\u20113 Second', '2\u20111 Third'].map((shown) => ({
        instr: ['STYLEREF 1 \\s', 'SEQ Figure \\* ARABIC \\s 1'],
        text: `Figure ${shown}`,
      })),
    )
    // the reference keeps the non-breaking hyphen as one
    expect(xml).toMatch(
      / REF _Ref\d+ \\h <\/w:instrText>[\s\S]*?<w:t[^>]*>Figure 2<\/w:t><w:noBreakHyphen\/><w:t[^>]*>1 Third</,
    )

    // ---- reopened: the same numbers; a chapter put in the middle moves the chapters after it ----
    await closeAndSaveVideo(launched)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    caption = page.locator('.caption-dialog')
    await expect(captionsOnPage(page)).toHaveText(
      [
        'Figure 1\u20111 Between',
        'Figure 1\u20112 First',
        'Figure 1\u20113 Second',
        'Figure 2\u20111 Third',
      ],
      { timeout: 30_000 },
    )
    await opened(page)
    // a numbered heading split in two: the first half is a new chapter in front of the old one
    await caretIn(page, 'Beta', true)
    await page.keyboard.type('Inserted')
    await page.keyboard.press('Enter')
    await expect(page.locator('.doc-page [data-chapter]')).toHaveCount(3)
    await expect(page.locator('.doc-page [data-chapter]', { hasText: 'Beta' })).toHaveAttribute(
      'data-chapter',
      '1:3',
    )
    await openCaption('Inserted')
    await caption.locator('input[name="caption-text"]').fill('Fourth')
    await caption.locator('button[name="ok"]').click()
    await expect(captionsOnPage(page)).toHaveText([
      'Figure 1\u20111 Between',
      'Figure 1\u20112 First',
      'Figure 1\u20113 Second',
      'Figure 2\u20111 Fourth',
      'Figure 3\u20111 Third',
    ])
    await expect(page.locator('.doc-page .doc-ref-field')).toHaveText(['Figure 3\u20111 Third'])
    xml = await save(app, page, docPath, 'Fourth')
    expect(xml).not.toContain('w:dirty')
    expect(savedCaptions(xml).map((c) => c.text)).toEqual([
      'Figure 1\u20111 Between',
      'Figure 1\u20112 First',
      'Figure 1\u20113 Second',
      'Figure 2\u20111 Fourth',
      'Figure 3\u20111 Third',
    ])
    await page.screenshot({ path: test.info().outputPath('chapters-page.png') })
    writeFileSync(test.info().outputPath('chapters-saved.docx'), readFileSync(docPath))
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a document from Word with chapter-numbered captions opens with its numbers, and Format… changes every caption of the label', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-caption-chapter-word-')))
  const docPath = join(dir, 'from-word.docx')
  writeFileSync(
    docPath,
    await docx(
      chapter('Alpha') +
        wordCaption('1', '1', 'One') +
        wordCaption('1', '2', 'Two') +
        chapter('Beta') +
        wordCaption('2', '1', 'Three') +
        para('Closing line'),
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(captionsOnPage(page)).toHaveText(
      ['Figure 1\u20111 One', 'Figure 1\u20112 Two', 'Figure 2\u20111 Three'],
      { timeout: 30_000 },
    )
    await opened(page)
    const caption = page.locator('.caption-dialog')
    const numbering = page.locator('.caption-dialog[aria-label="Caption Numbering"]')

    // the dialog opens on what the label's captions have
    await caretIn(page, 'Closing line')
    await clickMenu(app, ['Insert', 'Caption…'])
    await caption.locator('button[name="caption-numbering"]').click()
    await expect(numbering.locator('input[name="caption-chapter"]')).toBeChecked()
    await expect(numbering.locator('select[name="caption-chapter-style"]')).toHaveValue('1')
    await expect(numbering.locator('select[name="caption-chapter-sep"]')).toHaveValue('hyphen')
    await numbering.locator('select[name="caption-numbering-format"]').selectOption('ROMAN')
    await numbering.locator('select[name="caption-chapter-sep"]').selectOption('colon')
    await numbering.locator('button[name="ok"]').click()
    await expect(caption.locator('select[name="caption-format"]')).toHaveValue('ROMAN')
    await expect(caption.locator('input[name="caption-text"]')).toHaveAttribute(
      'placeholder',
      /Figure 2:II /,
    )
    // Cancel leaves the document as it was
    await page.keyboard.press('Escape')
    await expect(captionsOnPage(page)).toHaveText([
      'Figure 1\u20111 One',
      'Figure 1\u20112 Two',
      'Figure 2\u20111 Three',
    ])

    await clickMenu(app, ['Insert', 'Caption…'])
    await caption.locator('button[name="caption-numbering"]').click()
    await numbering.locator('select[name="caption-numbering-format"]').selectOption('ROMAN')
    await numbering.locator('select[name="caption-chapter-sep"]').selectOption('colon')
    await numbering.locator('button[name="ok"]').click()
    await caption.locator('input[name="caption-text"]').fill('Four')
    await caption.locator('button[name="ok"]').click()
    await expect(captionsOnPage(page)).toHaveText([
      'Figure 1:I One',
      'Figure 1:II Two',
      'Figure 2:I Three',
      'Figure 2:II Four',
    ])
    let xml = await save(app, page, docPath, 'Four')
    expect(xml).not.toContain('w:dirty')
    expect(savedCaptions(xml)).toEqual(
      ['1:I One', '1:II Two', '2:I Three', '2:II Four'].map((shown) => ({
        instr: ['STYLEREF 1 \\s', 'SEQ Figure \\* ROMAN \\s 1'],
        text: `Figure ${shown}`,
      })),
    )

    // without the chapter number the label counts through the document again
    await caretIn(page, 'Closing line')
    await clickMenu(app, ['Insert', 'Caption…'])
    await caption.locator('button[name="caption-numbering"]').click()
    await expect(numbering.locator('select[name="caption-chapter-sep"]')).toHaveValue('colon')
    // unticked and ticked again, the separator is still the one the captions have
    await numbering.locator('input[name="caption-chapter"]').uncheck()
    await numbering.locator('input[name="caption-chapter"]').check()
    await expect(numbering.locator('select[name="caption-chapter-sep"]')).toHaveValue('colon')
    await numbering.locator('input[name="caption-chapter"]').uncheck()
    await numbering.locator('select[name="caption-numbering-format"]').selectOption('alphabetic')
    await numbering.locator('button[name="ok"]').click()
    await caption.locator('input[name="caption-text"]').fill('Five')
    await caption.locator('button[name="ok"]').click()
    await expect(captionsOnPage(page)).toHaveText([
      'Figure a One',
      'Figure b Two',
      'Figure c Three',
      'Figure d Five',
      'Figure e Four',
    ])
    // one undo takes the caption and the label's new numbering back
    await expect
      .poll(() =>
        page.evaluate(() => {
          const { editor } = (window as unknown as EditorWindow).__aidocs
          if (!editor.isFocused) editor.commands.focus()
          return editor.isFocused
        }),
      )
      .toBe(true)
    await page.keyboard.press('ControlOrMeta+z')
    await expect(captionsOnPage(page)).toHaveText([
      'Figure 1:I One',
      'Figure 1:II Two',
      'Figure 2:I Three',
      'Figure 2:II Four',
    ])
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect(captionsOnPage(page)).toHaveCount(5)
    xml = await save(app, page, docPath, 'Five')
    expect(xml).not.toContain('w:dirty')
    expect(xml).not.toContain('STYLEREF')
    expect(savedCaptions(xml)).toEqual(
      ['a One', 'b Two', 'c Three', 'd Five', 'e Four'].map((shown) => ({
        instr: ['SEQ Figure \\* alphabetic'],
        text: `Figure ${shown}`,
      })),
    )
    writeFileSync(test.info().outputPath('from-word-saved.docx'), readFileSync(docPath))
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})
