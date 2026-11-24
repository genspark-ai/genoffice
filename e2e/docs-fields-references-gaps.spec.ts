/**
 * Fields and references that follow the headings of the document: Cross-reference lists a
 * heading numbered through its style under Heading and refers to it by its number, and a
 * caption AutoCaption adds takes the chapter number the label's captions already carry.
 * Driven through the dialogs; read on the page and in the saved file.
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
import { docxWith, saveFromMenu, withWordPart } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-fields-references-gaps'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'
const STYLES =
  `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/></w:style>' +
  '</w:styles>'
const NUMBERING =
  `<?xml version="1.0" encoding="UTF-8"?><w:numbering ${W}>` +
  '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
  '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'

const para = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
const chapter = (text: string, numbered = true) =>
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/>' +
  (numbered ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' : '') +
  `</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
const simple = (instr: string, result: string) =>
  `<w:fldSimple w:instr=" ${instr} "><w:r><w:rPr><w:noProof/></w:rPr><w:t>${result}</w:t></w:r></w:fldSimple>`
/** a table caption as Word's Caption dialog writes it with "Include chapter number" */
const wordCaption = (no: string, n: string, text: string) =>
  '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">Table </w:t></w:r>' +
  `${simple('STYLEREF 1 \\s', no)}<w:r><w:noBreakHyphen/></w:r>${simple('SEQ Table \\* ARABIC \\s 1', n)}` +
  `<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p>`

async function chapterDocx(body: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(
    await withWordPart(await docxWith(body + SECT_PR), 'styles', STYLES),
  )
  zip.file('word/numbering.xml', NUMBERING)
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
    ),
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

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

/** the caret at the end (or the start) of the paragraph that starts with `text` */
const caretIn = (page: Page, text: string, atStart = false) =>
  page.evaluate(
    ([text, atStart]) => {
      const { editor } = (window as unknown as EditorWindow).__aidocs
      let to = -1
      editor.state.doc.descendants((node, pos) => {
        if (to < 0 && node.isTextblock && node.textContent.startsWith(text as string))
          to = atStart ? pos + 1 : pos + node.nodeSize - 1
        return to < 0
      })
      if (to < 0) throw new Error(`no paragraph starts with ${text}`)
      editor.chain().focus().setTextSelection(to).run()
    },
    [text, atStart] as const,
  )

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

/** each field of the saved body: its code and the text it shows */
const savedFields = (xml: string) =>
  [
    ...xml.matchAll(
      /<w:fldChar w:fldCharType="begin"[^>]*\/><\/w:r><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:instrText[^>]*>\s*([^<]*?)\s*<\/w:instrText>.*?<w:t[^>]*>([^<]*)<\/w:t>/g,
    ),
  ].map((m) => `${m[1].replace(/_Ref\d+/, '_Ref')} = ${m[2]}`)

test('Cross-reference refers to a numbered heading by its number, and AutoCaption carries the chapter number of its label', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-fields-refs-')))
  const docPath = join(dir, 'chapters.docx')
  writeFileSync(
    docPath,
    await chapterDocx(
      chapter('Alpha') +
        wordCaption('1', '1', 'First') +
        para('Body one') +
        chapter('Plain', false) +
        chapter('Beta') +
        para('Body two') +
        para('By number: ') +
        para('Unnumbered: ') +
        para('By text: ') +
        para('End'),
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
    await expect(page.locator('.doc-page p', { hasText: 'By text:' }).first()).toBeVisible({
      timeout: 30_000,
    })

    // ---- Cross-reference > Heading ----
    const crossRef = page.locator('.crossref-dialog')
    const refer = async (at: string, heading: string, content: string) => {
      await caretIn(page, at)
      await clickMenu(app, ['Insert', 'Cross-reference…'])
      await expect(crossRef).toBeVisible()
      await crossRef.locator('select[name="ref-type"]').selectOption('heading')
      await crossRef.locator('select[name="ref-content"]').selectOption({ label: content })
      await crossRef.locator('.bookmark-row', { hasText: heading }).click()
      await crossRef.locator('button[name="insert"]').click()
      if (await crossRef.isVisible()) await page.keyboard.press('Escape')
      await expect(crossRef).toBeHidden()
    }
    await caretIn(page, 'By number: ')
    await clickMenu(app, ['Insert', 'Cross-reference…'])
    await crossRef.locator('select[name="ref-type"]').selectOption('heading')
    // the numbered headings are listed with the unnumbered one, each with its number
    await expect(crossRef.locator('.bookmark-row')).toHaveText(['1. Alpha', 'Plain', '2. Beta'])
    await expect(crossRef.locator('select[name="ref-content"] option')).toHaveText([
      'Heading text',
      'Page Number',
      'Heading number',
      'Heading number (no context)',
      'Heading number (full context)',
      'Above/below',
    ])
    await page.screenshot({ path: test.info().outputPath('crossref-heading-number.png') })
    await page.keyboard.press('Escape')
    await expect(crossRef).toBeHidden()

    await refer('By number: ', 'Beta', 'Heading number')
    await refer('Unnumbered: ', 'Plain', 'Heading number (no context)')
    await refer('By text: ', 'Beta', 'Heading text')
    await expect(page.locator('.doc-page p', { hasText: 'By number:' })).toHaveText('By number: 2')
    await expect(page.locator('.doc-page p', { hasText: 'Unnumbered:' })).toHaveText(
      'Unnumbered: 0',
    )
    await expect(page.locator('.doc-page p', { hasText: 'By text:' })).toHaveText('By text: Beta')

    // ---- AutoCaption: the label's captions carry a chapter number, so the new one does ----
    await caretIn(page, 'Body two')
    await clickMenu(app, ['Insert', 'Caption…'])
    const caption = page.locator('.caption-dialog')
    await caption.locator('button[name="auto-caption"]').click()
    const auto = page.getByRole('dialog', { name: 'AutoCaption' })
    await auto.locator('input[name="auto-caption-table"]').check()
    await expect(auto.locator('select[name="auto-caption-label"]')).toHaveValue('Table')
    await auto.locator('button[name="ok"]').click()
    await expect(auto).toBeHidden()
    await caption.locator('.btn-ghost').click()
    await expect(caption).toBeHidden()
    await caretIn(page, 'Body two')
    await clickMenu(app, ['Table', 'Insert', 'Table…'])
    await page.locator('.modal .btn-primary').click()
    await expect(page.locator('.doc-page table')).toHaveCount(1)
    const captions = page.locator('.doc-page .doc-protected-field-text')
    await expect(captions).toHaveText(['Table 1\u20111 First', 'Table 2\u20111'])
    await page.screenshot({ path: test.info().outputPath('auto-caption-chapter.png') })

    // a chapter put in front: the reference and the captions follow
    await caretIn(page, 'Alpha', true)
    await page.keyboard.type('Zero')
    await page.keyboard.press('Enter')
    await expect(page.locator('.doc-page p', { hasText: 'By number:' })).toHaveText('By number: 3')

    await saveFromMenu(app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('Zero')
    const xml = await documentXml(docPath)
    expect(xml).toMatch(/<w:t[^>]*>Zero<\/w:t>.*<w:t[^>]*>Alpha<\/w:t>/)
    expect(xml).not.toContain('w:dirty')
    expect(savedFields(xml)).toEqual([
      'STYLEREF 1 \\s = 3',
      'SEQ Table \\* ARABIC \\s 1 = 1',
      'REF _Ref \\r \\h = 3',
      'REF _Ref \\n \\h = 0',
      'REF _Ref \\h = Beta',
    ])
    // the caption the file came with is renumbered too
    expect(xml).toMatch(/STYLEREF 1 \\s "><w:r><w:rPr><w:noProof\/><\/w:rPr><w:t>2<\/w:t>/)
    copyFileSync(docPath, test.info().outputPath('chapters-saved.docx'))
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
