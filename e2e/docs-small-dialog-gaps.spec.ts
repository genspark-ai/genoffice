/**
 * Controls Word has in its smaller dialogs: Field's Preserve formatting, the
 * Language dialog's Do not check spelling or grammar, Date and Time's Language
 * and full-width characters, Drop Cap's Font and Distance from text, Phonetic
 * Guide's Font, Size and distance, and the Styles pane's Modify….
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-small-dialog-gaps'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'
const LONG =
  'pon a time there was a paragraph long enough to run over several lines of the page, ' +
  'so that a letter dropped at its start has text to sit beside on every one of the lines ' +
  'it spans and a little more text after that to show where the paragraph goes on below it.'

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

/** selects `text` inside the paragraph that holds it, or with `caret` puts the caret right after it */
const selectText = (page: Page, text: string, caret = false) =>
  page.evaluate(
    ([text, caret]) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: {
                doc: {
                  descendants(
                    f: (node: { isText: boolean; text?: string }, pos: number) => void,
                  ): void
                }
              }
              chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } }
            }
          }
        }
      ).__aidocs
      let from = -1
      editor.state.doc.descendants((node, pos) => {
        const at = node.isText ? (node.text ?? '').indexOf(text as string) : -1
        if (from < 0 && at >= 0) from = pos + at
      })
      if (from < 0) throw new Error(`no text "${text}"`)
      const to = from + (text as string).length
      editor
        .chain()
        .focus()
        .setTextSelection({ from: caret ? to : from, to })
        .run()
    },
    [text, caret] as const,
  )

const caretAt = (page: Page, text: string) => selectText(page, text, true)

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''
const documentXml = (docPath: string) => part(docPath, 'word/document.xml')

const paragraphXml = (xml: string, text: string) =>
  (xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find((p) => p.includes(text)) ?? ''

test('Field keeps formatting on update, Date and Time follows a language and writes full width, Language turns proofing off', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-small-dialogs-')))
  const docPath = join(dir, 'fields.docx')
  writeFileSync(
    docPath,
    await docxWith(
      ['Year: ', 'Pages: ', 'Updated: ', 'Printed: ', 'Wide: ']
        .map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`)
        .join('') +
        para('Spelled mispeled here') +
        '<w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r>' +
        '<w:r><w:rPr><w:noProof/></w:rPr><w:t>skipme</w:t></w:r></w:p>' +
        SECT_PR,
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
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(paragraph('Year:')).toBeVisible({ timeout: 30_000 })
    const year = String(new Date().getFullYear())

    // Field: a new field keeps its formatting on update unless the box is cleared
    const field = page.locator('.field-dialog')
    const preserve = field.locator('input[name="preserve"]')
    await caretAt(page, 'Year: ')
    await clickMenu(app, ['Insert', 'Field…'])
    await expect(preserve).toBeChecked()
    await expect(field.getByRole('button', { name: 'Options…' })).toBeDisabled()
    await field.getByLabel('Field code').fill('DATE \\@ "yyyy"')
    await page.screenshot({ path: test.info().outputPath('field-dialog.png') })
    await field.locator('button.btn-primary').click()
    await expect(paragraph('Year:')).toContainText(`Year: ${year}`)
    await caretAt(page, 'Pages: ')
    await clickMenu(app, ['Insert', 'Field…'])
    // the switch typed into the code ticks the box; clearing the box takes it out of the code
    await field.getByLabel('Field code').fill('NUMPAGES \\* MERGEFORMAT')
    await expect(preserve).toBeChecked()
    await preserve.uncheck()
    await expect(field.getByLabel('Field code')).toHaveValue('NUMPAGES')
    await field.locator('button.btn-primary').click()
    await expect(paragraph('Pages:')).toContainText('Pages: 1')

    // Date and Time: the formats and the names follow the language picked
    const date = page.locator('.date-dialog')
    const formats = date.locator('.date-dialog-list button')
    await caretAt(page, 'Updated: ')
    await clickMenu(app, ['Insert', 'Date and Time…'])
    await expect(date.locator('select[name="language"]')).toHaveValue('en-US')
    const now = new Date()
    const english = now.toLocaleDateString('en-US', { month: 'long' })
    const german = now.toLocaleDateString('de-DE', { month: 'long' })
    const germanDay = now.toLocaleDateString('de-DE', { weekday: 'long' })
    await expect(formats.nth(1)).toContainText(english)
    await date.locator('select[name="language"]').selectOption('de-DE')
    await expect(formats.nth(0)).toHaveText(/^\d\d\.\d\d\.\d\d$/)
    await expect(formats.nth(1)).toHaveText(new RegExp(`^${germanDay}, \\d+\\. ${german} ${year}$`))
    await formats.nth(1).click()
    await date.getByLabel('Update automatically').check()
    await page.screenshot({ path: test.info().outputPath('date-language.png') })
    await date.locator('button.primary').click()
    await expect(paragraph('Updated:')).toContainText(`${germanDay}, `)
    // as text, in the language picked
    await caretAt(page, 'Printed: ')
    await clickMenu(app, ['Insert', 'Date and Time…'])
    await date.locator('select[name="language"]').selectOption('fr-FR')
    const french = now.toLocaleDateString('fr-FR', { month: 'long' })
    await formats.filter({ hasText: french }).first().click()
    await date.locator('button.primary').click()
    await expect(paragraph('Printed:')).toContainText(french)
    // full width: the digits of the date as double-byte characters, in the list and on the page
    await caretAt(page, 'Wide: ')
    await clickMenu(app, ['Insert', 'Date and Time…'])
    const fullWidth = date.getByLabel('Use full width characters')
    await expect(fullWidth).not.toBeChecked()
    await date.locator('select[name="language"]').selectOption('ja-JP')
    await fullWidth.check()
    await expect(formats.nth(0)).toHaveText(/^[\uff10-\uff19\uff0f]+$/)
    await date.getByLabel('Update automatically').check()
    await page.screenshot({ path: test.info().outputPath('date-full-width.png') })
    await date.locator('button.primary').click()
    await expect(paragraph('Wide:')).toContainText(/[\uff10-\uff19]{4}/)
    // F9 recomputes the fields: each keeps its language and width
    await page.keyboard.press('Meta+A')
    await page.keyboard.press('F9')
    await expect(paragraph('Updated:')).toContainText(`${germanDay}, `)
    await expect(paragraph('Wide:')).toContainText(/[\uff10-\uff19]{4}/)

    // Language: text marked not to be checked
    const language = page.locator('.language-dialog')
    const noProof = language.locator('input[name="noProof"]')
    await selectText(page, 'mispeled')
    await clickMenu(app, ['Tools', 'Language…'])
    await expect(noProof).not.toBeChecked()
    await expect(language.getByLabel('Detect language automatically')).not.toBeChecked()
    await noProof.check()
    await page.screenshot({ path: test.info().outputPath('language-no-proof.png') })
    await language.locator('button.primary').click()
    await expect(
      page.locator('.doc-page span[spellcheck="false"]', { hasText: 'mispeled' }),
    ).toHaveCount(1)
    await clickMenu(app, ['Tools', 'Language…'])
    await expect(noProof).toBeChecked()
    await page.keyboard.press('Escape')
    await expect(language).toBeHidden()
    // text the file already marks opens marked, and the mark comes off
    await expect(
      page.locator('.doc-page span[spellcheck="false"]', { hasText: 'skipme' }),
    ).toHaveCount(1)
    await selectText(page, 'skipme')
    await clickMenu(app, ['Tools', 'Language…'])
    await expect(noProof).toBeChecked()
    await noProof.uncheck()
    await language.locator('button.primary').click()
    await expect(
      page.locator('.doc-page span[spellcheck="false"]', { hasText: 'skipme' }),
    ).toHaveCount(0)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const updated = paragraphXml(xml, 'Updated: ')
          return [
            /DATE \\@ "yyyy" \\\* MERGEFORMAT/.test(paragraphXml(xml, 'Year: ')),
            /<w:instrText[^>]*>\s*NUMPAGES\s*</.test(paragraphXml(xml, 'Pages: ')),
            /DATE \\@ "dddd, d\. MMMM yyyy"/.test(updated),
            /<w:lang w:val="de-DE"\/>/.test(updated),
            paragraphXml(xml, 'Printed: ').includes(french),
            /DATE \\@ "[^"]+" \\\* DBCHAR/.test(paragraphXml(xml, 'Wide: ')),
            /<w:lang w:eastAsia="ja-JP"\/>/.test(paragraphXml(xml, 'Wide: ')),
            /<w:noProof\/>[\s\S]*mispeled/.test(paragraphXml(xml, 'mispeled')),
            paragraphXml(xml, 'skipme').includes('w:noProof'),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true, true, true, true, false])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Drop Cap takes a font and a distance from the text, Phonetic Guide a font, size and distance', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-small-dialogs-')))
  const docPath = join(dir, 'caps.docx')
  const ruby =
    '<w:ruby><w:rubyPr><w:rubyAlign w:val="center"/><w:hps w:val="16"/><w:hpsRaise w:val="30"/>' +
    '<w:hpsBaseText w:val="24"/><w:lid w:val="ja-JP"/></w:rubyPr>' +
    '<w:rt><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="16"/></w:rPr>' +
    '<w:t>ki</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>tree</w:t></w:r></w:rubyBase></w:ruby>'
  writeFileSync(
    docPath,
    await docxWith(
      para(`O${LONG}`) +
        '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="2" w:hSpace="288" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/></w:pPr>' +
        `<w:r><w:t>U${LONG}</w:t></w:r></w:p>` +
        para('Before noon') +
        `<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:r>${ruby}</w:r><w:r><w:t xml:space="preserve"> grows</w:t></w:r></w:p>` +
        SECT_PR,
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
    const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    const first = paragraph('Opon a time')
    const second = paragraph('Upon a time')
    await expect(first).toBeVisible({ timeout: 30_000 })
    /** the space the dropped letter keeps on its text side, in px */
    const gapOf = (p: Locator) =>
      p.evaluate((el) => parseFloat(getComputedStyle(el, '::first-letter').marginRight))
    /** where the text beside the dropped letter starts, from the paragraph's left edge */
    const secondLineIndent = (p: Locator) =>
      p.evaluate((el) => {
        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT)
        let node = walker.nextNode()
        while (node && !(node.textContent ?? '').includes('pon a time')) node = walker.nextNode()
        if (!node) throw new Error('no text beside the letter')
        const at = (node.textContent ?? '').indexOf('pon a time')
        const range = document.createRange()
        range.setStart(node, at)
        range.setEnd(node, at + 1)
        return range.getBoundingClientRect().left - el.getBoundingClientRect().left
      })

    // a drop cap the file sets 0.2 inch from its text opens with that distance, on the page and in the dialog
    const dialog = page.locator('.dropcap-dialog')
    const distance = dialog.getByLabel('Distance from text')
    await expect.poll(() => gapOf(second)).toBeGreaterThan(19)
    await caretAt(page, 'Upon')
    await clickMenu(app, ['Format', 'Drop Cap…'])
    await expect(dialog.getByLabel('Dropped')).toBeChecked()
    await expect(distance).toHaveValue(/^(0\.2"|0\.51 cm)$/)
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()

    // a new one: dropped, then moved away from the text and set in another font
    await caretAt(page, 'Opon')
    await clickMenu(app, ['Format', 'Drop Cap…'])
    await expect(distance).toBeDisabled()
    await expect(dialog.locator('select[name="font"]')).toBeDisabled()
    await dialog.getByLabel('Dropped').check()
    await dialog.locator('button.primary').click()
    await expect(first).toHaveClass(/has-drop-cap/)
    const tight = await gapOf(first)
    const tightIndent = await secondLineIndent(first)
    await page.screenshot({ path: test.info().outputPath('dropcap-plain.png') })
    await clickMenu(app, ['Format', 'Drop Cap…'])
    await distance.fill('18 pt')
    await dialog.locator('select[name="font"]').selectOption('Georgia')
    await page.screenshot({ path: test.info().outputPath('dropcap-dialog.png') })
    await dialog.locator('button.primary').click()
    // 18 pt is 24 px more room between the letter and the text beside it
    await expect.poll(() => gapOf(first)).toBeCloseTo(tight + 24, 0)
    await expect
      .poll(() =>
        first.evaluate((el) => (el.firstElementChild as HTMLElement | null)?.style.fontFamily),
      )
      .toContain('Georgia')
    await expect(first.locator('span').first()).toHaveText('O')
    expect(await secondLineIndent(first)).toBeGreaterThan(tightIndent + 15)
    await page.screenshot({ path: test.info().outputPath('dropcap-font-distance.png') })
    await clickMenu(app, ['Format', 'Drop Cap…'])
    await expect(dialog.locator('select[name="font"]')).toHaveValue('Georgia')
    await expect(distance).toHaveValue(/^(0\.25"|0\.64 cm|0\.63 cm)$/)
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()

    // Phonetic Guide: the guide in the file opens with its reading's font, size and distance
    const phonetic = page.locator('.phonetic-guide-dialog')
    const guides = page.locator('.doc-page ruby.doc-ruby')
    const rtStyle = (guide: Locator) =>
      guide.locator('rt').evaluate((el) => {
        const cs = getComputedStyle(el)
        return [cs.fontSize, cs.bottom, cs.fontFamily.includes('Georgia')]
      })
    await expect(guides).toHaveCount(1)
    // 8 pt, raised 4 pt beyond where a 12 pt base puts it
    await expect.poll(() => rtStyle(guides.first())).toEqual(['10.6667px', '5.33333px', true])
    await guides.first().click()
    await clickMenu(app, ['Format', 'Asian Layout', 'Phonetic Guide…'])
    await expect(phonetic.locator('select[name="font"]')).toHaveValue('Georgia')
    await expect(phonetic.locator('input[name="size"]')).toHaveValue('8')
    await expect(phonetic.locator('input[name="offset"]')).toHaveValue('4')
    // brought back down onto the base text
    const raised = await guides.first().locator('rt').boundingBox()
    await phonetic.locator('input[name="offset"]').fill('0')
    await phonetic.locator('button.primary').click()
    await expect.poll(async () => (await rtStyle(guides.first()))[1]).toBe('auto')
    const lowered = await guides.first().locator('rt').boundingBox()
    expect(lowered!.y - raised!.y).toBeCloseTo(5.33, 0)

    // a new guide: reading in Arial, 9 pt, 3 pt above the text
    await selectText(page, 'Before')
    await clickMenu(app, ['Format', 'Asian Layout', 'Phonetic Guide…'])
    await phonetic.getByLabel('Ruby text', { exact: true }).fill('early')
    await phonetic.locator('select[name="font"]').selectOption('Arial')
    await phonetic.locator('input[name="size"]').fill('9')
    await phonetic.locator('input[name="offset"]').fill('3')
    await page.screenshot({ path: test.info().outputPath('phonetic-dialog.png') })
    await phonetic.locator('button.primary').click()
    await expect(guides).toHaveCount(2)
    const made = guides.first()
    await expect
      .poll(() =>
        made.locator('rt').evaluate((el) => {
          const cs = getComputedStyle(el)
          return [cs.fontSize, cs.bottom, cs.fontFamily.includes('Arial')]
        }),
      )
      .toEqual(['12px', '4px', true])
    await page.screenshot({ path: test.info().outputPath('phonetic-result.png') })
    await clickMenu(app, ['Format', 'Asian Layout', 'Phonetic Guide…'])
    await expect(phonetic.locator('select[name="font"]')).toHaveValue('Arial')
    await expect(phonetic.locator('input[name="size"]')).toHaveValue('9')
    await expect(phonetic.locator('input[name="offset"]')).toHaveValue('3')
    await page.keyboard.press('Escape')
    await expect(phonetic).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const cap = paragraphXml(xml, 'pon a time')
          const made = paragraphXml(xml, 'early')
          const kept = paragraphXml(xml, 'grows')
          const base = Number(/<w:hpsBaseText w:val="(\d+)"/.exec(made)?.[1])
          return [
            /<w:framePr w:dropCap="drop" w:lines="3" w:hSpace="360"/.test(cap),
            /<w:rFonts\b[^>]*w:ascii="Georgia"[^>]*\/>[\s\S]*?<w:t[^>]*>O<\/w:t>/.test(cap),
            /w:dropCap="drop" w:lines="2" w:hSpace="288"/.test(xml),
            /<w:hps w:val="18"\/>/.test(made),
            // three points over the resting place of a reading on this base
            Number(/<w:hpsRaise w:val="(\d+)"/.exec(made)?.[1]) === base - 2 + 6,
            /<w:rt><w:r><w:rPr><w:rFonts w:ascii="Arial"[^>]*\/><w:sz w:val="18"\/>/.test(made),
            /<w:hps w:val="16"\/><w:hpsRaise w:val="22"\/><w:hpsBaseText w:val="24"\/>/.test(kept),
            /<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"\/><w:sz w:val="16"\/>/.test(kept),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('the Styles pane modifies the current style; Hyphenation shows the zone of the document', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-small-dialogs-')))
  const docPath = join(dir, 'styles.docx')
  const zip = await JSZip.loadAsync(await docxWith(para('Styled words') + SECT_PR))
  zip.file(
    'word/styles.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>',
  )
  zip.file(
    'word/settings.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:hyphenationZone w:val="720"/></w:settings>',
  )
  writeFileSync(docPath, await zip.generateAsync({ type: 'nodebuffer' }))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const styled = page.locator('.doc-page p', { hasText: 'Styled words' }).first()
    await expect(styled).toBeVisible({ timeout: 30_000 })

    // Hyphenation: the zone the file carries is shown
    const hyphenation = page.locator('.hyphenation-dialog')
    await clickMenu(app, ['Tools', 'Hyphenation…'])
    const zone = hyphenation.getByLabel('Hyphenation zone')
    await expect(zone).toBeEnabled()
    await expect(zone).toHaveValue(/^(0\.5"|1\.27 cm)$/)
    await expect(hyphenation.getByRole('button', { name: 'Manual…' })).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('hyphenation.png') })
    await page.keyboard.press('Escape')
    await expect(hyphenation).toBeHidden()

    // Style…: Modify… opens the style of the text with the caret
    await caretAt(page, 'Styled')
    await clickMenu(app, ['Format', 'Style…'])
    const pane = page.locator('.styles-pane')
    await expect(pane).toBeVisible()
    await expect(pane.getByRole('button', { name: 'Delete', exact: true })).toBeDisabled()
    await expect(pane.getByRole('button', { name: 'Organizer…' })).toBeEnabled()
    await page.screenshot({ path: test.info().outputPath('styles-pane.png') })
    await pane.getByRole('button', { name: 'Modify…' }).click()
    const style = page.locator('.style-dialog')
    await expect(style.locator('h2')).toHaveText('Modify Style')
    await expect(style.getByLabel('Name')).toHaveValue('Normal')
    await style.getByLabel('Size').fill('20')
    await style.locator('button.btn-primary').click()
    await expect(style).toBeHidden()
    await expect
      .poll(() => styled.evaluate((el) => getComputedStyle(el).fontSize), { timeout: 15_000 })
      .toBe('26.6667px')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const styles = await part(docPath, 'word/styles.xml')
          const normal =
            (styles.match(/<w:style\b[\s\S]*?<\/w:style>/g) ?? []).find((s) =>
              /w:styleId="Normal"/.test(s),
            ) ?? ''
          return [
            /<w:sz w:val="40"\/>/.test(normal),
            /<w:hyphenationZone w:val="720"\/>/.test(await part(docPath, 'word/settings.xml')),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
