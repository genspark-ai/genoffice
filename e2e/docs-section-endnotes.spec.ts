/**
 * Endnotes at the end of each section (Footnote and Endnote dialog > End of
 * section), Format > Document… > Layout > Suppress endnotes, and the Phonetic
 * Guide's Right alignment. Each is checked on the page, on the sheets that
 * print, in the saved file and after reopening it.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-section-endnotes'

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const para = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`
const noted = (t: string, id: number) =>
  `<w:p><w:r><w:t>${t}</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r></w:p>`
const sectionEnd = (text: string, type = 'continuous', extra = '') =>
  `<w:p><w:pPr><w:sectPr><w:type w:val="${type}"/>${PAGE}${extra}</w:sectPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

async function endnoteDocx(
  body: string,
  notes: string[],
  pos?: string,
  lastType = 'continuous',
): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
  )
  zip.file(
    'word/settings.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:settings ${NS}>${pos ? `<w:endnotePr><w:pos w:val="${pos}"/><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>` : ''}</w:settings>`,
  )
  zip.file(
    'word/endnotes.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:endnotes ${NS}>` +
      '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' +
      '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>' +
      notes
        .map(
          (text, i) =>
            `<w:endnote w:id="${i + 1}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:endnote>`,
        )
        .join('') +
      '</w:endnotes>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}<w:sectPr><w:type w:val="${lastType}"/>${PAGE}</w:sectPr></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const ruby = (base: string, reading: string, align: string) =>
  `<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="${align}"/><w:hps w:val="12"/><w:hpsRaise w:val="22"/><w:hpsBaseText w:val="24"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt><w:r><w:rPr><w:sz w:val="12"/></w:rPr><w:t>${reading}</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>${base}</w:t></w:r></w:rubyBase></w:ruby></w:r>`

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

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

async function openDoc(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p').first()).toBeVisible({ timeout: 30_000 })
  return { launched, app: launched.app, page }
}

/** the caret goes into the paragraph; a click can lose to the focus a closing dialog hands back */
async function caretIn(page: Page, paragraph: string) {
  const target = page.locator('.doc-page p', { hasText: paragraph }).first()
  await expect(async () => {
    await target.click()
    await page.keyboard.press('End')
    const at = await page.evaluate(() => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { state: { selection: { $from: { parent: { textContent: string } } } } }
          }
        }
      ).__aidocs
      return editor.state.selection.$from.parent.textContent
    })
    expect(at).toContain(paragraph)
  }).toPass({ timeout: 15_000 })
}

/**
 * The body in reading order: paragraphs by their text, section-end endnote rows
 * as "#mark text", page turns as "GAP"; then the end-of-document area's rows.
 */
const flow = (page: Page) =>
  page.evaluate(() => {
    const pm = document.querySelector('.doc-page .ProseMirror, .doc-page.ProseMirror')!
    const body: string[] = []
    for (const el of Array.from(pm.children)) {
      if (el.classList.contains('page-sect-note')) {
        const mark = el.querySelector('sup')?.textContent ?? ''
        body.push(`#${mark} ${(el.textContent ?? '').slice(mark.length).trim()}`)
      } else if (el.classList.contains('page-gap')) body.push('GAP')
      else if (el.tagName === 'P') body.push((el.textContent ?? '').replace(/[ivxl]+$/, ''))
    }
    const tail = Array.from(document.querySelectorAll('.page-endnotes .page-note-text')).map(
      (el) => el.textContent ?? '',
    )
    return { body, tail }
  })

async function openFootnoteDialog(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Insert', 'Footnote…'])
  const dialog = page.locator('.footnote-dialog')
  await expect(dialog).toBeVisible()
  await dialog.locator('input[name="note-kind"][value="endnote"]').check()
  return dialog
}

async function openLayoutTab(app: ElectronApplication, page: Page) {
  const dialog = page.locator('.doc-dialog:not(.borders-shading-dialog)')
  await clickMenu(app, ['Format', 'Document…'])
  await dialog.locator('[role="tab"]').nth(1).click()
  return { dialog, suppress: dialog.locator('input[name="suppressEndnotes"]') }
}

interface DebugWindow {
  __pageDebug?: {
    slices: Array<{ start: number; end: number }>
    gapBoundaryMismatch: number
    pageCutMismatch: number
    pageCutDiag: string[]
  }
  __previewDebug?: { previewCutMismatch: number; previewCutDiag: string[]; shadowed: boolean }
  __previewShadow?: boolean
  __openPagePreview?: () => void
  __exportPdf?: () => Promise<boolean>
}

const NOTES = [
  'note for alpha',
  'first note for bravo',
  'second note for bravo',
  'note for charlie',
]
const THREE_SECTIONS =
  noted('Alpha starts.', 1) +
  sectionEnd('Alpha ends.') +
  noted('Bravo starts.', 2) +
  noted('Bravo goes on.', 3) +
  sectionEnd('Bravo ends.') +
  noted('Charlie starts.', 4) +
  para('Charlie ends.')

test('End of section puts each section’s endnotes under it, and Suppress endnotes hands them on', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-section-endnotes-')))
  const docPath = join(dir, 'sections.docx')
  writeFileSync(docPath, await endnoteDocx(THREE_SECTIONS, NOTES))
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    const atDocumentEnd = {
      body: [
        'Alpha starts.',
        'Alpha ends.',
        'Bravo starts.',
        'Bravo goes on.',
        'Bravo ends.',
        'Charlie starts.',
        'Charlie ends.',
      ],
      tail: NOTES,
    }
    await expect.poll(() => flow(page), { timeout: 30_000 }).toEqual(atDocumentEnd)

    // with endnotes at the document end there is nothing for a section to hand on
    await caretIn(page, 'Alpha starts.')
    let layout = await openLayoutTab(app, page)
    await expect(layout.suppress).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(layout.dialog).toBeHidden()

    const dialog = await openFootnoteDialog(app, page)
    const pos = dialog.locator('select[name="endnote-pos"]')
    await expect(pos).toHaveValue('docEnd')
    await pos.selectOption('sectEnd')
    await page.screenshot({ path: test.info().outputPath('footnote-dialog.png') })
    await dialog.locator('button[name="apply"]').click()
    await expect(dialog).toBeHidden()
    const atSectionEnds = {
      body: [
        'Alpha starts.',
        'Alpha ends.',
        '#i note for alpha',
        'Bravo starts.',
        'Bravo goes on.',
        'Bravo ends.',
        '#ii first note for bravo',
        '#iii second note for bravo',
        'Charlie starts.',
        'Charlie ends.',
      ],
      tail: ['note for charlie'],
    }
    await expect.poll(() => flow(page)).toEqual(atSectionEnds)
    // the rows take room in the text: the next section starts below them
    const top = (selector: string, text: string) =>
      page
        .locator(selector, { hasText: text })
        .first()
        .evaluate((el) => el.getBoundingClientRect().top)
    const alphaNote = await top('.page-sect-note', 'note for alpha')
    expect(alphaNote).toBeGreaterThan(await top('.doc-page p', 'Alpha ends.'))
    expect(await top('.doc-page p', 'Bravo starts.')).toBeGreaterThan(alphaNote + 10)
    await page.screenshot({ path: test.info().outputPath('section-ends.png') })

    // an endnote added in the first section joins that section's notes
    await caretIn(page, 'Alpha ends.')
    const insert = await openFootnoteDialog(app, page)
    await insert.locator('button.primary').click()
    await expect(insert).toBeHidden()
    const prompt = page.locator('.modal textarea.prompt-textarea')
    await prompt.fill('late note for alpha')
    await page.keyboard.press('ControlOrMeta+Enter')
    await expect(prompt).toBeHidden()
    await expect
      .poll(async () => (await flow(page)).body.slice(1, 5))
      .toEqual(['Alpha ends.', '#i note for alpha', '#ii late note for alpha', 'Bravo starts.'])

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/settings.xml'), { timeout: 30_000 })
      .toContain('<w:endnotePr><w:pos w:val="sectEnd"/>')

    // Suppress endnotes on the first section: its notes print with the second section's
    await caretIn(page, 'Alpha starts.')
    layout = await openLayoutTab(app, page)
    await expect(layout.suppress).toBeEnabled()
    await expect(layout.suppress).not.toBeChecked()
    await layout.suppress.check()
    await page.screenshot({ path: test.info().outputPath('document-dialog.png') })
    await layout.dialog.locator('button.primary').click()
    await expect(layout.dialog).toBeHidden()
    const handedOn = [
      'Alpha starts.',
      'Alpha ends.',
      'Bravo starts.',
      'Bravo goes on.',
      'Bravo ends.',
      '#i note for alpha',
      '#ii late note for alpha',
      '#iii first note for bravo',
      '#iv second note for bravo',
      'Charlie starts.',
      'Charlie ends.',
    ]
    await expect.poll(async () => (await flow(page)).body).toEqual(handedOn)
    await page.screenshot({ path: test.info().outputPath('suppressed.png') })

    // the sheets that print carry the same rows in the same place
    await page.evaluate(() => (window as unknown as DebugWindow).__openPagePreview?.())
    await expect(page.locator('.pv-page')).toHaveCount(1, { timeout: 30_000 })
    const printed = () =>
      page.evaluate(() =>
        Array.from(document.querySelectorAll('.pv-page p, .pv-page .page-sect-note'))
          .filter((el) => el.getBoundingClientRect().height > 0)
          .map((el) => [el.getBoundingClientRect().top, el.textContent ?? ''] as const)
          .sort((a, b) => a[0] - b[0])
          .map(([, text]) => text),
      )
    await expect
      .poll(async () => (await printed()).slice(4, 8))
      .toEqual([
        'Bravo ends.',
        'inote for alpha',
        'iilate note for alpha',
        'iiifirst note for bravo',
      ])
    await page.screenshot({ path: test.info().outputPath('suppressed-preview.png') })
    await page.keyboard.press('Escape')
    await expect(page.locator('.pv-page')).toHaveCount(0)

    await saveFromMenu(app)
    const sectPrs = async () =>
      (await part(docPath, 'word/document.xml')).match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []
    await expect
      .poll(async () => (await sectPrs()).map((s) => s.includes('<w:noEndnote/>')), {
        timeout: 30_000,
      })
      .toEqual([true, false, false])
    expect((await sectPrs())[0]).toContain('w:gutter="0"/><w:noEndnote/></w:sectPr>')

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect.poll(async () => (await flow(page)).body, { timeout: 30_000 }).toEqual(handedOn)
    const reopened = await openFootnoteDialog(app, page)
    await expect(reopened.locator('select[name="endnote-pos"]')).toHaveValue('sectEnd')
    await page.keyboard.press('Escape')
    await expect(reopened).toBeHidden()
    // the last section has no later one to hand its notes to
    await caretIn(page, 'Charlie starts.')
    layout = await openLayoutTab(app, page)
    await expect(layout.suppress).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(layout.dialog).toBeHidden()
    // switched off again, the first section prints its own notes
    await caretIn(page, 'Alpha starts.')
    layout = await openLayoutTab(app, page)
    await expect(layout.suppress).toBeChecked()
    await layout.suppress.uncheck()
    await layout.dialog.locator('button.primary').click()
    await expect(layout.dialog).toBeHidden()
    await expect
      .poll(async () => (await flow(page)).body.slice(1, 5))
      .toEqual(['Alpha ends.', '#i note for alpha', '#ii late note for alpha', 'Bravo starts.'])
    await saveFromMenu(app)
    await expect
      .poll(async () => (await sectPrs()).some((s) => s.includes('noEndnote')), { timeout: 30_000 })
      .toBe(false)

    // back at the end of the document, every note returns to the end area
    const back = await openFootnoteDialog(app, page)
    await back.locator('select[name="endnote-pos"]').selectOption('docEnd')
    await back.locator('button[name="apply"]').click()
    await expect(back).toBeHidden()
    await expect.poll(async () => (await flow(page)).body).toEqual(atDocumentEnd.body)
    await expect.poll(async () => (await flow(page)).tail).toHaveLength(5)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a file with endnotes at section ends pages them between entries, on the sheets and in the PDF', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-section-endnotes-')))
  const docPath = join(dir, 'pages.docx')
  const count = 20
  const body =
    Array.from({ length: count }, (_, i) => noted(`Alpha line ${i + 1}.`, i + 1)).join('') +
    sectionEnd('Alpha ends.', 'nextPage') +
    noted('Bravo starts.', count + 1) +
    para('Bravo ends.')
  const notes = Array.from({ length: count + 1 }, (_, i) => `note ${i + 1}`)
  writeFileSync(docPath, await endnoteDocx(body, notes, 'sectEnd', 'nextPage'))
  const session = await openDoc(docPath)
  try {
    const { app, page } = session
    await expect(page.locator('.page-sect-note')).toHaveCount(count, { timeout: 30_000 })
    // the first section's notes run over its page: the page turns between two entries,
    // and the second section still starts on a page of its own
    const pages = async () => {
      const { body } = await flow(page)
      const sheets: string[][] = [[]]
      for (const entry of body) {
        if (entry === 'GAP') sheets.push([])
        else sheets[sheets.length - 1]!.push(entry)
      }
      return sheets
    }
    await expect.poll(async () => (await pages()).length, { timeout: 30_000 }).toBe(3)
    const sheets = await pages()
    expect(sheets[0]![0]).toBe('Alpha line 1.')
    expect(sheets[0]!.at(-1)).toMatch(/^#\w+ note \d+$/)
    expect(sheets[1]![0]).toMatch(/^#\w+ note \d+$/)
    expect(sheets[1]!.at(-1)).toBe('#xx note 20')
    expect(sheets[2]).toEqual(['Bravo starts.', 'Bravo ends.'])
    expect((await flow(page)).tail).toEqual(['note 21'])
    const debug = await page.evaluate(() => (window as unknown as DebugWindow).__pageDebug)
    expect(debug?.slices).toHaveLength(3)
    expect(debug?.gapBoundaryMismatch).toBe(0)
    expect(debug?.pageCutMismatch, debug?.pageCutDiag.join(' | ')).toBe(0)
    await page.screenshot({ path: test.info().outputPath('paged.png') })

    await page.evaluate(() => {
      const w = window as unknown as DebugWindow
      w.__previewShadow = true
      w.__openPagePreview!()
    })
    await expect(page.locator('.pv-page')).toHaveCount(3, { timeout: 30_000 })
    await expect
      .poll(() => page.evaluate(() => (window as unknown as DebugWindow).__previewDebug?.shadowed))
      .toBe(true)
    const preview = await page.evaluate(() => (window as unknown as DebugWindow).__previewDebug)
    expect(preview?.previewCutMismatch, preview?.previewCutDiag.join(' | ')).toBe(0)

    const pdfPath = join(dir, 'pages.pdf')
    await app.evaluate(({ dialog, shell }, out) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    expect(await page.evaluate(() => (window as unknown as DebugWindow).__exportPdf!())).toBe(true)
    const pdf = readFileSync(pdfPath).toString('latin1')
    expect(pdf.match(/\/Type\s*\/Page\b/g)).toHaveLength(3)
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Phonetic Guide sets a reading flush right, on the page, the sheets and in the file', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ruby-right-')))
  const docPath = join(dir, 'ruby.docx')
  const line =
    `<w:p><w:r><w:t xml:space="preserve">Short reading </w:t></w:r>${ruby('Wide base text', 'ab', 'left')}` +
    `<w:r><w:t xml:space="preserve"> and long reading </w:t></w:r>${ruby('cd', 'a long reading', 'right')}` +
    `<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>`
  writeFileSync(docPath, await endnoteDocx(line + para('Second line.'), []))
  let session = await openDoc(docPath)
  try {
    let { app, page } = session
    /** the ink of a guide's base and reading (a range's boxes, not the element's), and its line */
    const drawn = (scope: string, index: number) =>
      page.evaluate(
        ([scope, index]) => {
          const el = document.querySelectorAll<HTMLElement>(`${scope} ruby.doc-ruby`)[
            index as number
          ]!
          const ink = (node: Node) => {
            const range = document.createRange()
            range.selectNodeContents(node)
            const box = range.getBoundingClientRect()
            return { left: box.left, right: box.right, top: box.top, bottom: box.bottom }
          }
          const base = el.querySelector('.doc-ruby-base')
          return {
            align: el.dataset.rubyAlign,
            display: getComputedStyle(el).display,
            base: base ? ink(base) : null,
            reading: ink(el.querySelector('rt')!),
            lineHeight: el.closest('p')!.getBoundingClientRect().height,
          }
        },
        [scope, index] as const,
      )
    const guides = page.locator('.doc-page ruby.doc-ruby')
    await expect(guides).toHaveCount(2, { timeout: 30_000 })
    // a guide that came flush right is drawn so: the short base ends where its long reading ends
    const fromFile = await drawn('.doc-page', 1)
    expect(fromFile.align).toBe('right')
    expect(fromFile.display).toBe('inline-grid')
    expect(Math.abs(fromFile.base!.right - fromFile.reading.right)).toBeLessThan(0.6)
    expect(fromFile.base!.left - fromFile.reading.left).toBeGreaterThan(10)
    expect(fromFile.reading.bottom).toBeLessThan(fromFile.base!.top + 4)
    const lineBefore = fromFile.lineHeight

    // the dialog offers Right and the page follows it
    await guides.first().click()
    await clickMenu(app, ['Format', 'Asian Layout', 'Phonetic Guide…'])
    const phonetic = page.locator('.phonetic-guide-dialog')
    const align = phonetic.locator('select[name="align"]')
    await expect(align).toHaveValue('left')
    await expect(align.locator('option[value="right"]')).toBeEnabled()
    await align.selectOption('right')
    await page.screenshot({ path: test.info().outputPath('phonetic-dialog.png') })
    await phonetic.locator('button.primary').click()
    await expect(phonetic).toBeHidden()
    await expect(guides.first()).toHaveAttribute('data-ruby-align', 'right')
    const set = await drawn('.doc-page', 0)
    // the short reading ends where its wide base ends, and the line is no taller
    expect(Math.abs(set.base!.right - set.reading.right)).toBeLessThan(0.6)
    expect(set.reading.left - set.base!.left).toBeGreaterThan(40)
    expect(set.reading.bottom).toBeLessThan(set.base!.top + 4)
    expect(Math.abs(set.lineHeight - lineBefore)).toBeLessThan(0.6)
    await page.screenshot({ path: test.info().outputPath('phonetic-right.png') })

    await page.evaluate(() => (window as unknown as DebugWindow).__openPagePreview?.())
    await expect(page.locator('.pv-page')).toHaveCount(1, { timeout: 30_000 })
    await expect(page.locator('.pv-page ruby.doc-ruby').first()).toBeVisible()
    const printed = await drawn('.pv-page', 0)
    expect(printed.display).toBe('inline-grid')
    expect(Math.abs(printed.base!.right - printed.reading.right)).toBeLessThan(0.6)
    expect(printed.reading.left - printed.base!.left).toBeGreaterThan(40)
    await page.keyboard.press('Escape')
    await expect(page.locator('.pv-page')).toHaveCount(0)

    await saveFromMenu(app)
    await expect
      .poll(
        async () =>
          (await part(docPath, 'word/document.xml')).match(/<w:rubyAlign w:val="(\w+)"\/>/g),
        {
          timeout: 30_000,
        },
      )
      .toEqual(['<w:rubyAlign w:val="right"/>', '<w:rubyAlign w:val="right"/>'])

    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    session = await openDoc(docPath)
    ;({ app, page } = session)
    await expect(page.locator('.doc-page ruby.doc-ruby')).toHaveCount(2, { timeout: 30_000 })
    const reopened = await drawn('.doc-page', 0)
    expect(reopened.align).toBe('right')
    expect(Math.abs(reopened.base!.right - reopened.reading.right)).toBeLessThan(0.6)
    await page.locator('.doc-page ruby.doc-ruby').first().click()
    await clickMenu(app, ['Format', 'Asian Layout', 'Phonetic Guide…'])
    await expect(page.locator('.phonetic-guide-dialog select[name="align"]')).toHaveValue('right')
    // back to centred, the guide is the browser's own ruby again
    await page.locator('.phonetic-guide-dialog select[name="align"]').selectOption('center')
    await page.locator('.phonetic-guide-dialog button.primary').click()
    await expect.poll(async () => (await drawn('.doc-page', 0)).display).toBe('ruby')
  } finally {
    await closeAndSaveVideo(session.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
