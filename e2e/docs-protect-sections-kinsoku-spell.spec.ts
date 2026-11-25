/**
 * Three dialog entries that used to be hidden: Protect Document > Sections
 * (forms protection section by section), Typography > Kinsoku > Custom (the
 * lists of characters that may not start or end a line) and AutoCorrect's
 * "Automatically use suggestions from the spelling checker".
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-protect-sections-kinsoku-spell'
const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
const sectPr = (inner = '') => `<w:sectPr>${PAGE}<w:cols w:space="720"/>${inner}</w:sectPr>`
const breakPara = (text: string, inner = '') =>
  `<w:p><w:pPr>${sectPr(inner)}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

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

async function withSettings(docx: Buffer, settings: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx)
  zip.file(
    'word/settings.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      `${settings}</w:settings>`,
  )
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    ),
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function open(docPath: string, shown: string, userDataDir?: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
    userDataDir,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page p', { hasText: shown }).first()).toBeVisible({
    timeout: 30_000,
  })
  return { launched, app: launched.app, page }
}

/** types at the end of the paragraph and says what the paragraph reads afterwards */
async function typeInto(page: Page, paragraph: Locator, text: string) {
  await paragraph.click()
  await page.keyboard.press('End')
  await page.keyboard.type(text)
  return paragraph.textContent()
}

test('Protect Document > Sections\u2026 keeps the unchecked section editable under forms protection', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-formprot-')))
  const docPath = join(dir, 'sections.docx')
  const mixedPath = join(dir, 'mixed.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('locked text') + breakPara('end of one') + para('open text') + para('tail') + sectPr(),
    ),
  )
  // as another program saved it: protected for forms, the first section left open
  writeFileSync(
    mixedPath,
    await withSettings(
      await docxWith(
        para('free text') +
          breakPara('end of one', '<w:formProt w:val="false"/>') +
          para('fixed text') +
          para('tail') +
          sectPr(),
      ),
      '<w:documentProtection w:edit="forms" w:enforcement="1"/>',
    ),
  )
  const paragraph = (page: Page, text: string) =>
    page.locator('.doc-page p', { hasText: text }).first()

  let run = await open(docPath, 'tail')
  try {
    const { app, page } = run
    await clickMenu(app, ['Tools', 'Protect Document\u2026'])
    const protect = page.locator('.protect-dialog:not(.protect-sections-dialog)')
    const sectionsButton = protect.getByRole('button', { name: 'Sections\u2026' })
    await expect(sectionsButton).toBeDisabled()
    // the native inputs sit under their drawn boxes: click the labels
    await protect.locator('label.protect-check', { hasText: 'Restrict editing' }).click()
    await expect(sectionsButton).toBeDisabled()
    await protect.locator('label.protect-check', { hasText: 'Forms' }).click()
    await expect(sectionsButton).toBeEnabled()
    await sectionsButton.click()
    const sections = page.getByRole('dialog', { name: 'Section Protection' })
    await expect(sections.locator('label.protect-check')).toHaveText(['Section 1', 'Section 2'])
    await expect(sections.getByRole('checkbox')).toHaveCount(2)
    for (const box of await sections.getByRole('checkbox').all()) await expect(box).toBeChecked()
    await sections.locator('label.protect-check', { hasText: 'Section 2' }).click()
    await expect(sections.getByRole('checkbox').nth(1)).not.toBeChecked()
    await page.screenshot({ path: test.info().outputPath('section-protection.png') })
    await sections.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(sections).toBeHidden()
    await protect.locator('button.btn-primary').first().click()
    await expect(protect).toBeHidden()

    expect(await typeInto(page, paragraph(page, 'open text'), ' edited')).toBe('open text edited')
    expect(await typeInto(page, paragraph(page, 'locked text'), ' edited')).toBe('locked text')
    await expect(page.locator('.status-msg')).toContainText('protected for forms')
    await page.keyboard.press('Backspace')
    await expect(paragraph(page, 'locked text')).toHaveText('locked text')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await part(docPath, 'word/document.xml')
          const [first, last] = xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []
          return [
            /<w:documentProtection\b[^>]*w:edit="forms"[^>]*w:enforcement="1"/.test(
              await part(docPath, 'word/settings.xml'),
            ),
            first?.includes('w:formProt'),
            last?.includes('<w:cols w:space="720"/><w:formProt w:val="0"/>'),
            xml.includes('open text edited'),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, false, true, true])
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
  }

  // reopened, the saved file behaves the same and the dialog shows what it holds
  run = await open(docPath, 'tail')
  try {
    const { app, page } = run
    expect(await typeInto(page, paragraph(page, 'tail'), '!')).toBe('tail!')
    expect(await typeInto(page, paragraph(page, 'locked text'), '!')).toBe('locked text')
    await clickMenu(app, ['Tools', 'Protect Document\u2026'])
    const protect = page.locator('.protect-dialog:not(.protect-sections-dialog)')
    await protect.getByRole('button', { name: 'Sections\u2026' }).click()
    const sections = page.getByRole('dialog', { name: 'Section Protection' })
    await expect(sections.getByRole('checkbox').nth(0)).toBeChecked()
    await expect(sections.getByRole('checkbox').nth(1)).not.toBeChecked()
    // the other way round: the first section open, the last one protected
    await sections.locator('label.protect-check', { hasText: 'Section 2' }).click()
    await sections.locator('label.protect-check', { hasText: 'Section 1' }).click()
    await sections.getByRole('button', { name: 'OK', exact: true }).click()
    await protect.locator('button.btn-primary').first().click()
    await expect(protect).toBeHidden()
    expect(await typeInto(page, paragraph(page, 'tail'), '?')).toBe('tail!')
    expect(await typeInto(page, paragraph(page, 'locked text'), ' now open')).toBe(
      'locked text now open',
    )
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await part(docPath, 'word/document.xml')
          const [first, last] = xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []
          return [first?.includes('<w:formProt w:val="0"/>'), last?.includes('w:formProt')]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, false])
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
  }

  // every section protected: the body takes no typing at all, and no section says otherwise
  run = await open(docPath, 'tail')
  try {
    const { app, page } = run
    await clickMenu(app, ['Tools', 'Protect Document\u2026'])
    const protect = page.locator('.protect-dialog:not(.protect-sections-dialog)')
    await protect.getByRole('button', { name: 'Sections\u2026' }).click()
    const sections = page.getByRole('dialog', { name: 'Section Protection' })
    await expect(sections.getByRole('checkbox').nth(0)).not.toBeChecked()
    await sections.locator('label.protect-check', { hasText: 'Section 1' }).click()
    await sections.getByRole('button', { name: 'OK', exact: true }).click()
    await protect.locator('button.btn-primary').first().click()
    await expect(protect).toBeHidden()
    await expect(page.locator('.doc-page[contenteditable="true"]')).toHaveCount(0)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).includes('w:formProt'), {
        timeout: 30_000,
      })
      .toBe(false)
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
  }

  run = await open(mixedPath, 'tail')
  try {
    const { page } = run
    expect(await typeInto(page, paragraph(page, 'free text'), ' edited')).toBe('free text edited')
    expect(await typeInto(page, paragraph(page, 'fixed text'), ' edited')).toBe('fixed text')
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Typography > Kinsoku > Custom replaces the line-start and line-end lists for its language', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-kinsoku-custom-')))
  const docPath = join(dir, 'custom.docx')
  const KANJI = '\u56fd'
  const KANA = '\u3042'
  const MARK = '\u3005'
  const ja = (text: string) =>
    `<w:p><w:r><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`
  // one of these lengths puts the probed character exactly where the second line starts
  const sweep = (probe: string) =>
    Array.from({ length: 36 }, (_, i) => ja(KANJI.repeat(22 + i) + probe + KANJI.repeat(3)))
  writeFileSync(
    docPath,
    await docxWith(para('head') + sweep(KANA).join('') + sweep(MARK).join('') + sectPr()),
  )
  /** how many lines start with each probed character */
  const lineStarts = (page: Page) =>
    page.evaluate(
      ([kana, mark]) => {
        const count = { kana: 0, mark: 0 }
        for (const para of document.querySelectorAll('.doc-page p')) {
          const rects: Array<{ ch: string; rect: DOMRect }> = []
          const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT)
          while (walker.nextNode()) {
            const node = walker.currentNode as Text
            for (let i = 0; i < node.data.length; i++) {
              if (node.data[i] === '\u2060') continue
              const range = document.createRange()
              range.setStart(node, i)
              range.setEnd(node, i + 1)
              rects.push({ ch: node.data[i], rect: range.getBoundingClientRect() })
            }
          }
          rects.forEach((c, i) => {
            if (i === 0 || rects[i - 1].rect.bottom > c.rect.top + 1) return
            if (c.ch === kana) count.kana++
            if (c.ch === mark) count.mark++
          })
        }
        return count
      },
      [KANA, MARK],
    )
  const openTypography = async (app: ElectronApplication, page: Page) => {
    await page.locator('.doc-page p', { hasText: 'head' }).first().click()
    await clickMenu(app, ['Format', 'Paragraph\u2026'])
    const dialog = page.locator('.modal', { hasText: 'Paragraph' }).first()
    await dialog.getByRole('tab', { name: 'Asian Typography' }).click()
    await dialog.getByRole('button', { name: 'Options\u2026' }).click()
    const typography = page.getByRole('dialog', { name: 'Typography' })
    await expect(typography).toBeVisible()
    return { dialog, typography }
  }
  const confirm = async (dialog: Locator, typography: Locator) => {
    await typography.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(typography).toBeHidden()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()
  }

  let run = await open(docPath, 'head')
  try {
    const { app, page } = run
    // Standard: a plain kana may start a line, the iteration mark may not
    const standard = await lineStarts(page)
    expect(standard.kana).toBeGreaterThan(0)
    expect(standard.mark).toBe(0)

    const { dialog, typography } = await openTypography(app, page)
    await expect(typography.getByLabel('Standard')).toBeChecked()
    await expect(typography.locator('textarea')).toHaveCount(0)
    await typography.getByLabel('Custom').check()
    // Custom starts from the standard lists of the language shown
    const [noStart, noEnd] = [
      typography.getByLabel('Cannot start line:'),
      typography.getByLabel('Cannot end line:'),
    ]
    await expect(noStart).toHaveValue(/\u3005/)
    await expect(noEnd).toHaveValue(/\uff08/)
    await noStart.fill(KANA)
    await noEnd.fill('\uff08')
    await page.screenshot({ path: test.info().outputPath('typography-custom.png') })
    await confirm(dialog, typography)

    await expect.poll(async () => (await lineStarts(page)).kana).toBe(0)
    await expect.poll(async () => (await lineStarts(page)).mark).toBeGreaterThan(0)

    await saveFromMenu(app)
    await expect
      .poll(
        async () =>
          (await part(docPath, 'word/settings.xml')).match(/<w:noLineBreaks[^>]*>/g) ?? [],
        {
          timeout: 30_000,
        },
      )
      .toEqual([
        `<w:noLineBreaksAfter w:lang="ja-JP" w:val="\uff08"/>`,
        `<w:noLineBreaksBefore w:lang="ja-JP" w:val="${KANA}"/>`,
      ])
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
  }

  // the saved lists are honoured on open and shown in the dialog; Standard takes them away
  run = await open(docPath, 'head')
  try {
    const { app, page } = run
    await expect.poll(async () => (await lineStarts(page)).kana).toBe(0)
    expect((await lineStarts(page)).mark).toBeGreaterThan(0)
    const { dialog, typography } = await openTypography(app, page)
    await expect(typography.getByLabel('Custom')).toBeChecked()
    await expect(typography.getByLabel('Cannot start line:')).toHaveValue(KANA)
    await expect(typography.getByLabel('Cannot end line:')).toHaveValue('\uff08')
    await typography.getByLabel('Standard').check()
    await confirm(dialog, typography)
    await expect.poll(async () => (await lineStarts(page)).mark).toBe(0)
    expect((await lineStarts(page)).kana).toBeGreaterThan(0)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/settings.xml')).includes('noLineBreaks'), {
        timeout: 30_000,
      })
      .toBe(false)
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('AutoCorrect replaces a misspelling with the spelling checker\u2019s only suggestion, and Backspace brings it back', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-spell-suggest-')))
  const docPath = join(dir, 'spell.docx')
  writeFileSync(docPath, await docxWith(para('head') + para('') + sectPr()))
  const option = 'Automatically use suggestions from the spelling checker'
  const CANDIDATES = [
    'recieve',
    'definately',
    'seperate',
    'occured',
    'untill',
    'beleive',
    'goverment',
    'tommorow',
    'accomodate',
    'adress',
    'calender',
    'freind',
    'truely',
    'neccessary',
    'wich',
    'becuase',
    'enviroment',
  ]
  const openAutoCorrect = async (app: ElectronApplication, page: Page) => {
    await clickMenu(app, ['Tools', 'AutoCorrect Options\u2026'])
    const dialog = page.getByRole('dialog', { name: 'AutoCorrect Options' })
    await expect(dialog).toBeVisible()
    return dialog
  }
  const lastParagraph = (page: Page) => page.locator('.doc-page p').last()

  let run = await open(docPath, 'head')
  const { userDataDir } = run.launched
  try {
    const { app, page } = run
    // a word the checker of this machine knows exactly one suggestion for
    const pick = await page.evaluate((words) => {
      const api = (
        window as unknown as {
          desktop: {
            spellIsMisspelled?: (word: string) => boolean
            spellSuggestions?: (word: string) => string[]
          }
        }
      ).desktop
      for (const word of words) {
        if (!api.spellIsMisspelled?.(word)) continue
        const offered = api.spellSuggestions?.(word) ?? []
        if (offered.length === 1) return { word, to: offered[0] }
      }
      return null
    }, CANDIDATES)
    test.skip(!pick, 'the spelling checker here has no single suggestion for any candidate')
    const { word, to } = pick!

    // off until asked for: the word stays as typed
    await lastParagraph(page).click()
    await page.keyboard.type(`${word} `)
    await expect(lastParagraph(page)).toHaveText(`${word} `, { useInnerText: false })

    const dialog = await openAutoCorrect(app, page)
    const box = dialog.getByLabel(option)
    await expect(box).not.toBeChecked()
    await box.setChecked(true, { force: true })
    await page.screenshot({ path: test.info().outputPath('autocorrect-spelling.png') })
    await dialog.getByRole('button', { name: 'Close', exact: true }).click()
    await expect(dialog).toBeHidden()

    await lastParagraph(page).click()
    await page.keyboard.press('End')
    await page.keyboard.type(`${word} `)
    await expect.poll(() => lastParagraph(page).textContent()).toBe(`${word} ${to} `)
    // Backspace right after the correction puts the typed word back
    await page.keyboard.press('Backspace')
    await expect.poll(() => lastParagraph(page).textContent()).toBe(`${word} ${word} `)
    await page.keyboard.type(`${word}.`)
    await expect.poll(() => lastParagraph(page).textContent()).toBe(`${word} ${word} ${to}.`)

    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).includes(`${to}.`), {
        timeout: 30_000,
      })
      .toBe(true)
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
  }

  // an app setting: it is still on after a restart
  run = await open(docPath, 'head', userDataDir)
  try {
    const { app, page } = run
    const dialog = await openAutoCorrect(app, page)
    await expect(dialog.getByLabel(option)).toBeChecked()
    await dialog.getByLabel(option).setChecked(false, { force: true })
    await dialog.getByRole('button', { name: 'Close', exact: true }).click()
    await lastParagraph(page).click()
    await page.keyboard.press('End')
    await page.keyboard.type(' recieve ')
    await expect.poll(() => lastParagraph(page).textContent()).toContain(' recieve ')
  } finally {
    await closeAndSaveVideo(run.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
