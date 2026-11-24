/**
 * The commands behind the last chords of Word for Mac to get one: Normal
 * style on the keypad, Mark Table of Contents Entry, Copy Text, Insert Merge
 * Field, Link to Previous and Next Misspelling; and the menu bar showing the
 * shortcut a document assigns. On the Mac the chord is pressed; elsewhere,
 * where the chords are not bound, the command is sent as the menu sends one.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-word-keys-leftovers'
const IS_MAC = process.platform === 'darwin'

const styled = (style: string, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
const PAGE =
  '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'

const scratch = (name: string) =>
  realpathSync(mkdtempSync(join(tmpdir(), `genoffice-e2e-${name}-`)))

async function open(docPath: string, shows: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )
  await expect(page.locator('.doc-page p', { hasText: shows }).first()).toBeAttached({
    timeout: 30_000,
  })
  // the open-time respell kick focuses the body and shields keystrokes for a moment
  await page.waitForTimeout(1500)
  return { launched, page }
}

async function close(launched: LaunchedApp, dir: string) {
  await closeAndSaveVideo(launched, VIDEO_DIR)
  rmSync(dir, { recursive: true, force: true })
}

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')!.async('string')

/** the command as the menu bar sends it to the document */
const send = (app: ElectronApplication, command: string) =>
  app.evaluate(({ webContents }, command) => {
    for (const contents of webContents.getAllWebContents())
      if (contents.getURL().includes('://docs/')) contents.send('menu:command', command)
  }, command)

/** Word's chord on the Mac; the command itself where the chord is not bound */
async function fire(app: ElectronApplication, page: Page, chord: string, command: string) {
  if (IS_MAC) await page.keyboard.press(chord)
  else await send(app, command)
  // the command crosses to the page on its own channel: let it land before the next key
  await page.waitForTimeout(300)
}

/** the caret into the paragraph that holds `needle`, `offset` characters in, selecting `length` */
const caretIn = (page: Page, needle: string, offset = 0, length = 0) =>
  page.evaluate(
    ({ needle, offset, length }) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: {
                doc: {
                  descendants(
                    visit: (
                      node: { isTextblock: boolean; textContent: string },
                      pos: number,
                    ) => void,
                  ): void
                }
              }
              chain(): {
                focus(): {
                  setTextSelection(range: { from: number; to: number }): { run(): boolean }
                }
              }
            }
          }
        }
      ).__aidocs
      let at = -1
      editor.state.doc.descendants((node, pos) => {
        const found = node.isTextblock ? node.textContent.indexOf(needle) : -1
        if (at < 0 && found >= 0) at = pos + 1 + found
      })
      if (at < 0) throw new Error(`no paragraph contains "${needle}"`)
      editor
        .chain()
        .focus()
        .setTextSelection({ from: at + offset, to: at + offset + length })
        .run()
    },
    { needle, offset, length },
  )

const selectedText = (page: Page): Promise<string> =>
  page.evaluate(() => {
    const { state } = (
      window as never as {
        __aidocs: {
          editor: {
            state: {
              doc: { textBetween(from: number, to: number, sep: string): string }
              selection: { from: number; to: number }
            }
          }
        }
      }
    ).__aidocs.editor
    return state.doc.textBetween(state.selection.from, state.selection.to, '\n')
  })

test('Normal style, Mark Table of Contents Entry and Copy Text run from their keys and reach the file', async () => {
  test.setTimeout(180_000)
  const dir = scratch('keys-leftovers')
  const docPath = join(dir, 'keys.docx')
  writeFileSync(
    docPath,
    await docxWith(
      styled('Heading1', 'Alpha heading') +
        para('Intro text here') +
        para('Second para') +
        para('Last paragraph'),
    ),
  )
  const { launched, page } = await open(docPath, 'Last paragraph')
  const { app } = launched
  try {
    // ---- Normal style ----
    await expect(page.locator('.doc-page h1', { hasText: 'Alpha heading' })).toHaveCount(1)
    await caretIn(page, 'Alpha heading', 2)
    // the driver's Numpad5 is the Mac keypad's Clear
    await fire(app, page, 'Alt+Shift+Numpad5', 'style-normal')
    await expect(page.locator('.doc-page p', { hasText: 'Alpha heading' })).toHaveCount(1)
    await expect(page.locator('.doc-page h1')).toHaveCount(0)

    // ---- Copy Text: what is selected goes to where Return is pressed ----
    await caretIn(page, 'Second para', 0, 'Second'.length)
    await fire(app, page, 'Shift+F2', 'copy-text')
    await expect(page.getByText('Copy to where?')).toBeVisible()
    await caretIn(page, 'Last paragraph', 'Last paragraph'.length)
    await page.keyboard.press('Enter')
    await expect(page.locator('.doc-page p', { hasText: 'Last paragraphSecond' })).toHaveCount(1)
    await expect(page.locator('.doc-page p', { hasText: /^Second para$/ })).toHaveCount(1)
    // with a bare caret it asks what to copy, and brings that to the caret
    await caretIn(page, 'Second para', 'Second para'.length)
    await fire(app, page, 'Shift+F2', 'copy-text')
    await expect(page.getByText('Copy from where?')).toBeVisible()
    await caretIn(page, 'Alpha heading', 0, 'Alpha'.length)
    await page.keyboard.press('Enter')
    await expect(page.locator('.doc-page p', { hasText: /^Second paraAlpha$/ })).toHaveCount(1)
    // Escape leaves the mode: Return is a new paragraph again
    await caretIn(page, 'Intro text here', 0, 'Intro'.length)
    await fire(app, page, 'Shift+F2', 'copy-text')
    await page.keyboard.press('Escape')
    const paragraphs = await page.locator('.doc-page p').count()
    await caretIn(page, 'Last paragraphSecond', 'Last paragraphSecond'.length)
    await page.keyboard.press('Enter')
    await expect(page.locator('.doc-page p')).toHaveCount(paragraphs + 1)
    await expect(page.locator('.doc-page p', { hasText: 'Last paragraphSecond' })).toHaveText(
      'Last paragraphSecond',
    )

    // ---- Mark Table of Contents Entry ----
    await caretIn(page, 'Intro text here', 0, 'Intro'.length)
    await fire(app, page, 'Alt+Shift+Meta+O', 'mark-toc-entry')
    const dialog = page.locator('.mark-toc-entry-dialog')
    await expect(dialog).toBeVisible()
    await expect(dialog.locator('h2')).toHaveText('Mark Table of Contents Entry')
    await expect(dialog.locator('input[name="entry"]')).toHaveValue('Intro')
    await expect(dialog.locator('select[name="table"]')).toHaveValue('C')
    await dialog.locator('select[name="level"]').selectOption('2')
    await dialog.getByRole('button', { name: 'Mark', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    await expect(page.locator('.doc-page .doc-tc-mark')).toHaveCount(1)

    await saveFromMenu(app)
    await expect
      .poll(async () => (await documentXml(docPath)).includes('Last paragraphSecond'), {
        timeout: 20_000,
      })
      .toBe(true)
    const xml = await documentXml(docPath)
    const codes = [...xml.matchAll(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g)].map((m) =>
      m[1].replace(/&quot;/g, '"').trim(),
    )
    expect(codes).toEqual(['TC "Intro" \\l 2'])
    expect(xml).not.toContain('w:val="Heading1"')
    expect(xml).toContain('Second paraAlpha')
  } finally {
    await close(launched, dir)
  }
})

test('Insert Merge Field opens the list of merge fields once a recipient list is attached', async () => {
  test.setTimeout(180_000)
  const dir = scratch('merge-field-key')
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(para('Dear') + para('Last paragraph')))
  const listPath = join(dir, 'people.csv')
  writeFileSync(listPath, 'Name,City\nAda,London\nAlan,Wilmslow\n')
  const { launched, page } = await open(docPath, 'Last paragraph')
  const { app } = launched
  try {
    const tab = (name: string) =>
      page.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).click()
    const fields = page.locator('.rb-field-item')
    // nothing to pick from yet
    await caretIn(page, 'Dear', 'Dear'.length)
    await fire(app, page, 'Control+Shift+F', 'insert-merge-field')
    await page.waitForTimeout(500)
    await expect(fields).toHaveCount(0)

    await app.evaluate(({ dialog }, list) => {
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [list] })
    }, listPath)
    await tab('Mailings')
    await page
      .locator('[data-ribbon-body] button:is(.rb-big, .rb-small)', { hasText: 'Select Recipients' })
      .click()
    await expect(
      page.locator('[data-ribbon-body] [data-testid="mailings-record-status"]', {
        hasText: 'Record 1 of 2',
      }),
    ).toHaveCount(1)

    await tab('Home')
    await caretIn(page, 'Dear', 'Dear'.length)
    await fire(app, page, 'Control+Shift+F', 'insert-merge-field')
    await expect(fields).toHaveText(['Name', 'City'])
    await fields.filter({ hasText: /^City$/ }).click()
    await expect(fields).toHaveCount(0)
    await saveFromMenu(app)
    await expect
      .poll(async () => /MERGEFIELD\s+"?City/.test(await documentXml(docPath)), { timeout: 20_000 })
      .toBe(true)
  } finally {
    await close(launched, dir)
  }
})

/** two sections, each long enough to span pages, section 1 carrying a plain header */
async function twoSections(): Promise<Buffer> {
  const filler = (label: string) =>
    Array.from({ length: 70 }, (_, i) => para(`${label} paragraph ${i + 1}`)).join('')
  const zip = await JSZip.loadAsync(
    await docxWith(
      filler('First') +
        `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/>${PAGE}</w:sectPr></w:pPr></w:p>` +
        filler('Second') +
        para('Last paragraph') +
        `<w:sectPr>${PAGE}</w:sectPr>`,
    ),
  )
  zip.file(
    'word/header1.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Plain header</w:t></w:r></w:p></w:hdr>',
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>',
    ),
  )
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('Link to Previous turns off and on for the header being edited', async () => {
  test.setTimeout(180_000)
  const dir = scratch('hf-link-key')
  const docPath = join(dir, 'sections.docx')
  writeFileSync(docPath, await twoSections())
  const { launched, page } = await open(docPath, 'Last paragraph')
  const { app } = launched
  try {
    // with no header open the command has nothing to act on
    await caretIn(page, 'First paragraph 1', 2)
    await fire(app, page, 'Control+Shift+R', 'hf-link-previous')

    const gapHeader = page.locator('.page-gap-hf.page-hf-header[data-hf-section="1"]').first()
    await gapHeader.scrollIntoViewIfNeeded()
    await expect(gapHeader).toContainText('Plain header')
    await gapHeader.dblclick()
    await expect(gapHeader.locator('.page-hf-editor')).toBeVisible()
    const link = page.getByRole('button', { name: 'Link to Previous' })
    await expect(link).toHaveClass(/active/)
    await fire(app, page, 'Control+Shift+R', 'hf-link-previous')
    await expect(link).not.toHaveClass(/active/)
    // unlinking re-paginates; the editor follows the strip into its new widget and takes focus back
    await expect(gapHeader.locator('.page-hf-editor')).toBeFocused()
    page.once('dialog', (d) => void d.accept())
    await fire(app, page, 'Control+Shift+R', 'hf-link-previous')
    // linked again: the button says so, or the strip that was discarded closed its editor
    await expect
      .poll(
        async () =>
          (await link.count()) === 0 || /\bactive\b/.test((await link.getAttribute('class')) ?? ''),
        { timeout: 15_000 },
      )
      .toBe(true)
    await expect(gapHeader).toContainText('Plain header')
  } finally {
    await close(launched, dir)
  }
})

const KEYS_REL = 'http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations'
const KEYS_TYPE = 'application/vnd.ms-word.keyMapCustomizations+xml'
// Control+Shift+Y runs Word Count
const CUSTOMIZATIONS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<wne:tcg xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"><wne:keymaps>' +
  '<wne:keymap wne:kcmPrimary="0359"><wne:fci wne:fciName="ToolsWordCount" wne:swArg="0000"/></wne:keymap>' +
  '</wne:keymaps></wne:tcg>'

async function withKeys(docx: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(docx)
  zip.file('word/customizations.xml', CUSTOMIZATIONS)
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      `<Relationship Id="rIdKeys" Type="${KEYS_REL}" Target="customizations.xml"/></Relationships>`,
    ),
  )
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      `<Override PartName="/word/customizations.xml" ContentType="${KEYS_TYPE}"/></Types>`,
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('the menu bar shows the shortcut the open document assigns, and not in another document', async () => {
  test.setTimeout(180_000)
  const dir = scratch('doc-keys-menu')
  const docPath = join(dir, 'keys.docx')
  writeFileSync(docPath, await withKeys(await docxWith(para('Keys') + para('Last paragraph'))))
  const { launched, page } = await open(docPath, 'Last paragraph')
  const { app } = launched
  try {
    const shortcut = () =>
      app.evaluate(({ Menu }) => {
        const tools = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Tools')
        const item = tools?.submenu?.items.find((candidate) => candidate.label === 'Word Count…')
        return item ? (item.accelerator ?? '') : 'no item'
      })
    await expect.poll(shortcut, { timeout: 20_000 }).toBe('Ctrl+Shift+Y')

    // a second document without keys of its own: its menu has no shortcut for the command
    await expect
      .poll(
        async () => {
          for (const candidate of app.windows()) {
            const opened = await candidate
              .evaluate(async () => {
                const api = (window as never as { aiOffice?: { newDoc(): Promise<void> } }).aiOffice
                if (!api) return false
                await api.newDoc()
                return true
              })
              .catch(() => false)
            if (opened) return true
          }
          return false
        },
        { timeout: 30_000 },
      )
      .toBe(true)
    const docsPages = () => app.windows().filter((p) => p.url().includes('://docs/'))
    await expect.poll(() => docsPages().length, { timeout: 30_000 }).toBe(2)
    const second = docsPages().find((p) => p !== page)!
    await expect(second.locator('.status-page')).toContainText(/\d/, { timeout: 30_000 })
    await expect.poll(shortcut, { timeout: 20_000 }).toBe('')
  } finally {
    await close(launched, dir)
  }
})

test('Next Misspelling selects the next misspelled word and shows its menu', async () => {
  test.setTimeout(180_000)
  const dir = scratch('next-misspelling')
  const docPath = join(dir, 'spell.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('This has a qzvxwk word') + para('and then a zzkqpv one') + para('Last paragraph'),
    ),
  )
  const { launched, page } = await open(docPath, 'Last paragraph')
  const { app } = launched
  try {
    const knowsWords = await page.evaluate(async () => {
      const api = (window as never as { desktop: { spellIsMisspelled?(word: string): boolean } })
        .desktop
      for (let i = 0; i < 40; i++) {
        if (api.spellIsMisspelled?.('qzvxwk') && !api.spellIsMisspelled?.('word')) return true
        await new Promise((resolve) => setTimeout(resolve, 500))
      }
      return false
    })
    test.skip(!knowsWords, 'no spelling dictionary is loaded on this machine')

    const menu = page.locator('.ctx-menu')
    await caretIn(page, 'This has a qzvxwk word', 0)
    await fire(app, page, 'Alt+F7', 'spell-next')
    await expect(menu).toBeVisible()
    await expect(menu).toContainText('Add to Dictionary')
    expect(await selectedText(page)).toBe('qzvxwk')
    await page.keyboard.press('Escape')
    await expect(menu).toHaveCount(0)
    await fire(app, page, 'Alt+F7', 'spell-next')
    await expect(menu).toBeVisible()
    expect(await selectedText(page)).toBe('zzkqpv')
    await page.keyboard.press('Escape')
    // round the end of the document
    await fire(app, page, 'Alt+F7', 'spell-next')
    await expect.poll(() => selectedText(page)).toBe('qzvxwk')
  } finally {
    await close(launched, dir)
  }
})
