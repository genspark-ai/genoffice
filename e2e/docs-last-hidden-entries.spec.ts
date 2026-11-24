/**
 * Dialog entries that were left out until their function existed: Paste
 * Special's Paste link, AutoText's Look in, AutoFormat As You Type's Date
 * style and Define styles based on your formatting, and the original
 * document's label in Combine Documents. Each is driven through its dialog,
 * checked in the saved package, and read again after a reopen.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { existsSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-last-hidden-entries'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

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

const part = async (file: string, name: string): Promise<string> =>
  (await (await JSZip.loadAsync(readFileSync(file))).file(name)?.async('string')) ?? ''

const saveAs = async (app: ElectronApplication, out: string) => {
  await app.evaluate(({ dialog }, out) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
  }, out)
  await clickMenu(app, ['File', 'Save As…'])
  await expect.poll(() => existsSync(out), { timeout: 30_000 }).toBe(true)
}

const pickFile = (app: ElectronApplication, file: string) =>
  app.evaluate(({ dialog }, picked) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [picked] })
  }, file)

type TestEditor = {
  state: {
    doc: {
      descendants(
        f: (
          node: { isTextblock: boolean; textContent: string; nodeSize: number },
          pos: number,
        ) => boolean,
      ): void
    }
  }
  chain(): {
    focus(): {
      setTextSelection(at: number | { from: number; to: number }): { run(): boolean }
    }
  }
  commands: { focus(at: string): boolean }
}

/** selects from the start of the paragraph that starts with `first` to the end of the one that starts with `last` */
const selectParagraphs = (page: Page, first: string, last: string) =>
  page.evaluate(
    ([first, last]) => {
      const editor = (window as never as { __aidocs: { editor: TestEditor } }).__aidocs.editor
      let from = -1
      let to = -1
      editor.state.doc.descendants((node, pos) => {
        if (!node.isTextblock) return true
        if (from < 0 && node.textContent.startsWith(first)) from = pos + 1
        if (node.textContent.startsWith(last)) to = pos + node.nodeSize - 1
        return false
      })
      editor.chain().focus().setTextSelection({ from, to }).run()
      return new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
    },
    [first, last],
  )

const caretAtDocEnd = (page: Page) =>
  page.evaluate(() => {
    const editor = (window as never as { __aidocs: { editor: TestEditor } }).__aidocs.editor
    editor.commands.focus('end')
    return new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
  })

async function caretAtEndOf(page: Page, text: string) {
  await page.locator('.doc-page p', { hasText: text }).first().click()
  await page.keyboard.press('End')
}

/** the native inputs sit under their drawn boxes: the label takes the click */
const labelOf = (dialog: Locator, text: string) =>
  dialog.locator('label.protect-check', { hasText: text })

async function openLinks(app: ElectronApplication, page: Page) {
  await clickMenu(app, ['Edit', 'Links…'])
  const dialog = page.locator('.links-dialog')
  await expect(dialog).toBeVisible({ timeout: 15_000 })
  return dialog
}

async function openAutoCorrect(app: ElectronApplication, page: Page, tab: string) {
  await clickMenu(app, ['Tools', 'AutoCorrect Options…'])
  const dialog = page.getByRole('dialog', { name: 'AutoCorrect Options' })
  await expect(dialog).toBeVisible()
  await dialog.getByRole('tab', { name: tab, exact: true }).click()
  return dialog
}

async function closeDialog(dialog: Locator) {
  await dialog.getByRole('button', { name: 'Close', exact: true }).click()
  await expect(dialog).toBeHidden()
}

test('Paste link: a copy from a saved document goes in as a LINK field that Edit > Links updates', async () => {
  test.setTimeout(360_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-paste-link-')))
  const sourcePath = join(dir, 'source.docx')
  const hostPath = join(dir, 'host.docx')
  writeFileSync(
    sourcePath,
    await docxWith(
      ['Intro stays', 'Quarterly figure 100', 'Second linked line', 'Outro stays']
        .map(para)
        .join(''),
    ),
  )
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: sourcePath,
  })
  try {
    let { app } = launched
    const shell = await waitForPageWithUrl(app, 'shell/out')
    const source = await waitForPageWithUrl(app, '://docs/')
    await expect(source.locator('.doc-page p', { hasText: 'Intro stays' })).toBeVisible({
      timeout: 30_000,
    })

    // a second, untitled document to paste into
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
    const host = docsPages().find((p) => p !== source)!
    await expect(host.locator('.status-page')).toContainText(/\d/, { timeout: 30_000 })
    const tab = (name: string | RegExp, not = false) =>
      shell.locator(
        '.tab-bar .tab-item:not(.tab-home)',
        not ? { hasNotText: name } : { hasText: name },
      )

    // text that is not from a document cannot be linked: Word greys Paste link
    await app.evaluate(({ clipboard }) =>
      clipboard.write({ text: 'loose words', html: '<b>loose words</b>' }),
    )
    await caretAtDocEnd(host)
    await clickMenu(app, ['Edit', 'Paste Special…'])
    let paste = host.locator('.paste-special-dialog')
    await expect(paste.locator('.paste-special-as select')).toHaveValue('html')
    await expect(paste.getByRole('radio')).toHaveCount(2)
    await expect(paste.getByRole('radio').nth(0)).toBeChecked()
    await expect(paste.getByRole('radio').nth(1)).toBeDisabled()
    await expect(paste.getByText('Display as icon')).toHaveCount(0)
    await host.keyboard.press('Escape')
    await expect(paste).toBeHidden()

    // two paragraphs copied in the saved document
    await tab('source.docx').click()
    await selectParagraphs(source, 'Quarterly figure', 'Second linked line')
    await source.keyboard.press('ControlOrMeta+c')
    await expect
      .poll(() => app.evaluate(({ clipboard }) => clipboard.readHTML()))
      .toContain('genoffice-link')
    // the clipboard names no file
    expect(await app.evaluate(({ clipboard }) => clipboard.readHTML())).not.toContain(dir)

    await tab('source.docx', true).click()
    await caretAtDocEnd(host)
    await host.keyboard.type('Host line')
    await host.keyboard.press('Enter')
    await clickMenu(app, ['Edit', 'Paste Special…'])
    paste = host.locator('.paste-special-dialog')
    const link = paste.getByRole('radio').nth(1)
    await expect(link).toBeEnabled()
    await labelOf(paste, 'Paste link').click()
    await expect(link).toBeChecked()
    await expect(paste.locator('.paste-special-as select option')).toHaveText([
      'HTML Format',
      'Unformatted Unicode Text',
    ])
    await expect(paste.locator('.paste-special-result')).toContainText('link to the source file')
    await host.screenshot({ path: test.info().outputPath('paste-link.png') })
    await paste.locator('button.primary').click()
    await expect(paste).toBeHidden()
    const hostBody = host.locator('.doc-page')
    await expect(hostBody).toContainText('Quarterly figure 100')
    await expect(hostBody).toContainText('Second linked line')

    // the source document now bookmarks exactly the copied range
    await tab('source.docx').click()
    await source.locator('.doc-page p', { hasText: 'Intro stays' }).click()
    await saveFromMenu(app)
    await expect
      .poll(() => part(sourcePath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('w:name="LinkSource1"')
    let xml = await part(sourcePath, 'word/document.xml')
    const start = /<w:bookmarkStart w:id="(\d+)" w:name="LinkSource1"\/>/.exec(xml)!
    expect(xml.match(/w:name="LinkSource1"/g)).toHaveLength(1)
    const order = [
      'Intro stays',
      start[0],
      'Quarterly figure 100',
      'Second linked line',
      `<w:bookmarkEnd w:id="${start[1]}"/>`,
      'Outro stays',
    ].map((s) => xml.indexOf(s))
    expect(order.every((at) => at > 0)).toBe(true)
    expect([...order].sort((a, b) => a - b)).toEqual(order)

    // the source changes and is saved; a second save keeps one bookmark
    await caretAtEndOf(source, 'Quarterly figure 100')
    await source.keyboard.type(' and 250')
    await saveFromMenu(app)
    await expect
      .poll(() => part(sourcePath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('Quarterly figure 100 and 250')
    xml = await part(sourcePath, 'word/document.xml')
    expect(xml.match(/<w:bookmarkStart [^>]*w:name="LinkSource1"/g)).toHaveLength(1)

    // Edit > Links in the other document lists the link and reads the source again
    await tab('source.docx', true).click()
    await caretAtDocEnd(host)
    let links = await openLinks(app, host)
    const rows = links.locator('button.links-row')
    await expect(rows).toHaveCount(1)
    await expect(rows.first()).toContainText('source.docx')
    await expect(rows.first()).toContainText('Document')
    await expect(links.locator('.links-source')).toHaveText(sourcePath)
    await links.locator('button[name="update"]').click()
    await expect(rows.first()).toContainText('Updated')
    await expect(hostBody).toContainText('Quarterly figure 100 and 250')
    await expect(hostBody).toContainText('Second linked line')
    await expect(hostBody).toContainText('Host line')
    await host.keyboard.press('Escape')
    await expect(links).toBeHidden()

    await saveAs(app, hostPath)
    await expect
      .poll(() => part(hostPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('LINK Word.Document.12')
    xml = await part(hostPath, 'word/document.xml')
    expect(xml).toContain(`LINK Word.Document.12 "${sourcePath}" "LinkSource1" \\r`)
    expect(
      [...xml.matchAll(/w:fldCharType="(begin|separate|end)"/g)].map((m) => m[1][0]).join(''),
    ).toBe('bse')
    const result = [
      '"separate"',
      'Quarterly figure 100 and 250',
      'Second linked line',
      '"end"',
    ].map((s) => xml.indexOf(s))
    expect(result[0]).toBeGreaterThan(0)
    expect([...result].sort((a, b) => a - b)).toEqual(result)

    // reopened, the link is still a link: it is listed and can be broken
    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: hostPath })
    app = launched.app
    const reopened = await waitForPageWithUrl(app, '://docs/')
    await expect(reopened.locator('.doc-page')).toContainText('Second linked line', {
      timeout: 30_000,
    })
    await reopened.locator('.doc-page p', { hasText: 'Host line' }).click()
    links = await openLinks(app, reopened)
    await expect(links.locator('button.links-row')).toHaveCount(1)
    await expect(links.locator('button.links-row').first()).toContainText('source.docx')
    await expect(links.locator('button[name="update"]')).toBeEnabled()
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'

test('Date style and Define styles based on your formatting apply styles as the line is entered', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-autoformat-styles-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(para('Opening line') + para('') + SECT_PR))
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Opening line' })).toBeVisible({
      timeout: 30_000,
    })

    // both are off until asked for, as in Word: a date and a bold line stay plain paragraphs
    let dialog = await openAutoCorrect(app, page, 'AutoFormat As You Type')
    const dateStyle = dialog.getByLabel('Date style')
    const defineStyles = dialog.getByLabel('Define styles based on your formatting')
    await expect(dateStyle).not.toBeChecked()
    await expect(defineStyles).not.toBeChecked()
    await closeDialog(dialog)
    await caretAtDocEnd(page)
    await page.keyboard.type('November 20, 2026')
    await page.keyboard.press('Enter')
    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('November 20, 2026')
    expect(await part(docPath, 'word/document.xml')).not.toContain('w:pStyle w:val="Date"')

    dialog = await openAutoCorrect(app, page, 'AutoFormat As You Type')
    await labelOf(dialog, 'Date style').click()
    await labelOf(dialog, 'Define styles based on your formatting').click()
    await expect(dialog.getByLabel('Date style')).toBeChecked()
    await expect(dialog.getByLabel('Define styles based on your formatting')).toBeChecked()
    await page.screenshot({ path: test.info().outputPath('autoformat-styles.png') })
    await closeDialog(dialog)

    await caretAtDocEnd(page)
    await page.keyboard.type('November 23, 2026')
    await page.keyboard.press('Enter')
    await page.keyboard.type('Dear all, the report of November 23, 2026 follows')
    await page.keyboard.press('Enter')
    // a short line typed in bold is a heading in the making
    await page.keyboard.press('ControlOrMeta+b')
    await page.keyboard.type('Quarterly results')
    await page.keyboard.press('Enter')
    await page.keyboard.type('Body under the heading')
    const heading = page.locator('.doc-page :is(h1, h2, h3, h4, h5, h6)', {
      hasText: 'Quarterly results',
    })
    await expect(heading).toHaveCount(1)
    await expect(page.locator('.doc-page p', { hasText: 'Body under the heading' })).toHaveCSS(
      'font-weight',
      '400',
    )

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('Body under the heading')
    const paragraphOf = async (text: string) =>
      (await part(docPath, 'word/document.xml')).split('</w:p>').find((p) => p.includes(text)) ?? ''
    expect(await paragraphOf('November 23, 2026</w:t>')).toContain('<w:pStyle w:val="Date"/>')
    expect(await paragraphOf('November 20, 2026')).not.toContain('w:pStyle')
    expect(await paragraphOf('Dear all')).not.toContain('w:pStyle')
    const headingXml = await paragraphOf('Quarterly results')
    const headingStyle = /<w:pStyle w:val="([^"]+)"\/>/.exec(headingXml)?.[1]
    expect(headingStyle).toBeTruthy()
    // the bold is the style's: the run carries none
    expect(headingXml).not.toMatch(/<w:b\/>|<w:b w:val/)
    expect(await paragraphOf('Body under the heading')).not.toMatch(/<w:b\/>|w:pStyle/)
    const styles = await part(docPath, 'word/styles.xml')
    const styleOf = (id: string) =>
      new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`).exec(styles)?.[0] ??
      ''
    expect(styleOf('Date')).toContain('<w:name w:val="Date"/>')
    expect(styleOf('Date')).not.toContain('w:customStyle')
    expect(styleOf(headingStyle!)).toMatch(/<w:name w:val="heading \d"\/>/i)
    expect(styleOf(headingStyle!)).toMatch(/<w:b\/>/)

    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect(
      page.locator('.doc-page :is(h1, h2, h3, h4, h5, h6)', { hasText: 'Quarterly results' }),
    ).toHaveCount(1, { timeout: 30_000 })
    await expect(
      page.locator('.doc-page :is(h1, h2, h3, h4, h5, h6)', { hasText: 'Quarterly results' }),
    ).not.toHaveCSS('font-weight', '400')
    await expect(page.locator('.doc-page p', { hasText: 'Body under the heading' })).toHaveCSS(
      'font-weight',
      '400',
    )
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

const r = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`
const docPart = (name: string, gallery: string, body: string) =>
  `<w:docPart><w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/>` +
  `<w:gallery w:val="${gallery}"/></w:category><w:behaviors><w:behavior w:val="content"/></w:behaviors>` +
  `</w:docPartPr><w:docPartBody>${body}</w:docPartBody></w:docPart>`

/** a document that carries AutoText in its glossary part, as one made from a template does */
async function docxWithAutoText(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(para('Letter body') + para('') + SECT_PR))
  zip.file(
    'word/glossary/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:glossaryDocument ${W}><w:docParts>` +
      docPart(
        'Sign off',
        'autoTxt',
        `<w:p>${r('Kind regards, ')}${r('Finance team', '<w:b/>')}</w:p>`,
      ) +
      docPart('Prompt', 'placeholder', `<w:p>${r('Click to type')}</w:p>`) +
      docPart(
        'Office address',
        'autoTxt',
        `<w:p>${r('1 Main Street')}</w:p><w:p>${r('Springfield')}</w:p>`,
      ) +
      '</w:docParts></w:glossaryDocument>',
  )
  const rels = 'word/_rels/document.xml.rels'
  zip.file(
    rels,
    (await zip.file(rels)!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdGlossary" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/></Relationships>',
    ),
  )
  const types = '[Content_Types].xml'
  zip.file(
    types,
    (await zip.file(types)!.async('string')).replace(
      '</Types>',
      '<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/></Types>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('AutoText Look in lists the entries kept on this computer and the ones the document carries', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-autotext-lookin-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWithAutoText())
  let launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: docPath })
  try {
    let { app } = launched
    let page = await waitForPageWithUrl(app, '://docs/')
    const body = page.locator('.doc-page')
    await expect(body).toContainText('Letter body', { timeout: 30_000 })

    // an entry of this computer, made from the selection
    await selectParagraphs(page, 'Letter body', 'Letter body')
    let dialog = await openAutoCorrect(app, page, 'AutoText')
    const lookIn = dialog.locator('.ac-look-in select')
    const rows = dialog.locator('.ac-table .ac-row')
    await expect(lookIn).toBeEnabled()
    await expect(lookIn.locator('option')).toHaveText([
      'All active templates',
      'Normal',
      'letter.docx',
    ])
    await expect(lookIn).toHaveValue('all')
    await expect(rows).toHaveText(['Sign off', 'Office address'])
    await dialog.locator('.ac-entry-fields input').fill('Opening')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(rows).toHaveText(['Opening', 'Sign off', 'Office address'])

    await lookIn.selectOption('normal')
    await expect(rows).toHaveText(['Opening'])
    await lookIn.selectOption('document')
    await expect(rows).toHaveText(['Sign off', 'Office address'])
    // what the document carries is read, not edited here
    await rows.filter({ hasText: 'Office address' }).click()
    await expect(dialog.locator('.ac-preview')).toContainText('1 Main Street')
    await expect(dialog.getByRole('button', { name: 'Delete', exact: true })).toBeDisabled()
    await page.screenshot({ path: test.info().outputPath('autotext-look-in.png') })
    await closeDialog(dialog)

    // a one-paragraph entry goes into the line, with its formatting
    await caretAtDocEnd(page)
    await page.keyboard.type('Yours: ')
    dialog = await openAutoCorrect(app, page, 'AutoText')
    await dialog.locator('.ac-look-in select').selectOption('document')
    await dialog.locator('.ac-table .ac-row', { hasText: 'Sign off' }).click()
    await dialog.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(dialog).toBeHidden()
    await expect(body.locator('p', { hasText: 'Yours: Kind regards, Finance team' })).toHaveCount(1)

    // a two-paragraph entry comes in as its paragraphs
    await caretAtDocEnd(page)
    await page.keyboard.press('Enter')
    dialog = await openAutoCorrect(app, page, 'AutoText')
    await dialog.locator('.ac-table .ac-row', { hasText: 'Office address' }).click()
    await dialog.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(dialog).toBeHidden()
    await expect(body.locator('p', { hasText: '1 Main Street' })).toHaveCount(1)
    await expect(body.locator('p', { hasText: 'Springfield' })).toHaveCount(1)

    await saveFromMenu(app)
    await expect
      .poll(() => part(docPath, 'word/document.xml'), { timeout: 30_000 })
      .toContain('Springfield')
    const xml = await part(docPath, 'word/document.xml')
    const signOff = xml.split('</w:p>').find((p) => p.includes('Finance team')) ?? ''
    expect(signOff).toContain('Yours: ')
    expect(signOff).toMatch(
      /<w:rPr>(?:(?!<\/w:rPr>).)*<w:b\/>(?:(?!<\/w:rPr>).)*<\/w:rPr><w:t[^>]*>Finance team/,
    )
    expect(xml.indexOf('1 Main Street')).toBeLessThan(xml.indexOf('Springfield'))
    // the entries themselves stay in the package
    expect(await part(docPath, 'word/glossary/document.xml')).toContain('w:val="Office address"')

    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: VIDEO_DIR,
      openFile: docPath,
      userDataDir: launched.userDataDir,
    })
    app = launched.app
    page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page')).toContainText('Springfield', { timeout: 30_000 })
    await page.locator('.doc-page p', { hasText: 'Letter body' }).click()
    dialog = await openAutoCorrect(app, page, 'AutoText')
    await expect(dialog.locator('.ac-look-in select option')).toHaveText([
      'All active templates',
      'Normal',
      'letter.docx',
    ])
    await expect(dialog.locator('.ac-table .ac-row')).toHaveText([
      'Opening',
      'Sign off',
      'Office address',
    ])
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Combine Documents: the original side has a label, taken by its tracked changes that name no author', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-combine-label-')))
  const originalPath = join(dir, 'original.docx')
  const revisedPath = join(dir, 'revised.docx')
  const outPath = join(dir, 'combined.docx')
  const ins = (attrs: string, text: string) => `<w:ins w:id="${attrs}>${r(text)}</w:ins>`
  writeFileSync(
    originalPath,
    await docxWith(
      para('Kept as it is.') +
        `<w:p>${r('The plan ')}${ins('7" w:date="2026-11-01T08:00:00Z"', 'quietly ')}` +
        `${ins('8" w:author="Abe Origin" w:date="2026-11-01T08:00:00Z"', 'and slowly ')}${r('moves on.')}</w:p>`,
    ),
  )
  writeFileSync(
    revisedPath,
    await docxWith(para('Kept as it is.') + para('The plan moves on today.')),
  )
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: revisedPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'Kept as it is.' })).toBeVisible({
      timeout: 30_000,
    })

    // comparing takes one name only: the original's box stays grey, as in Word
    await clickMenu(app, ['Tools', 'Track Changes', 'Compare Documents…'])
    let dialog = page.locator('.cmp-dialog')
    await expect(dialog.locator('input[name="compare-label-original"]')).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()

    await pickFile(app, originalPath)
    await clickMenu(app, ['Tools', 'Combine Documents…'])
    dialog = page.locator('.cmp-dialog')
    const originalLabel = dialog.locator('input[name="compare-label-original"]')
    // written into the open document, its own changes are not renamed
    await expect(originalLabel).toBeDisabled()
    await dialog.locator('button[name="browse-original"]').click()
    await expect(dialog.locator('select[name="compare-original"] option:checked')).toHaveText(
      'original.docx',
    )
    await dialog.locator('select[name="compare-revised"]').selectOption('current')
    await expect(originalLabel).toBeEnabled()
    await expect(originalLabel).not.toHaveValue('')
    await originalLabel.fill('Olga Original')
    await dialog.locator('input[name="compare-label-revised"]').fill('Rae Visor')
    await dialog.locator('button[name="compare-options"]').click()
    await dialog.locator('input[name="compare-target"][value="new"]').check({ force: true })
    await expect(originalLabel).toHaveValue('Olga Original')
    await page.screenshot({ path: test.info().outputPath('combine-original-label.png') })
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)

    // the combination opens as a new document
    let result: Page | undefined
    await expect
      .poll(
        async () => {
          for (const candidate of app.windows()) {
            if (candidate === page) continue
            const href = await candidate.evaluate(() => window.location.href).catch(() => '')
            if (!href.includes('://docs/')) continue
            if (
              (await candidate
                .locator('.doc-page .doc-ins')
                .count()
                .catch(() => 0)) > 0
            )
              result = candidate
          }
          return !!result
        },
        { timeout: 60_000 },
      )
      .toBe(true)
    await expect(result!.locator('.doc-page')).toContainText('quietly')
    await saveAs(app, outPath)
    const xml = await part(outPath, 'word/document.xml')
    const authorOf = (text: string) =>
      new RegExp(`<w:ins [^>]*w:author="([^"]*)"[^>]*>(?:(?!</w:ins>).)*${text}`).exec(xml)?.[1]
    expect(authorOf('quietly')).toBe('Olga Original')
    expect(authorOf('and slowly')).toBe('Abe Origin')
    expect(authorOf('today')).toBe('Rae Visor')

    await closeAndSaveVideo(launched, VIDEO_DIR)
    launched = await launchShell({ onboardingSeen: true, videoDir: VIDEO_DIR, openFile: outPath })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(reopened.locator('.doc-page .doc-ins', { hasText: 'quietly' })).toHaveCount(1, {
      timeout: 30_000,
    })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR).catch(() => {})
    rmSync(dir, { recursive: true, force: true })
  }
})
