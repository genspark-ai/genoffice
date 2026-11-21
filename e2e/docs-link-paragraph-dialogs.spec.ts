/**
 * Insert > Link beyond a web address (a place in this document, an e-mail
 * address, a ScreenTip) and the Format > Paragraph options Mirror indents,
 * Don't hyphenate, Snap to grid, Tabs… and Set As Default: each is set in the
 * dialog, shows on the page, is saved and comes back when the file is reopened.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const STYLES =
  `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>` +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  '</w:styles>'

/** the fixture document plus a styles part and, when given, a settings part */
async function docxWithStyles(body: string, settings = ''): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body))
  const add = (file: string, path: string, type: string, rel: string, xml: string) => {
    zip.file(path, xml)
    return [
      `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`,
      `<Relationship Id="rId${file}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${rel}" Target="${file}.xml"/>`,
    ]
  }
  const parts = [add('styles', 'word/styles.xml', 'styles', 'styles', STYLES)]
  if (settings)
    parts.push(
      add(
        'settings',
        'word/settings.xml',
        'settings',
        'settings',
        `<?xml version="1.0" encoding="UTF-8"?><w:settings ${W}>${settings}</w:settings>`,
      ),
    )
  const patch = async (path: string, close: string, index: 0 | 1) => {
    const xml = await zip.file(path)!.async('string')
    zip.file(path, xml.replace(close, parts.map((p) => p[index]).join('') + close))
  }
  await patch('[Content_Types].xml', '</Types>', 0)
  await patch('word/_rels/document.xml.rels', '</Relationships>', 1)
  return zip.generateAsync({ type: 'nodebuffer' })
}

type AidocsWindow = {
  __aidocs?: {
    editor?: { state: { selection: { $from: { parent: { textContent: string } } } } }
  }
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

const partOf = async (docPath: string, part: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(part)!.async('string')

async function openDoc(docPath: string, video: string) {
  const launched = await launchShell({ onboardingSeen: true, videoDir: video, openFile: docPath })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  return { launched, app: launched.app, page }
}

/** whether the block holding `text` lies inside the scrolled view */
const inView = (page: Page, text: string) =>
  page.evaluate((text) => {
    const el = Array.from(document.querySelectorAll('.doc-page p, .doc-page h1')).find(
      (candidate) => candidate.textContent === text,
    )
    const scroller = document.querySelector('main.editor-scroll')
    if (!el || !scroller) return false
    const r = el.getBoundingClientRect()
    const s = scroller.getBoundingClientRect()
    return r.top >= s.top - 1 && r.bottom <= s.bottom + 1
  }, text)

const FILLER = (from: number, count: number) =>
  Array.from({ length: count }, (_, i) => para(`filler line ${from + i}`)).join('')

const LINK_BODY =
  para('intro') +
  FILLER(1, 60) +
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Chapter Two</w:t></w:r></w:p>' +
  FILLER(61, 60) +
  '<w:p><w:bookmarkStart w:id="7" w:name="figures"/><w:r><w:t>figure list</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>' +
  FILLER(121, 60) +
  para('the end')

test('Insert > Link reaches a heading, a bookmark, the top of the document and an e-mail address', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-link-dialog-')))
  const docPath = join(dir, 'links.docx')
  writeFileSync(docPath, await docxWithStyles(LINK_BODY))
  const mod = process.platform === 'darwin' ? 'Meta' : 'Control'
  let session = await openDoc(docPath, 'docs-link-paragraph-dialogs')
  try {
    let { app, page } = session
    const block = (text: string) =>
      page.locator('.doc-page p, .doc-page h1').filter({ hasText: new RegExp(`^${text}`) })
    const dialog = page.locator('.link-dialog')
    const tabOf = (name: string) => dialog.locator('[role="tab"]', { hasText: name })
    const place = (name: string) =>
      dialog.locator('[role="option"]').filter({ hasText: new RegExp(`^${name}`) })
    const link = (text: string) => page.locator('.doc-page a.doc-link', { hasText: text })
    const openDialogAtEndOf = async (text: string) => {
      await block(text).first().click()
      await page.keyboard.press('End')
      await page.keyboard.type(' ')
      await page.keyboard.press(`${mod}+k`)
      await expect(dialog).toBeVisible()
    }
    // the right-click menu puts the caret on the link before the dialog reads it
    const editLink = async (text: string) => {
      await link(text).click({ button: 'right' })
      await page.locator('.ctx-menu .ctx-item', { hasText: 'Edit Hyperlink…' }).click()
      await expect(dialog).toBeVisible()
    }
    const scrollToTop = async () => {
      await page.evaluate(() => document.querySelector('main.editor-scroll')!.scrollTo(0, 0))
      await expect.poll(() => inView(page, 'filler line 1')).toBe(true)
    }

    await expect(block('intro').first()).toBeVisible({ timeout: 30_000 })

    // a heading: listed under Headings, the link rides on a hidden bookmark put on it
    await openDialogAtEndOf('intro')
    await expect(tabOf('Web Page or File')).toHaveAttribute('aria-selected', 'true')
    await dialog.locator('input[name="linkText"]').fill('to chapter')
    await tabOf('This Document').click()
    await expect(place('Top of the Document')).toHaveCount(1)
    await expect(place('figures')).toHaveCount(1)
    await expect(dialog.locator('button.btn-primary')).toBeDisabled()
    await place('Chapter Two').click()
    await page.screenshot({ path: test.info().outputPath('link-this-document.png') })
    await dialog.locator('button.btn-primary').click()
    await expect(dialog).toBeHidden()
    await expect(link('to chapter')).toHaveAttribute('href', '#_Chapter_Two')

    // a bookmark
    await openDialogAtEndOf('intro')
    await dialog.locator('input[name="linkText"]').fill('to figures')
    await tabOf('This Document').click()
    await place('figures').click()
    await dialog.locator('button.btn-primary').click()
    await expect(link('to figures')).toHaveAttribute('href', '#figures')

    // an e-mail address with a subject
    await openDialogAtEndOf('intro')
    await dialog.locator('input[name="linkText"]').fill('mail us')
    await tabOf('Email Address').click()
    await dialog.locator('input[name="linkEmail"]').fill('team@example.com')
    await dialog.locator('input[name="linkSubject"]').fill('Hello there')
    await page.screenshot({ path: test.info().outputPath('link-email.png') })
    await dialog.locator('input[name="linkSubject"]').press('Enter')
    await expect(link('mail us')).toHaveAttribute(
      'href',
      'mailto:team@example.com?subject=Hello%20there',
    )

    // following the links moves to their targets
    expect(await inView(page, 'Chapter Two')).toBe(false)
    await link('to chapter').click({ modifiers: [mod] })
    await expect.poll(() => inView(page, 'Chapter Two')).toBe(true)
    await scrollToTop()
    await link('to figures').click({ modifiers: [mod] })
    await expect.poll(() => inView(page, 'figure list')).toBe(true)

    // the top of the document, from its end
    await block('the end')
      .first()
      .evaluate((el) => el.scrollIntoView({ block: 'center' }))
    await expect.poll(() => inView(page, 'the end')).toBe(true)
    await openDialogAtEndOf('the end')
    await dialog.locator('input[name="linkText"]').fill('back up')
    await tabOf('This Document').click()
    await place('Top of the Document').click()
    await dialog.locator('button.btn-primary').click()
    await expect(link('back up')).toHaveAttribute('href', '#_top')
    expect(await inView(page, 'filler line 1')).toBe(false)
    await link('back up').click({ modifiers: [mod] })
    await expect.poll(() => inView(page, 'filler line 1')).toBe(true)

    // an existing link opens on its tab with its values; a ScreenTip is added to it
    await editLink('to chapter')
    await expect(dialog.locator('h2')).toHaveText('Edit Hyperlink')
    await expect(tabOf('This Document')).toHaveAttribute('aria-selected', 'true')
    await expect(place('Chapter Two')).toHaveAttribute('aria-selected', 'true')
    await expect(dialog.locator('input[name="linkText"]')).toHaveValue('to chapter')
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await editLink('to figures')
    await expect(tabOf('This Document')).toHaveAttribute('aria-selected', 'true')
    await expect(place('figures')).toHaveAttribute('aria-selected', 'true')
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await editLink('mail us')
    await expect(tabOf('Email Address')).toHaveAttribute('aria-selected', 'true')
    await expect(dialog.locator('input[name="linkEmail"]')).toHaveValue('team@example.com')
    await expect(dialog.locator('input[name="linkSubject"]')).toHaveValue('Hello there')
    await dialog.locator('input[name="linkScreenTip"]').fill('Write to us')
    await dialog.locator('button.btn-primary').click()
    await expect(link('mail us')).toHaveAttribute('title', 'Write to us')

    // following the mail link hands it to the system's mail program
    await app.evaluate(({ shell }) => {
      const opened: string[] = []
      ;(globalThis as { __openedLinks?: string[] }).__openedLinks = opened
      shell.openExternal = async (url: string) => {
        opened.push(url)
      }
    })
    await link('mail us').click({ modifiers: [mod] })
    await expect
      .poll(() => app.evaluate(() => (globalThis as { __openedLinks?: string[] }).__openedLinks))
      .toEqual(['mailto:team@example.com?subject=Hello%20there'])

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await partOf(docPath, 'word/document.xml')
          const rels = await partOf(docPath, 'word/_rels/document.xml.rels')
          const heading = /<w:p>(?:(?!<\/w:p>).)*Chapter Two.*?<\/w:p>/.exec(xml)?.[0] ?? ''
          const mailRel =
            /Id="([^"]+)"[^>]*Target="mailto:team@example\.com\?subject=Hello%20there"[^>]*TargetMode="External"/.exec(
              rels,
            )?.[1]
          return [
            /<w:hyperlink w:anchor="_Chapter_Two">/.test(xml),
            /<w:bookmarkStart [^>]*w:name="_Chapter_Two"\/>/.test(heading),
            /<w:hyperlink w:anchor="figures">/.test(xml),
            /<w:hyperlink w:anchor="_top">/.test(xml),
            !!mailRel && xml.includes(`<w:hyperlink r:id="${mailRel}" w:tooltip="Write to us">`),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true])

    // read back: the links are there and the heading link still finds its heading
    await closeAndSaveVideo(session.launched, 'docs-link-paragraph-dialogs')
    session = await openDoc(docPath, 'docs-link-paragraph-dialogs')
    ;({ app, page } = session)
    const reopened = (text: string) => page.locator('.doc-page a.doc-link', { hasText: text })
    await expect(reopened('to chapter')).toHaveAttribute('href', '#_Chapter_Two', {
      timeout: 30_000,
    })
    await expect(reopened('mail us')).toHaveAttribute('title', 'Write to us')
    expect(await inView(page, 'Chapter Two')).toBe(false)
    await reopened('to chapter').click({ modifiers: [mod] })
    await expect.poll(() => inView(page, 'Chapter Two')).toBe(true)
    await reopened('mail us').click({ button: 'right' })
    await page.locator('.ctx-menu .ctx-item', { hasText: 'Edit Hyperlink…' }).click()
    const again = page.locator('.link-dialog')
    await expect(again.locator('[role="tab"]', { hasText: 'Email Address' })).toHaveAttribute(
      'aria-selected',
      'true',
    )
    await expect(again.locator('input[name="linkSubject"]')).toHaveValue('Hello there')
    await expect(again.locator('input[name="linkScreenTip"]')).toHaveValue('Write to us')
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-link-paragraph-dialogs')
    rmSync(dir, { recursive: true, force: true })
  }
})

/** x (pt) of the first text item on each page of a PDF */
async function pdfTextLefts(file: string): Promise<number[]> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(file)) }).promise
  const lefts: number[] = []
  for (let n = 1; n <= doc.numPages; n++) {
    const items = (await (await doc.getPage(n)).getTextContent()).items as Array<{
      str: string
      transform: number[]
    }>
    const first = items.find((item) => item.str.trim())
    if (first) lefts.push(first.transform[4])
  }
  return lefts
}

const GRID_SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/><w:docGrid w:type="lines" w:linePitch="480"/></w:sectPr>'

const LONG =
  'Internationalization responsibilities notwithstanding, counterrevolutionary characterizations ' +
  'misrepresent uncharacteristically straightforward interdepartmental communications.'

const PARA_BODY =
  para('apple') +
  `<w:p><w:pPr><w:adjustRightInd w:val="0"/></w:pPr><w:r><w:t>${LONG}</w:t></w:r></w:p>` +
  '<w:p><w:pPr><w:pageBreakBefore/><w:ind w:left="1440"/></w:pPr><w:r><w:t>banana</w:t></w:r></w:p>' +
  para('cherry') +
  GRID_SECT

test('Format > Paragraph: Mirror indents, hyphenation, the grid, Tabs… and Set As Default', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-paragraph-dialog-')))
  const docPath = join(dir, 'paragraph.docx')
  writeFileSync(docPath, await docxWithStyles(PARA_BODY, '<w:autoHyphenation/>'))
  let session = await openDoc(docPath, 'docs-link-paragraph-dialogs')
  try {
    const ui = (page: Page, app: ElectronApplication) => {
      const paragraph = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
      const dialog = page.locator('.modal', { has: page.locator('h2', { hasText: /^Paragraph$/ }) })
      const tabOf = (name: string) => dialog.locator('[role="tab"]', { hasText: name })
      const box = (name: string) => dialog.locator(`input[type="checkbox"][name="${name}"]`)
      const openParagraph = async (text: string) => {
        await paragraph(text).click()
        // the dialog reads the editor's selection, which follows the click a moment later
        await expect
          .poll(() =>
            page.evaluate((text) => {
              const editor = (window as unknown as AidocsWindow).__aidocs?.editor
              return editor?.state.selection.$from.parent.textContent.includes(text) ?? false
            }, text),
          )
          .toBe(true)
        await clickMenu(app, ['Format', 'Paragraph…'])
        await expect(dialog).toBeVisible()
      }
      const ok = async () => {
        await dialog.locator('button.btn-primary').click()
        await expect(dialog).toBeHidden()
      }
      const left = (text: string) =>
        paragraph(text).evaluate((el) => el.getBoundingClientRect().left)
      const height = (text: string) =>
        paragraph(text).evaluate((el) => el.getBoundingClientRect().height)
      // hyphens: whether the break points Docs places in the paragraph draw (the --doc-shy cascade)
      const css = (text: string, prop: 'hyphens' | 'textAlign') =>
        paragraph(text).evaluate((el, prop) => {
          const cs = getComputedStyle(el)
          if (prop !== 'hyphens') return cs[prop]
          const shy = cs.getPropertyValue('--doc-shy').trim()
          return shy && shy !== 'none' ? 'auto' : 'manual'
        }, prop)
      return { page, app, paragraph, dialog, tabOf, box, openParagraph, ok, left, css, height }
    }
    let { page, app, paragraph, dialog, tabOf, box, openParagraph, ok, left, css, height } = ui(
      session.page,
      session.app,
    )

    await expect(paragraph('banana')).toBeVisible({ timeout: 30_000 })

    // Mirror indents: banana is on page 2, so its left indent becomes the outside one
    const indent = (await left('banana')) - (await left('cherry'))
    expect(indent).toBeGreaterThan(50)
    await openParagraph('banana')
    await expect(dialog.getByLabel('Left indent')).toBeVisible()
    await expect(box('mirrorIndents')).not.toBeChecked()
    await box('mirrorIndents').check()
    await expect(dialog.getByLabel('Inside')).toBeVisible()
    await expect(dialog.getByLabel('Outside')).toBeVisible()
    // the right indent is not pulled onto a character grid here: the box is not offered
    await expect(box('adjustRightInd')).toHaveCount(0)
    await page.screenshot({ path: test.info().outputPath('paragraph-indents.png') })
    await ok()
    await expect
      .poll(async () => Math.abs((await left('banana')) - (await left('cherry'))))
      .toBeLessThan(1)

    // Don't hyphenate: the document hyphenates, this paragraph no longer does
    expect(await css('Internationalization', 'hyphens')).toBe('auto')
    await openParagraph('Internationalization')
    await tabOf('Line and Page Breaks').click()
    await expect(box('suppressAutoHyphens')).not.toBeChecked()
    await box('suppressAutoHyphens').check()
    await page.screenshot({ path: test.info().outputPath('paragraph-breaks.png') })
    await ok()
    await expect.poll(() => css('Internationalization', 'hyphens')).toBe('manual')
    expect(await css('apple', 'hyphens')).toBe('auto')

    // Snap to grid: off the 24pt line grid the one-line paragraph is shorter
    const snapped = await height('cherry')
    await openParagraph('cherry')
    await expect(box('snapToGrid')).toBeChecked()
    await box('snapToGrid').uncheck()
    await ok()
    await expect.poll(() => height('cherry')).toBeLessThan(snapped - 4)

    // Tabs… keeps what the dialog changed and opens the Tabs dialog
    await openParagraph('cherry')
    await tabOf('Line and Page Breaks').click()
    await box('keepNext').check()
    await dialog.locator('button', { hasText: /^Tabs…$/ }).click()
    await expect(dialog).toBeHidden()
    const tabs = page.locator('.tabs-dialog')
    await expect(tabs).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(tabs).toBeHidden()

    // Set As Default: the default paragraph style takes what was changed, every paragraph follows
    expect(await css('banana', 'textAlign')).not.toBe('justify')
    await openParagraph('apple')
    await dialog.locator('button.gs-dd-btn[aria-label="Alignment"]').click()
    await page.locator('.gs-dd-pop [role="option"][aria-label="Justified"]').click()
    await dialog.getByLabel('Space after').fill('12')
    await dialog.locator('button', { hasText: /^Set as Default$/ }).click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => css('banana', 'textAlign')).toBe('justify')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await partOf(docPath, 'word/document.xml')
          const styles = await partOf(docPath, 'word/styles.xml')
          const p = (text: string) =>
            xml
              .split('<w:p>')
              .find((chunk) => chunk.includes(`>${text}`))
              ?.split('</w:pPr>')[0] ?? ''
          const normal = /<w:style [^>]*w:styleId="Normal">.*?<\/w:style>/.exec(styles)?.[0] ?? ''
          return [
            /<w:ind w:left="1440"\/><w:mirrorIndents\/>/.test(p('banana')),
            /<w:suppressAutoHyphens\/><w:adjustRightInd w:val="0"\/>/.test(
              p('Internationalization'),
            ),
            /<w:keepNext\/><w:snapToGrid w:val="0"\/>/.test(p('cherry')),
            /<w:spacing [^>]*w:after="240"/.test(normal),
            /<w:jc w:val="both"\/>/.test(normal),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true, true])

    // the exported PDF agrees with the screen: on page 2 banana starts at the margin
    const pdfPath = join(dir, 'paragraph.pdf')
    await app.evaluate(({ dialog: d, shell }, out) => {
      d.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    const lefts = await pdfTextLefts(pdfPath)
    expect(lefts).toHaveLength(2)
    expect(lefts[1]).toBeCloseTo(lefts[0], 0)

    // read back
    await closeAndSaveVideo(session.launched, 'docs-link-paragraph-dialogs')
    session = await openDoc(docPath, 'docs-link-paragraph-dialogs')
    ;({ page, app, paragraph, dialog, tabOf, box, openParagraph, ok, left, css, height } = ui(
      session.page,
      session.app,
    ))
    await expect(paragraph('banana')).toBeVisible({ timeout: 30_000 })
    await expect
      .poll(async () => Math.abs((await left('banana')) - (await left('cherry'))))
      .toBeLessThan(1)
    expect(await css('Internationalization', 'hyphens')).toBe('manual')
    expect(await css('banana', 'textAlign')).toBe('justify')
    await openParagraph('banana')
    await expect(box('mirrorIndents')).toBeChecked()
    await expect(dialog.getByLabel('Inside')).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await openParagraph('cherry')
    await expect(box('snapToGrid')).not.toBeChecked()
    await tabOf('Line and Page Breaks').click()
    await expect(box('keepNext')).toBeChecked()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    await openParagraph('Internationalization')
    await tabOf('Line and Page Breaks').click()
    await expect(box('suppressAutoHyphens')).toBeChecked()
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-link-paragraph-dialogs')
    rmSync(dir, { recursive: true, force: true })
  }
})
