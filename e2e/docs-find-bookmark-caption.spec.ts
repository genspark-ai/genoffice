/**
 * The Bookmark, Cross-reference and Caption dialogs and the Find panel after
 * they took on more of Word's controls: bookmark sorting, hidden bookmarks and
 * Go To; references to headings, page numbers and above/below; caption labels,
 * number formats and the label left out; a search kept to the selection; and
 * the Go To kinds Field, Graphic, Equation and Object. Checked in the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-find-bookmark-caption'
const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>'
const HEADING =
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Scope</w:t></w:r></w:p>'

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

/** puts the caret at the end of the paragraph starting with `text`, or selects all of it */
const caretIn = (page: Page, text: string, whole = false) =>
  page.evaluate(
    ([text, whole]) => {
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
        .setTextSelection(whole ? { from, to } : to)
        .run()
    },
    [text, whole] as const,
  )

const selectedText = (page: Page) =>
  page.evaluate(() => {
    const { state } = (window as unknown as EditorWindow).__aidocs.editor
    return state.doc.textBetween(state.selection.from, state.selection.to)
  })

const caretBlock = (page: Page) =>
  page.evaluate(() => {
    const { state } = (window as unknown as EditorWindow).__aidocs.editor
    return state.doc.resolve(state.selection.from).parent.textContent
  })

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')?.async('string') ?? ''

test('bookmarks sort and go to, references point at headings and pages, captions take labels and formats, find keeps to the selection', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-find-bookmark-caption-')))
  const docPath = join(dir, 'refs.docx')
  writeFileSync(
    docPath,
    await docxWith(
      HEADING +
        para('alpha one alpha') +
        para('alpha two') +
        para('Target paragraph') +
        para('See also:') +
        para('Closing line') +
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
    const line = (text: string) => page.locator('.doc-page p', { hasText: text }).first()
    await expect(line('Target paragraph')).toBeVisible({ timeout: 30_000 })

    // ---- Bookmark: Add, Sort by, Go To, Delete ----
    const bookmark = page.locator('.bookmark-dialog')
    const addBookmark = async (at: string, name: string) => {
      await caretIn(page, at)
      await clickMenu(app, ['Insert', 'Bookmark…'])
      await bookmark.locator('input[name="bookmark-name"]').fill(name)
      await bookmark.locator('button[name="add"]').click()
      await expect(bookmark.locator('.bookmark-row', { hasText: name })).toHaveCount(1)
    }
    await addBookmark('Target paragraph', 'Able')
    await page.keyboard.press('Escape')
    await expect(bookmark).toBeHidden()
    await addBookmark('alpha two', 'Zed')
    const names = bookmark.locator('.bookmark-row .bookmark-name')
    await expect(bookmark.locator('input[value="name"]')).toBeChecked()
    await expect(names).toHaveText(['Able', 'Zed'])
    await bookmark.locator('input[value="location"]').check()
    await expect(names).toHaveText(['Zed', 'Able'])
    await expect(bookmark.locator('button[name="go-to"]')).toBeEnabled()
    await bookmark.locator('.bookmark-row', { hasText: 'Able' }).click()
    await page.screenshot({ path: test.info().outputPath('bookmark-dialog.png') })
    await bookmark.locator('button[name="go-to"]').click()
    await expect.poll(() => selectedText(page)).toBe('Target paragraph')
    // a name that is not a bookmark has nothing to go to or delete
    await bookmark.locator('input[name="bookmark-name"]').fill('Nobody')
    await expect(bookmark.locator('button[name="go-to"]')).toBeDisabled()
    await expect(bookmark.locator('button[name="delete"]')).toBeDisabled()
    await page.keyboard.press('Escape')

    // ---- Cross-reference: a heading as plain text, a bookmark as above/below and as a page ----
    const crossRef = page.locator('.crossref-dialog')
    const openCrossRef = async () => {
      await caretIn(page, 'See also:')
      await clickMenu(app, ['Insert', 'Cross-reference…'])
      await expect(crossRef).toBeVisible()
    }
    await openCrossRef()
    await expect(crossRef.locator('select[name="ref-type"]')).toHaveValue('heading')
    await expect(crossRef.locator('option[value="numbered"]')).toHaveCount(1)
    for (const kind of ['footnote', 'endnote'])
      await expect(crossRef.locator(`option[value="${kind}"]`)).not.toHaveAttribute('disabled', '')
    await expect(crossRef.locator('input[name="ref-separator"]')).toBeDisabled()
    await expect(crossRef.locator('input[name="ref-above-below"]')).toBeDisabled()
    await expect(crossRef.locator('.bookmark-row')).toHaveText(['Scope'])
    await crossRef.locator('input[name="ref-hyperlink"]').uncheck()
    await page.screenshot({ path: test.info().outputPath('crossref-dialog.png') })
    await crossRef.locator('button[name="insert"]').click()
    await expect(page.locator('.doc-page .doc-ref-field')).toHaveText(['Scope'])

    await openCrossRef()
    await crossRef.locator('select[name="ref-type"]').selectOption('bookmark')
    await crossRef.locator('select[name="ref-content"]').selectOption('aboveBelow')
    await crossRef.locator('.bookmark-row', { hasText: 'Able' }).click()
    await crossRef.locator('button[name="insert"]').click()
    await expect(page.locator('.doc-page .doc-ref-field')).toHaveText(['Scope', 'above'])

    await openCrossRef()
    await crossRef.locator('select[name="ref-type"]').selectOption('bookmark')
    await crossRef.locator('select[name="ref-content"]').selectOption('page')
    await crossRef.locator('input[name="ref-above-below"]').check()
    await crossRef.locator('.bookmark-row', { hasText: 'Able' }).click()
    await crossRef.locator('button[name="insert"]').click()
    // the bookmark is on this page, further up
    await expect(line('See also:')).toHaveText('See also:Scopeaboveabove')

    // the heading reference left a hidden bookmark behind
    await clickMenu(app, ['Insert', 'Bookmark…'])
    await expect(bookmark.locator('.bookmark-row')).toHaveCount(2)
    await bookmark.locator('input[name="bookmark-hidden"]').check()
    await expect(bookmark.locator('.bookmark-row')).toHaveCount(3)
    await expect(bookmark.locator('.bookmark-row .bookmark-name').first()).toHaveText(/^_Ref\d+$/)
    await page.keyboard.press('Escape')

    // ---- Caption: a new label, roman numbers, then the label left out ----
    const caption = page.locator('.caption-dialog')
    await caretIn(page, 'Closing line')
    await clickMenu(app, ['Insert', 'Caption…'])
    await expect(caption.locator('select[name="caption-position"]')).toBeDisabled()
    await expect(caption.locator('button[name="auto-caption"]')).toBeEnabled()
    await expect(caption.locator('button[name="delete-label"]')).toBeDisabled()
    await caption.locator('button[name="new-label"]').click()
    await page.locator('.modal input').fill('Plate')
    await page.keyboard.press('Enter')
    await expect(caption.locator('select[name="caption-label"]')).toHaveValue('Plate')
    await expect(caption.locator('button[name="delete-label"]')).toBeEnabled()
    await caption.locator('select[name="caption-format"]').selectOption('ROMAN')
    await caption.locator('input[name="caption-text"]').fill('Map')
    await page.screenshot({ path: test.info().outputPath('caption-dialog.png') })
    await caption.locator('button[name="ok"]').click()
    await expect(page.locator('.doc-page', { hasText: 'Plate I Map' })).toHaveCount(1)

    await caretIn(page, 'Closing line')
    await clickMenu(app, ['Insert', 'Caption…'])
    await caption.locator('select[name="caption-label"]').selectOption('Plate')
    // the label's captions are roman already
    await expect(caption.locator('select[name="caption-format"]')).toHaveValue('ROMAN')
    await caption.locator('input[name="caption-exclude-label"]').check()
    await caption.locator('input[name="caption-text"]').fill('Chart')
    await caption.locator('button[name="ok"]').click()
    // it went in above the first one, so it is number I and the first one moves to II
    await expect(page.locator('.doc-page .doc-protected-field-text')).toHaveText([
      'I Chart',
      'Plate II Map',
    ])

    await caretIn(page, 'Closing line')
    await clickMenu(app, ['Insert', 'Caption…'])
    await caption.locator('select[name="caption-label"]').selectOption('Plate')
    await caption.locator('button[name="delete-label"]').click()
    // still listed: the document uses it
    await expect(caption.locator('option[value="Plate"]')).toHaveCount(1)
    await expect(caption.locator('button[name="delete-label"]')).toBeDisabled()
    await page.keyboard.press('Escape')

    // ---- Find: kept to the selection, Replace All stays inside it ----
    await caretIn(page, 'alpha one alpha', true)
    await expect.poll(() => selectedText(page)).toBe('alpha one alpha')
    await clickMenu(app, ['Edit', 'Find', 'Replace…'])
    const panel = page.locator('.find-panel')
    await expect(panel.locator('.find-tab.on')).toHaveText('Replace')
    const scope = panel.locator('select[name="find-scope"]')
    await expect(scope).toHaveValue('main')
    await panel.locator('.find-input').first().fill('alpha')
    await expect(panel.locator('.find-count')).toHaveText('1/3')
    await scope.selectOption('selection')
    await expect(panel.locator('.find-count')).toHaveText('1/2')
    await page.screenshot({ path: test.info().outputPath('find-scope.png') })
    await panel.locator('.find-input').nth(1).fill('beta')
    await panel.locator('.find-action', { hasText: 'Replace All' }).click()
    await expect(line('beta one beta')).toBeVisible()
    await expect(line('alpha two')).toBeVisible()
    await scope.selectOption('main')
    await expect(panel.locator('.find-count')).toHaveText('1/1')

    // ---- Go To: Word's thirteen kinds; a field is found, an equation is not there ----
    await panel.locator('.find-tab', { hasText: 'Go To' }).click()
    const what = panel.locator('.find-select').first()
    await expect(what.locator('option')).toHaveText([
      'Page',
      'Section',
      'Line',
      'Bookmark',
      'Comment',
      'Footnote',
      'Endnote',
      'Field',
      'Table',
      'Graphic',
      'Equation',
      'Object',
      'Heading',
    ])
    await caretIn(page, 'Scope')
    await what.selectOption('field')
    await panel.locator('.find-goto-input').fill('1')
    await panel.locator('.find-action-primary').click()
    await expect.poll(() => caretBlock(page)).toBe('See also:Scopeaboveabove')
    await what.selectOption('equation')
    await panel.locator('.find-action-primary').click()
    await expect(panel.locator('.find-hint-miss')).toHaveCount(1)
    await page.screenshot({ path: test.info().outputPath('goto-kinds.png') })

    // the save takes the page layout as it stands: wait for the one of this document
    await page.waitForFunction(
      () =>
        (
          window as unknown as { __aidocs: { pagesSettled: () => boolean } }
        ).__aidocs.pagesSettled(),
      undefined,
      { timeout: 30_000 },
    )
    await saveFromMenu(app)
    await expect
      .poll(() => documentXml(docPath), { timeout: 20_000 })
      .toContain('SEQ Plate \\* ROMAN')
    const xml = await documentXml(docPath)
    // bookmarks hold their paragraph's text
    expect(xml).toMatch(
      /<w:bookmarkStart w:id="(\d+)" w:name="Able"\/><w:r>(?:(?!<\/w:p>).)*Target paragraph<\/w:t><\/w:r><w:bookmarkEnd w:id="\1"\/>/,
    )
    expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="Zed"\/>/)
    const ref =
      /<w:bookmarkStart w:id="\d+" w:name="(_Ref\d+)"\/><w:r>(?:(?!<\/w:p>).)*Scope</.exec(xml)?.[1]
    expect(ref).toBeTruthy()
    // not a hyperlink: no \h
    expect(xml).toContain(`> REF ${ref} </w:instrText>`)
    expect(xml).toContain('> REF Able \\p \\h </w:instrText>')
    expect(xml).toMatch(/>\s*PAGEREF Able \\p \\h\s*<\/w:instrText>/)
    expect(xml.match(/SEQ Plate \\\* ROMAN/g)).toHaveLength(2)
    expect(xml).toContain('>Plate </w:t>')
    // the numbers are saved in document order as the fields' results
    expect(
      [...xml.matchAll(/SEQ Plate[^<]*<\/w:instrText>[\s\S]*?<w:t[^>]*>([^<]*)<\/w:t>/g)].map(
        (m) => m[1],
      ),
    ).toEqual(['I', 'II'])
    // the page reference carries what the layout shows (the bookmark is further up this page): nothing is left for Word to update
    expect(xml).not.toContain('w:dirty')
    expect(xml).toMatch(
      /<w:instrText[^>]*>\s*PAGEREF Able \\p \\h\s*<\/w:instrText>(?:(?!<\/w:p>).)*?<w:t[^>]*>above<\/w:t>/,
    )
    expect(xml).toContain('beta one beta')
    expect(xml).toContain('alpha two')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
