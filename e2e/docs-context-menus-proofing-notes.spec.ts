/**
 * The right-click menus of four scenes, with real right-clicks on a file that
 * was never typed in: a misspelled word (suggestions, Ignore Once, Ignore All,
 * AutoCorrect, Spelling…), a comment (its balloon and its anchor text), a
 * footnote (its text and its reference mark) and a tracked change. Each
 * checks the document and, where the action reaches the file, the saved file.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml'
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
const para = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`
const WHO = 'w:author="Reviewer" w:date="2026-01-05T10:00:00Z"'

async function fixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${TYPE}.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="${TYPE}.comments+xml"/><Override PartName="/word/footnotes.xml" ContentType="${TYPE}.footnotes+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/comments" Target="comments.xml"/><Relationship Id="rId2" Type="${REL}/footnotes" Target="footnotes.xml"/></Relationships>`,
  )
  zip.file(
    'word/comments.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:comments ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="1" ${WHO} w:initials="R"><w:p w14:paraId="1A2B3C4D">${run('A comment to act on')}</w:p></w:comment></w:comments>`,
  )
  zip.file(
    'word/footnotes.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:footnotes ${W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>${run(' The note to act on')}</w:p></w:footnote></w:footnotes>`,
  )
  const body =
    para(run('First the mispeled word.')) +
    para(run('Tracked the wrnog way.')) +
    para(run('Both qzxvw here and qzxvw there.')) +
    para(run('Once plmokn here and plmokn there.')) +
    para(run('Paired acheive entry.')) +
    para(
      run('Some '),
      '<w:commentRangeStart w:id="1"/>',
      run('commented phrase'),
      '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>',
      run(' here.'),
    ) +
    para(
      run('Text with '),
      `<w:ins w:id="11" ${WHO}>${run('inserted words')}</w:ins>`,
      run(' and '),
      `<w:del w:id="12" ${WHO}><w:r><w:delText xml:space="preserve">deleted words</w:delText></w:r></w:del>`,
      run(' end.'),
    ) +
    para(run('A note mark'), '<w:r><w:footnoteReference w:id="1"/></w:r>', run(' in a line.')) +
    para(run('Last line'))
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${body}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

const part = async (file: string, name: string): Promise<string> => {
  const zip = await JSZip.loadAsync(readFileSync(file))
  return (await zip.file(name)?.async('string')) ?? ''
}

interface Scene {
  launched: LaunchedApp
  app: ElectronApplication
  page: Page
  docPath: string
  dir: string
}

async function openScene(name: string): Promise<Scene> {
  const dir = realpathSync(mkdtempSync(join(process.env.TMPDIR ?? tmpdir(), 'genoffice-e2e-ctx-')))
  const userDataDir = join(dir, 'user-data')
  mkdirSync(userDataDir)
  const docPath = join(dir, 'scenes.docx')
  writeFileSync(docPath, await fixture())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: name,
    openFile: docPath,
    userDataDir,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )
  await expect(page.locator('.doc-page')).toContainText('Last line', { timeout: 30_000 })
  return { launched, app: launched.app, page, docPath, dir }
}

async function closeScene(scene: Scene, name: string) {
  try {
    await closeAndSaveVideo(scene.launched, name)
  } finally {
    rmSync(scene.dir, { recursive: true, force: true })
  }
}

/** viewport centre of the nth place `needle` shows in the page */
const centreOf = (page: Page, needle: string, nth = 0) =>
  page.evaluate(
    ({ needle, nth }) => {
      const root = document.querySelector('.doc-live') ?? document.body
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
      let seen = 0
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const text = node.textContent ?? ''
        for (let at = text.indexOf(needle); at !== -1; at = text.indexOf(needle, at + 1)) {
          const range = document.createRange()
          range.setStart(node, at)
          range.setEnd(node, at + needle.length)
          const box = range.getBoundingClientRect()
          if (box.width === 0) continue
          if (seen++ === nth) return { x: box.left + box.width / 2, y: box.top + box.height / 2 }
        }
      }
      return null
    },
    { needle, nth },
  )

/** where the text is once the page has stopped moving (a markup column coming or going shifts it) */
async function settledCentreOf(page: Page, needle: string, nth = 0) {
  let last = await centreOf(page, needle, nth)
  await expect(async () => {
    await page.waitForTimeout(150)
    const now = await centreOf(page, needle, nth)
    const still = !!now && !!last && now.x === last.x && now.y === last.y
    last = now
    expect(still, `"${needle}" is on the page and at rest`).toBe(true)
  }).toPass({ timeout: 15_000 })
  return last!
}

async function rightClickText(page: Page, needle: string, nth = 0) {
  const at = await settledCentreOf(page, needle, nth)
  await page.mouse.click(at.x, at.y, { button: 'right' })
  await page.locator('.ctx-menu').waitFor()
}

/** the top-level rows of the open menu, separators as --- */
const menuRows = (page: Page) =>
  page.evaluate(() =>
    [...document.querySelectorAll('.ctx-menu > *')].map((el) =>
      el.classList.contains('ctx-sep')
        ? '---'
        : (el.querySelector(':scope > .ctx-item .ctx-label, :scope > .ctx-label')?.textContent ??
          ''),
    ),
  )

const menuItem = (page: Page, label: string) =>
  page.locator('.ctx-menu > .ctx-item, .ctx-menu > .ctx-item-wrap > .ctx-item', {
    has: page.locator('.ctx-label', { hasText: new RegExp(`^${label}$`) }),
  })

/** a click in the page's corner: after a spelling item the recheck can swallow an Escape */
const closeMenu = async (page: Page) => {
  const box = (await page.locator('.doc-page').boundingBox())!
  await page.mouse.click(box.x + 10, box.y + 10)
  await expect(page.locator('.ctx-menu')).toHaveCount(0)
}

/** the recheck after a spelling item types a space and takes it back, deaf to clicks meanwhile */
const settle = (page: Page) => page.waitForTimeout(2000)

const bodyText = (page: Page) =>
  page.evaluate(() => {
    type Hook = { editor: { state: { doc: { textContent: string } } } }
    return (window as unknown as { __aidocs: Hook }).__aidocs.editor.state.doc.textContent
  })

/** the text carrying a mark of that name, piece by piece */
const marked = (page: Page, mark: string) =>
  page.evaluate((mark) => {
    type Node = { isText: boolean; text?: string; marks: { type: { name: string } }[] }
    type Hook = { editor: { state: { doc: { descendants(f: (n: Node) => void): void } } } }
    const out: string[] = []
    ;(window as unknown as { __aidocs: Hook }).__aidocs.editor.state.doc.descendants((node) => {
      if (node.isText && node.marks.some((m) => m.type.name === mark)) out.push(node.text ?? '')
    })
    return out
  }, mark)

const noteMarks = (page: Page) =>
  page.evaluate(() => {
    type Node = { type: { name: string }; attrs: { kind: string; id: string } }
    type Hook = { editor: { state: { doc: { descendants(f: (n: Node) => void): void } } } }
    const out: string[] = []
    ;(window as unknown as { __aidocs: Hook }).__aidocs.editor.state.doc.descendants((node) => {
      if (node.type.name === 'docNoteRef') out.push(`${node.attrs.kind}:${node.attrs.id}`)
    })
    return out
  })

const UNDO = process.platform === 'darwin' ? 'Meta+z' : 'Control+z'

test('a misspelled word of an opened file: suggestions, Ignore Once, Ignore All, AutoCorrect, Spelling…', async () => {
  test.setTimeout(240_000)
  const scene = await openScene('docs-ctx-spelling')
  const { page, app, docPath } = scene
  try {
    const checker = await page.evaluate(() => {
      const api = (
        window as unknown as {
          desktop: {
            spellIsMisspelled?(w: string): boolean
            spellSuggestions?(w: string): string[]
          }
        }
      ).desktop
      return {
        wrong: api.spellIsMisspelled?.('mispeled') === true,
        offers: (w: string) => api.spellSuggestions?.(w) ?? [],
        first: {
          mispeled: api.spellSuggestions?.('mispeled')?.[0],
          wrnog: api.spellSuggestions?.('wrnog')?.[0],
          acheive: api.spellSuggestions?.('acheive')?.[0],
        },
      }
    })
    test.skip(
      !checker.wrong || !checker.first.mispeled || !checker.first.wrnog || !checker.first.acheive,
      'the spelling checker has no dictionary in this environment',
    )
    const fix = checker.first as Record<'mispeled' | 'wrnog' | 'acheive', string>

    // the word was never typed, so the browser has no marker on it: the menu still knows it
    await rightClickText(page, 'mispeled')
    const rows = await menuRows(page)
    const cut = rows.indexOf('Cut')
    expect(rows[0]).toBe(fix.mispeled)
    expect(rows.slice(rows.indexOf('---') + 1, cut - 1).slice(0, 5)).toEqual([
      'Ignore Once',
      'Ignore All',
      'Add to Dictionary',
      'AutoCorrect',
      'Spelling…',
    ])
    await page.screenshot({ path: test.info().outputPath('misspelled.png') })
    await page.locator('.ctx-menu .ctx-item-strong').first().click()
    await expect.poll(() => bodyText(page)).toContain(`First the ${fix.mispeled} word.`)
    // one undo step takes it back, and the menu comes back with it
    await expect(page.locator('.doc-page')).toBeFocused()
    await page.keyboard.press(UNDO)
    await expect.poll(() => bodyText(page)).toContain('First the mispeled word.')

    // a word that is spelled right has no spelling section
    await rightClickText(page, 'First')
    expect((await menuRows(page))[0]).toBe('Cut')
    await closeMenu(page)

    // Spelling… opens the dialog on the word that was right-clicked
    await rightClickText(page, 'acheive')
    await menuItem(page, 'Spelling…').click()
    const dialog = page.locator('.spelling-dialog')
    await expect(dialog.locator('.spelling-context b')).toHaveText('acheive')
    await dialog.getByRole('button', { name: 'Close' }).click()
    await expect(dialog).toBeHidden()

    // Ignore Once: that occurrence only
    await rightClickText(page, 'plmokn', 0)
    await menuItem(page, 'Ignore Once').click()
    await settle(page)
    await rightClickText(page, 'plmokn', 0)
    expect((await menuRows(page))[0]).toBe('Cut')
    await closeMenu(page)
    await rightClickText(page, 'plmokn', 1)
    expect(await menuRows(page)).toContain('Ignore Once')
    await closeMenu(page)

    // Ignore All: every occurrence
    await rightClickText(page, 'qzxvw', 0)
    await menuItem(page, 'Ignore All').click()
    await settle(page)
    await expect
      .poll(async () => {
        await rightClickText(page, 'qzxvw', 1)
        const first = (await menuRows(page))[0]
        await closeMenu(page)
        return first
      })
      .toBe('Cut')
    await rightClickText(page, 'qzxvw', 0)
    expect((await menuRows(page))[0]).toBe('Cut')
    await closeMenu(page)

    // AutoCorrect: the pair corrects the word now and the same typing later
    await rightClickText(page, 'acheive')
    await menuItem(page, 'AutoCorrect').hover()
    const pair = page.locator('.ctx-submenu .ctx-item', {
      has: page.locator('.ctx-label', { hasText: new RegExp(`^${fix.acheive}$`) }),
    })
    await page.screenshot({ path: test.info().outputPath('autocorrect.png') })
    await pair.click()
    await expect.poll(() => bodyText(page)).toContain(`Paired ${fix.acheive} entry.`)
    const last = await settledCentreOf(page, 'Last line')
    await page.mouse.click(last.x, last.y)
    await page.keyboard.press(process.platform === 'darwin' ? 'Meta+ArrowRight' : 'End')
    await page.keyboard.type(' acheive again', { delay: 20 })
    await expect.poll(() => bodyText(page)).toContain(`Last line ${fix.acheive} again`)

    // with Track Changes on, a suggestion is a tracked replacement
    await page.locator('.status-track').click()
    await expect(page.locator('.status-track')).toHaveAttribute('aria-pressed', 'true')
    await rightClickText(page, 'wrnog')
    await page.locator('.ctx-menu .ctx-item-strong').first().click()
    await expect.poll(() => marked(page, 'ins')).toContain(fix.wrnog)
    expect(await marked(page, 'del')).toContain('wrnog')

    await saveFromMenu(app)
    await expect
      .poll(async () => await part(docPath, 'word/document.xml'), { timeout: 20_000 })
      .toContain(`Paired ${fix.acheive} entry.`)
    const xml = await part(docPath, 'word/document.xml')
    expect(xml).toMatch(new RegExp(`<w:ins [^>]*>(?:(?!</w:ins>).)*${fix.wrnog}`))
    expect(xml).toMatch(/<w:del [^>]*>(?:(?!<\/w:del>).)*<w:delText[^>]*>wrnog/)
    expect(xml).toContain('First the mispeled word.')
  } finally {
    await closeScene(scene, 'docs-ctx-spelling')
  }
})

test('a comment: the balloon gets Reply, Resolve and Delete; the anchor text lists the same three, with Reopen', async () => {
  test.setTimeout(240_000)
  const scene = await openScene('docs-ctx-comment')
  const { page, app, docPath } = scene
  try {
    const balloon = page.locator('.comment-bubble[data-comment-id]').first()
    await expect(balloon).toBeVisible({ timeout: 20_000 })

    // the balloon, also after it was picked with a click: the three items alone
    await balloon.click()
    await balloon.click({ button: 'right' })
    await page.locator('.ctx-menu').waitFor()
    expect(await menuRows(page)).toEqual(['Reply To Comment', 'Delete Comment', 'Resolve Comment'])
    await page.screenshot({ path: test.info().outputPath('balloon.png') })

    await menuItem(page, 'Reply To Comment').click()
    const reply = page.locator('.comment-reply-input:visible').first()
    await reply.click()
    await reply.fill('A reply from the menu')
    await reply.press('Enter')
    await expect(page.locator('.app')).toContainText('A reply from the menu')

    // the anchor text keeps the text menu, with the comment items in it
    await rightClickText(page, 'commented phrase')
    const rows = await menuRows(page)
    expect(rows[0]).toBe('Cut')
    // Word's text menu: the three comment items between the clipboard and Font…
    expect(rows.slice(0, rows.indexOf('Font…'))).toEqual([
      'Cut',
      'Copy',
      'Paste',
      'Paste Special…',
      '---',
      'Reply To Comment',
      'Delete Comment',
      'Resolve Comment',
      '---',
    ])
    for (const gone of ['Paste as Plain Text', 'Edit comment', 'Track Changes'])
      expect(rows).not.toContain(gone)
    await page.screenshot({ path: test.info().outputPath('anchor.png') })
    await closeMenu(page)

    // Resolve from the balloon, Reopen from the anchor text
    await balloon.click({ button: 'right' })
    await menuItem(page, 'Resolve Comment').click()
    await rightClickText(page, 'commented phrase')
    await expect(menuItem(page, 'Reopen Comment')).toBeEnabled()
    await expect(menuItem(page, 'Reply To Comment')).toBeDisabled()
    await menuItem(page, 'Reopen Comment').click()
    await rightClickText(page, 'commented phrase')
    await expect(menuItem(page, 'Resolve Comment')).toBeEnabled()
    await closeMenu(page)

    await saveFromMenu(app)
    await expect
      .poll(async () => await part(docPath, 'word/comments.xml'), { timeout: 20_000 })
      .toContain('A reply from the menu')
    expect(await part(docPath, 'word/comments.xml')).toContain('A comment to act on')

    // Delete from the balloon: the thread and its anchor go
    await expect(balloon).toBeVisible()
    await balloon.click({ button: 'right' })
    await menuItem(page, 'Delete Comment').click()
    await expect(page.locator('.doc-page .doc-comment')).toHaveCount(0)
    await expect(page.locator('.comment-bubble[data-comment-id]')).toHaveCount(0)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).includes('commentRangeStart'), {
        timeout: 20_000,
      })
      .toBe(false)
    expect(await part(docPath, 'word/comments.xml')).not.toContain('A comment to act on')
  } finally {
    await closeScene(scene, 'docs-ctx-comment')
  }
})

test('a footnote: its text and its mark go to each other, open the options and convert', async () => {
  test.setTimeout(240_000)
  const scene = await openScene('docs-ctx-note')
  const { page, app, docPath } = scene
  try {
    const noteText = page.locator('[data-note-row]', { hasText: 'The note to act on' }).first()
    await expect(noteText).toBeVisible({ timeout: 20_000 })
    const mark = page.locator('.doc-page .doc-note-ref').first()

    // the note's text: the three note items alone
    await noteText.locator('.page-note-text, span').first().click({ button: 'right' })
    await page.locator('.ctx-menu').waitFor()
    expect(await menuRows(page)).toEqual(['Go to Footnote', 'Note Options…', 'Convert to Endnote'])
    await page.screenshot({ path: test.info().outputPath('note-text.png') })
    await menuItem(page, 'Go to Footnote').click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          type Hook = { editor: { state: { selection: { node?: { type: { name: string } } } } } }
          const { selection } = (window as unknown as { __aidocs: Hook }).__aidocs.editor.state
          return selection.node?.type.name ?? null
        }),
      )
      .toBe('docNoteRef')

    // the mark in the body: the text menu with the note items in it
    await mark.click({ button: 'right' })
    await page.locator('.ctx-menu').waitFor()
    const rows = await menuRows(page)
    expect(rows[0]).toBe('Cut')
    const at = rows.indexOf('Go to Footnote')
    expect(rows.slice(at, at + 3)).toEqual([
      'Go to Footnote',
      'Note Options…',
      'Convert to Endnote',
    ])
    await page.screenshot({ path: test.info().outputPath('note-mark.png') })
    await menuItem(page, 'Go to Footnote').click()
    const prompt = page.locator('.modal', { hasText: 'Edit footnote' })
    await expect(prompt.locator('textarea, input').first()).toHaveValue(/The note to act on/)
    await prompt.locator('textarea, input').first().focus()
    await page.keyboard.press('Escape')
    await expect(prompt).toBeHidden()

    await mark.click({ button: 'right' })
    await menuItem(page, 'Note Options…').click()
    const options = page.locator('.footnote-dialog')
    await expect(options).toBeVisible()
    await options.getByRole('button', { name: 'Cancel' }).click()
    await expect(options).toBeHidden()

    // Convert to Endnote from the text, and the file takes it
    await noteText.locator('.page-note-text, span').first().click({ button: 'right' })
    await menuItem(page, 'Convert to Endnote').click()
    await expect.poll(() => noteMarks(page)).toEqual(['endnote:1'])
    const endnote = page.locator('[data-note-row="endnote"]', { hasText: 'The note to act on' })
    await expect(endnote.first()).toBeVisible()
    await saveFromMenu(app)
    await expect
      .poll(
        async () => (await part(docPath, 'word/document.xml')).includes('<w:endnoteReference'),
        {
          timeout: 20_000,
        },
      )
      .toBe(true)
    expect(await part(docPath, 'word/document.xml')).not.toContain('<w:footnoteReference')
    expect(await part(docPath, 'word/endnotes.xml')).toContain('The note to act on')

    // and back again from the endnote's mark
    await mark.click({ button: 'right' })
    await menuItem(page, 'Convert to Footnote').click()
    await expect.poll(() => noteMarks(page)).toEqual(['footnote:1'])
    await saveFromMenu(app)
    await expect
      .poll(
        async () => (await part(docPath, 'word/document.xml')).includes('<w:footnoteReference'),
        {
          timeout: 20_000,
        },
      )
      .toBe(true)
    expect(await part(docPath, 'word/footnotes.xml')).toContain('The note to act on')
  } finally {
    await closeScene(scene, 'docs-ctx-note')
  }
})

test('a tracked change is named at the top of the menu, and Tools > Translate Selection takes the caret word', async () => {
  test.setTimeout(240_000)
  const scene = await openScene('docs-ctx-tracked')
  const { page, app, docPath } = scene
  try {
    await rightClickText(page, 'inserted words')
    expect((await menuRows(page)).slice(0, 4)).toEqual([
      'Accept Insertion',
      'Reject Insertion',
      '---',
      'Cut',
    ])
    await page.screenshot({ path: test.info().outputPath('tracked-insert.png') })
    await menuItem(page, 'Accept Insertion').click()
    await expect.poll(() => marked(page, 'ins')).toEqual([])
    expect(await bodyText(page)).toContain('Text with inserted words and')

    await rightClickText(page, 'deleted words')
    expect((await menuRows(page)).slice(0, 3)).toEqual([
      'Accept Deletion',
      'Reject Deletion',
      '---',
    ])
    await page.screenshot({ path: test.info().outputPath('tracked-delete.png') })
    await menuItem(page, 'Reject Deletion').click()
    await expect.poll(() => marked(page, 'del')).toEqual([])
    expect(await bodyText(page)).toContain('and deleted words end.')

    await saveFromMenu(app)
    await expect
      .poll(async () => (await part(docPath, 'word/document.xml')).includes('<w:ins '), {
        timeout: 20_000,
      })
      .toBe(false)
    const xml = await part(docPath, 'word/document.xml')
    expect(xml).not.toContain('<w:del ')
    expect(xml).toContain('inserted words')
    expect(xml).toContain('deleted words')

    // a bare caret inside a word: the Tools item translates that word, as the right-click one does
    const word = await settledCentreOf(page, 'Last')
    await page.mouse.click(word.x, word.y)
    await expect
      .poll(() =>
        page.evaluate(() => {
          type Sel = { empty: boolean; $from: { parent: { textContent: string } } }
          type Hook = { editor: { state: { selection: Sel } } }
          const { selection } = (window as unknown as { __aidocs: Hook }).__aidocs.editor.state
          return selection.empty && selection.$from.parent.textContent
        }),
      )
      .toBe('Last line')
    await app.evaluate(({ Menu }) => {
      const tools = Menu.getApplicationMenu()?.items.find((item) => item.label === 'Tools')
      const translate = tools?.submenu?.items.find((item) => item.label === 'Translate')
      translate?.submenu?.items.find((item) => item.label === 'Translate Selection')?.click()
    })
    await expect(page.locator('.ai-dock')).not.toHaveClass(/collapsed/)
    await expect(page.locator('.ai-dock')).toContainText('"Last"')
    expect(await page.locator('.status-msg').textContent()).not.toContain('Select the text')
  } finally {
    await closeScene(scene, 'docs-ctx-tracked')
  }
})
