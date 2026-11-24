/**
 * Word's Expand / Collapse of the body under a heading, on a real document:
 * the triangle beside the heading, the right-click submenu, the keys, the
 * Paragraph dialog's "Collapsed by default" through a save and a reopen, the
 * exported PDF, Find stepping into a hidden body, and editing a collapsed
 * heading.
 */
import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { docxWith, saveFromMenu, withWordPart } from './docs-objects-fixture'
import { clickMenu } from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-collapsible-headings'
const IS_MAC = process.platform === 'darwin'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"'

const headingStyle = (n: number) =>
  `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/>` +
  `<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:outlineLvl w:val="${n - 1}"/></w:pPr>` +
  `<w:rPr><w:b/><w:sz w:val="${36 - 4 * n}"/></w:rPr></w:style>`
const STYLES =
  `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}>` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  `${headingStyle(1)}${headingStyle(2)}</w:styles>`

const p = (text: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t>${text}</w:t></w:r></w:p>`
const h = (level: number, text: string, extra = '') =>
  p(text, `<w:pStyle w:val="Heading${level}"/>${extra}`)
const filler = (count: number) =>
  Array.from({ length: count }, (_, i) => p(`alpha filler line ${i + 1}`)).join('')

/** Alpha's body runs over the first page; the last paragraph is the open marker */
const body = (alphaExtra = '') =>
  p('Intro paragraph') +
  h(1, 'Alpha', alphaExtra) +
  p('alpha body with the needle') +
  h(2, 'Alpha sub') +
  p('alpha sub body') +
  filler(60) +
  h(1, 'Beta') +
  p('beta body') +
  h(1, 'Gamma') +
  p('Last paragraph')

interface Snap {
  blocks: string[]
  /** texts of the blocks a collapsed heading hides */
  hidden: string[]
  collapsed: string[]
  pages: number
  headBlock: string
}

const snap = (page: Page): Promise<Snap> =>
  page.evaluate(() => {
    interface Node {
      textContent: string
      nodeSize: number
    }
    const win = window as never as {
      __aidocs: {
        editor: {
          view: { nodeDOM(pos: number): Element | null }
          state: {
            doc: { forEach(visit: (node: Node, pos: number) => void): void }
            selection: { head: number }
          }
        }
      }
      __pageDebug?: { pageCuts?: unknown[] }
    }
    const { editor } = win.__aidocs
    const blocks: string[] = []
    const hidden: string[] = []
    const collapsed: string[] = []
    let headBlock = ''
    editor.state.doc.forEach((node, pos) => {
      blocks.push(node.textContent)
      const el = editor.view.nodeDOM(pos)
      if (el?.classList.contains('doc-collapsed-body')) hidden.push(node.textContent)
      if (el?.classList.contains('doc-heading-collapsed')) collapsed.push(node.textContent)
      const { head } = editor.state.selection
      if (head >= pos && head < pos + node.nodeSize) headBlock = node.textContent
    })
    return { blocks, hidden, collapsed, pages: win.__pageDebug?.pageCuts?.length ?? 0, headBlock }
  })

/** a caret (or `length` characters selected) at `offset` of the block whose text is `text` */
const selectIn = (page: Page, text: string, offset: number, length = 0) =>
  page.evaluate(
    ({ text, offset, length }) => {
      interface Node {
        textContent: string
      }
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: {
              state: { doc: { forEach(visit: (node: Node, pos: number) => void): void } }
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
      editor.state.doc.forEach((node, pos) => {
        if (at < 0 && node.textContent === text) at = pos + 1
      })
      if (at < 0) throw new Error(`no block reads "${text}"`)
      editor
        .chain()
        .focus()
        .setTextSelection({ from: at + offset, to: at + offset + length })
        .run()
    },
    { text, offset, length },
  )

/** selectIn, then wait for the focus it asks for: keys pressed before it lands go elsewhere */
async function caretIn(page: Page, text: string, offset: number, length = 0) {
  await selectIn(page, text, offset, length)
  await expect
    .poll(() => page.evaluate(() => document.activeElement?.classList.contains('ProseMirror')))
    .toBe(true)
}

async function newDoc(alphaExtra = '') {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-collapse-')))
  const docPath = join(dir, 'headings.docx')
  const docx = await withWordPart(await docxWith(body(alphaExtra)), 'styles', STYLES)
  writeFileSync(docPath, docx)
  return { dir, docPath }
}

async function open(docPath: string) {
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
  await expect(page.locator('.doc-page p', { hasText: 'Last paragraph' })).toBeVisible({
    timeout: 30_000,
  })
  return { launched, page }
}

async function close(launched: LaunchedApp, dir: string) {
  await closeAndSaveVideo(launched, VIDEO_DIR)
  rmSync(dir, { recursive: true, force: true })
}

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')!.async('string')

async function pdfText(path: string): Promise<string> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(path)) }).promise
  const out: string[] = []
  for (let i = 1; i <= doc.numPages; i++) {
    const content = await (await doc.getPage(i)).getTextContent()
    out.push(content.items.map((it) => ('str' in it ? it.str : '')).join(' '))
  }
  return out.join('\n')
}

const heading = (page: Page, text: string) =>
  page.locator('.doc-page :is(h1, h2)', { hasText: new RegExp(`^${text}$`) })
const para = (page: Page, text: string) => page.locator('.doc-page p', { hasText: text }).first()
const ALPHA_BODY = ['alpha body with the needle', 'Alpha sub', 'alpha sub body']

async function foldFromMenu(page: Page, on: string, item: string) {
  await heading(page, on).click({ button: 'right' })
  const menu = page.locator('.ctx-menu')
  await menu.waitFor()
  await menu.locator('.ctx-item', { hasText: /^Expand\/Collapse/ }).hover()
  await menu.locator('.ctx-submenu .ctx-item', { hasText: new RegExp(`^${item}$`) }).click()
  await expect(menu).toBeHidden()
}

test('the triangle, the right-click submenu and the keys fold a heading; the pages follow', async () => {
  test.setTimeout(180_000)
  const { dir, docPath } = await newDoc()
  const { launched, page } = await open(docPath)
  try {
    const all = await snap(page)
    expect(all.hidden).toEqual([])
    expect(all.pages).toBeGreaterThan(1)
    // nothing is drawn in a document with no fold
    await expect(page.locator('.doc-page .heading-toggle, .doc-page .doc-fold-gap')).toHaveCount(0)

    // the triangle shows beside the heading under the pointer and collapses its body
    const hover = page.locator('.heading-toggle-hover')
    await expect(hover).toBeHidden()
    await heading(page, 'Alpha').hover()
    await expect(hover).toBeVisible()
    await hover.click()
    await expect.poll(async () => (await snap(page)).collapsed).toEqual(['Alpha'])
    let now = await snap(page)
    expect(now.hidden.slice(0, 3)).toEqual(ALPHA_BODY)
    expect(now.hidden).toHaveLength(ALPHA_BODY.length + 60)
    expect(now.blocks).toEqual(all.blocks)
    await expect(para(page, 'alpha body with the needle')).toBeHidden()
    await expect(heading(page, 'Beta')).toBeVisible()
    // the hidden body takes no space: the document is one page now
    await expect.poll(async () => (await snap(page)).pages).toBe(1)

    // a body paragraph has no triangle; the collapsed heading keeps its own, which expands it
    await para(page, 'beta body').hover()
    await expect(hover).toBeHidden()
    await page.locator('.doc-page .heading-toggle.is-collapsed').click()
    await expect.poll(async () => (await snap(page)).hidden).toEqual([])
    await expect.poll(async () => (await snap(page)).pages).toBe(all.pages)

    // right-click > Expand/Collapse
    await heading(page, 'Beta').click({ button: 'right' })
    const menu = page.locator('.ctx-menu')
    await menu.locator('.ctx-item', { hasText: /^Expand\/Collapse/ }).hover()
    const sub = menu.locator('.ctx-submenu .ctx-item')
    await expect(sub).toHaveText([
      'Expand Heading',
      'Collapse Heading',
      'Expand All Headings',
      'Collapse All Headings',
    ])
    await expect(sub.nth(0)).toBeDisabled()
    await sub.nth(1).click()
    await expect.poll(async () => (await snap(page)).hidden).toEqual(['beta body'])
    // a body paragraph has no such submenu
    await para(page, 'Intro paragraph').click({ button: 'right' })
    await menu.waitFor()
    await expect(menu.locator('.ctx-item', { hasText: /^Expand\/Collapse/ })).toHaveCount(0)
    await page.keyboard.press('Escape')

    await foldFromMenu(page, 'Beta', 'Collapse All Headings')
    // every heading with a body folds; the ones inside Alpha are hidden by it
    await expect.poll(async () => (await snap(page)).collapsed).toEqual(['Alpha', 'Beta', 'Gamma'])
    now = await snap(page)
    expect(now.blocks.filter((b) => !now.hidden.includes(b))).toEqual([
      'Intro paragraph',
      'Alpha',
      'Beta',
      'Gamma',
    ])
    await foldFromMenu(page, 'Beta', 'Expand All Headings')
    await expect.poll(async () => (await snap(page)).hidden).toEqual([])

    // Word for Mac's Collapse / Expand keys act on the heading the caret is under
    if (IS_MAC) {
      await caretIn(page, 'alpha sub body', 3)
      await page.keyboard.press('Control+Shift+Minus')
      await expect.poll(async () => (await snap(page)).collapsed).toEqual(['Alpha sub'])
      expect((await snap(page)).headBlock).toBe('Alpha sub')
      await page.keyboard.press('Control+Shift+Equal')
      await expect.poll(async () => (await snap(page)).collapsed).toEqual([])
    }
  } finally {
    await close(launched, dir)
  }
})

test('"Collapsed by default" is saved as w15:collapsed, reopens collapsed and prints without the body', async () => {
  test.setTimeout(240_000)
  const { dir, docPath } = await newDoc()
  let { launched, page } = await open(docPath)
  try {
    const ui = (page: Page, app: LaunchedApp['app']) => {
      const dialog = page.locator('.modal', { has: page.locator('h2', { hasText: /^Paragraph$/ }) })
      const box = dialog.locator('input[type="checkbox"][name="collapsed"]')
      const openDialog = async (text: string) => {
        await selectIn(page, text, 1)
        await clickMenu(app, ['Format', 'Paragraph…'])
        await expect(dialog).toBeVisible()
      }
      return { dialog, box, openDialog }
    }
    let { dialog, box, openDialog } = ui(page, launched.app)
    // offered to a paragraph with an outline level only
    await openDialog('beta body')
    await expect(box).toBeDisabled()
    await dialog.locator('button', { hasText: /^Cancel$/ }).click()
    await expect(dialog).toBeHidden()

    await openDialog('Alpha')
    await expect(box).toBeEnabled()
    await expect(box).not.toBeChecked()
    await box.check()
    await dialog.locator('button.btn-primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await snap(page)).collapsed).toEqual(['Alpha'])

    // folding Beta by its triangle is view state: it is not written
    await heading(page, 'Beta').hover()
    await page.locator('.heading-toggle-hover').click()
    await expect.poll(async () => (await snap(page)).collapsed).toEqual(['Alpha', 'Beta'])

    await saveFromMenu(launched.app)
    await expect.poll(() => documentXml(docPath), { timeout: 30_000 }).toContain('w15:collapsed')
    const xml = await documentXml(docPath)
    expect(xml.match(/<w15:collapsed\b/g)).toHaveLength(1)
    // the dialog also writes the spacing it shows; the flag follows it, in schema order
    expect(xml).toMatch(
      /<w:pPr><w:pStyle w:val="Heading1"\/>(<w:spacing [^>]*\/>)?<w15:collapsed\/><\/w:pPr><w:r><w:t[^>]*>Alpha</,
    )
    expect(xml).toContain(W15)
    // the save keeps what the view has folded
    await expect.poll(async () => (await snap(page)).collapsed).toEqual(['Alpha', 'Beta'])

    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, page } = await open(docPath))
    ;({ dialog, box, openDialog } = ui(page, launched.app))
    const reopened = await snap(page)
    expect(reopened.collapsed).toEqual(['Alpha'])
    expect(reopened.hidden.slice(0, 3)).toEqual(ALPHA_BODY)
    expect(reopened.pages).toBe(1)
    await openDialog('Alpha')
    await expect(box).toBeChecked()
    await dialog.locator('button', { hasText: /^Cancel$/ }).click()

    // Word prints a collapsed body hidden; so does the exported PDF
    const pdfPath = join(dir, 'headings.pdf')
    await launched.app.evaluate(({ dialog, shell }, out) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    const text = await pdfText(pdfPath)
    expect(text).toContain('Alpha')
    expect(text).toContain('beta body')
    expect(text).not.toContain('needle')
    expect(text).not.toContain('alpha filler')
    expect(text).not.toContain('Alpha sub')
  } finally {
    await close(launched, dir)
  }
})

test('Find opens a collapsed heading; Enter goes below the hidden body; deleting the heading takes it', async () => {
  test.setTimeout(180_000)
  const { dir, docPath } = await newDoc('<w15:collapsed/>')
  const { launched, page } = await open(docPath)
  try {
    const opened = await snap(page)
    expect(opened.collapsed).toEqual(['Alpha'])
    await expect(para(page, 'alpha body with the needle')).toBeHidden()

    // Find lands in the hidden body: the heading expands, as in Word
    await page.keyboard.press('ControlOrMeta+f')
    const panel = page.locator('.find-panel')
    await panel.waitFor()
    await panel.locator('.find-input').first().fill('needle')
    await expect(panel.locator('.find-count')).toHaveText('1/1')
    await expect.poll(async () => (await snap(page)).collapsed).toEqual([])
    await expect(page.locator('.doc-page .search-hit-active')).toBeVisible()
    await page.keyboard.press('Escape')

    // collapse again; typing at the end of the heading and Enter opens a paragraph below the body
    await heading(page, 'Alpha').hover()
    await page.locator('.heading-toggle-hover').click()
    await expect.poll(async () => (await snap(page)).collapsed).toEqual(['Alpha'])
    await caretIn(page, 'Alpha', 'Alpha'.length)
    await page.keyboard.type(' one')
    await page.keyboard.press('Enter')
    await page.keyboard.type('typed below')
    let now = await snap(page)
    expect(now.collapsed).toEqual(['Alpha one'])
    expect(now.hidden).toHaveLength(ALPHA_BODY.length + 60)
    const at = now.blocks.indexOf('typed below')
    expect(now.blocks[at - 1]).toBe('alpha filler line 60')
    expect(now.blocks[at + 1]).toBe('Beta')
    await expect(para(page, 'typed below')).toBeVisible()

    // the selected collapsed heading, deleted, takes its hidden body with it
    await caretIn(page, 'Alpha one', 0, 'Alpha one'.length)
    await page.keyboard.press('Backspace')
    await expect
      .poll(async () => (await snap(page)).blocks)
      .toEqual([
        'Intro paragraph',
        '',
        'typed below',
        'Beta',
        'beta body',
        'Gamma',
        'Last paragraph',
      ])
    // one undo brings the heading and the body back
    await page.keyboard.press('ControlOrMeta+z')
    now = await snap(page)
    expect(now.blocks).toContain('alpha body with the needle')
    expect(now.blocks).toContain('Alpha one')
  } finally {
    await close(launched, dir)
  }
})
