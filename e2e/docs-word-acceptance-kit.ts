/** What the Word acceptance specs share: opening a document, Word's entry points, reading the saved file. */
import { expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { docxWith } from './docs-objects-fixture'

export const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'

export async function clickMenu(app: ElectronApplication, path: string[]) {
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

export const ribbonTab = async (page: Page, name: string) => {
  const tab = page.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).first()
  const control = page.locator('[data-ribbon-body] button').first()
  // a click on the tab that is already open folds the ribbon away, and a contextual tab
  // that goes away hands the ribbon back to the tab that was open before it
  await expect(async () => {
    const open =
      (await tab.getAttribute('class'))?.includes('active') && (await control.isVisible())
    if (!open) await tab.click()
    await expect(tab).toHaveClass(/active/, { timeout: 1_000 })
    await expect(control).toBeVisible({ timeout: 1_000 })
  }).toPass({ timeout: 15_000 })
}

/** a ribbon control by the name Word gives it: its label, its tooltip or its caption */
export const ribbon = (page: Page, name: string): Locator => {
  const buttons = page.locator('[data-ribbon-body] button')
  const exact = new RegExp(`^\\s*${name}\\s*$`)
  return (
    page
      .locator(
        `[data-ribbon-body] button[aria-label="${name}"], [data-ribbon-body] button[data-tip^="${name}"]`,
      )
      .or(buttons.filter({ hasText: exact }))
      // an icon drawn with letters sits in front of the caption
      .or(buttons.filter({ has: page.locator('span', { hasText: exact }) }))
      .first()
  )
}

/** selects the paragraph holding `text`, through the one holding `toText` when given */
export const selectParagraph = (page: Page, text: string, toText = text) =>
  page.evaluate(
    ([text, toText]) => {
      const { editor } = (window as never as { __aidocs: { editor: any } }).__aidocs
      let from = -1
      let to = -1
      editor.state.doc.descendants((node: any, pos: number) => {
        if (!node.isTextblock) return true
        if (from < 0 && node.textContent.includes(text)) from = pos + 1
        if (to < 0 && node.textContent.includes(toText)) to = pos + node.nodeSize - 1
        return false
      })
      // the chain's focus() waits a frame when focus is elsewhere: keys typed next would miss the editor
      editor.view.focus()
      editor.chain().focus().setTextSelection({ from, to }).run()
    },
    [text, toText],
  )

/** selects the first occurrence of `word` in the body */
export const selectWord = (page: Page, word: string) =>
  page.evaluate((word) => {
    const { editor } = (window as never as { __aidocs: { editor: any } }).__aidocs
    let from = -1
    editor.state.doc.descendants((node: any, pos: number) => {
      if (from >= 0 || !node.isText) return true
      const at = (node.text as string).indexOf(word)
      if (at >= 0) from = pos + at
      return true
    })
    editor.view.focus()
    editor
      .chain()
      .focus()
      .setTextSelection({ from, to: from + word.length })
      .run()
  }, word)

/** the caret at the end of the paragraph holding `text` */
export const caretAtEnd = (page: Page, text: string) =>
  page.evaluate((text) => {
    const { editor } = (window as never as { __aidocs: { editor: any } }).__aidocs
    let at = -1
    editor.state.doc.descendants((node: any, pos: number) => {
      if (!node.isTextblock) return true
      if (at < 0 && node.textContent.includes(text)) at = pos + node.nodeSize - 1
      return false
    })
    editor.view.focus()
    editor.chain().focus().setTextSelection(at).run()
  }, text)

/** where the text of the paragraph holding `text` is drawn */
export const textBox = (page: Page, text: string) =>
  page.evaluate((text) => {
    const paragraph = Array.from(document.querySelectorAll('.doc-page :is(p, .doc-li)')).find(
      (el) => el.textContent?.includes(text),
    )!
    const range = document.createRange()
    range.selectNodeContents(paragraph)
    const { x, y, width, height } = range.getBoundingClientRect()
    return { x, y, width, height }
  }, text)

/** Home > Styles: picks `name` from the gallery */
export const pickStyle = async (page: Page, name: string) => {
  await ribbon(page, 'Styles').click()
  await page
    .locator('.style-gallery-menu .style-card')
    .filter({ has: page.locator('.style-card-label', { hasText: new RegExp(`^${name}$`) }) })
    .click()
}

export const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

export const savedBody = (docPath: string) => part(docPath, 'word/document.xml')

/** the text of each paragraph of the saved body */
export const savedText = async (docPath: string) =>
  (await savedBody(docPath))
    .replace(/<w:p[ >]/g, '\n<w:p>')
    .replace(/<[^>]+>/g, '')
    .trim()
    .split('\n')

/** the w:p of the saved body that holds `text` */
export const savedParagraph = async (docPath: string, text: string) =>
  ((await savedBody(docPath)).match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).find((p) =>
    p.includes(`>${text}<`),
  ) ?? ''

const heading = (level: number, size: number) =>
  `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/>` +
  `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
  `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
  `<w:rPr><w:b/><w:sz w:val="${size}"/></w:rPr></w:style>`

const STYLES =
  '<?xml version="1.0" encoding="UTF-8"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' +
  '<w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  heading(1, 40) +
  heading(2, 32) +
  heading(3, 28) +
  '</w:styles>'

/** `body` as a document that carries the styles a Word document always has */
async function styledDocx(body: string): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(body + SECT_PR))
  zip.file('word/styles.xml', STYLES)
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file(
    '[Content_Types].xml',
    types.replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    ),
  )
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  zip.file(
    'word/_rels/document.xml.rels',
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

export interface Session {
  dir: string
  docPath: string
  launched: LaunchedApp
  app: ElectronApplication
  page: Page
  close: () => Promise<void>
}

/** the app with a document of `body` open in Docs; `extra` are other files written beside it */
export async function open(
  videoDir: string,
  body: string,
  name: string,
  extra: Record<string, string> = {},
): Promise<Session> {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-acceptance-')))
  const docPath = join(dir, name)
  writeFileSync(docPath, await styledDocx(body))
  for (const [file, otherBody] of Object.entries(extra))
    writeFileSync(join(dir, file), await styledDocx(otherBody))
  const launched = await launchShell({
    onboardingSeen: true,
    // a new document's first save lands in this folder without asking
    settings: { defaultSaveDir: dir },
    videoDir,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  return {
    dir,
    docPath,
    launched,
    app: launched.app,
    page,
    close: async () => {
      await closeAndSaveVideo(launched, videoDir)
      rmSync(dir, { recursive: true, force: true })
    },
  }
}

/** the next native open dialog answers with `file` */
export const pickFile = (app: ElectronApplication, file: string) =>
  app.evaluate(({ dialog }, file) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [file] })
  }, file)

/** the next native save dialog answers with `file` */
export const pickSavePath = (app: ElectronApplication, file: string) =>
  app.evaluate(({ dialog }, file) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: file })
  }, file)
