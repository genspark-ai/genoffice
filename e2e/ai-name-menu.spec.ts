import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { mkdtemp, readFile, writeFile, cp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * AI file naming ("Name with AI…") lives in the ribbon's File tab, which docs
 * and slides only render on Windows/Linux — on macOS the whole tab is hidden,
 * and neither app's native File menu carried the action, so it had no entry at
 * all there. The action is in the native File menu now.
 *
 * A unit test cannot see this: the defect is a menu template (Electron) plus an
 * IPC hop to the renderer, so this drives the real application menu and the real
 * command channel, the same way menu-user-guide.spec.ts reads the menu back out
 * of the main process.
 */

async function minimalDocx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Name me</w:t></w:r></w:p></w:body></w:document>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function makeDeck(dir: string): Promise<string> {
  const fixture = join(dir, 'fixture')
  await cp(resolve('e2e/assets/font-manager-rubik'), fixture, { recursive: true })
  const slideXml = join(fixture, 'ppt/slides/slide1.xml')
  await writeFile(slideXml, (await readFile(slideXml, 'utf8')).replaceAll('Rubik', 'Arial'))
  const pptx = join(dir, 'deck.pptx')
  execFileSync('zip', ['-X', '-q', '-r', pptx, '.'], { cwd: fixture })
  return pptx
}

/** the File menu's item labels, read from the live application menu */
const fileMenuLabels = (app: ElectronApplication): Promise<string[]> =>
  app.evaluate(({ Menu }) => {
    const file = Menu.getApplicationMenu()?.items.find((item) => item.label === 'File')
    return (file?.submenu?.items ?? []).map((item) => item.label || item.type || '')
  })

const clickFileMenuItem = (app: ElectronApplication, label: string): Promise<boolean> =>
  app.evaluate(({ Menu }, label) => {
    const file = Menu.getApplicationMenu()?.items.find((item) => item.label === 'File')
    const item = file?.submenu?.items.find((candidate) => candidate.label === label)
    if (!item || typeof item.click !== 'function') return false
    item.click()
    return true
  }, label)

/** listen for the menu command in the editor renderer, click the menu, expect the command */
async function expectMenuCommandReachesEditor(app: ElectronApplication, page: Page): Promise<void> {
  await page.evaluate(() => {
    const w = window as unknown as {
      __menuCmds?: string[]
      desktop?: { onMenuCommand: (handler: (cmd: string) => void) => () => void }
      slidesApi?: { onMenuCommand: (handler: (cmd: string) => void) => () => void }
    }
    w.__menuCmds = []
    const subscribe = w.desktop?.onMenuCommand ?? w.slidesApi?.onMenuCommand
    subscribe?.((cmd) => w.__menuCmds?.push(cmd))
  })
  expect(await clickFileMenuItem(app, 'Name with AI…')).toBe(true)
  await expect
    .poll(async () =>
      page.evaluate(() => (window as unknown as { __menuCmds?: string[] }).__menuCmds),
    )
    .toContain('name-with-ai')
}

test('docs: the File menu names the document with AI', async () => {
  test.setTimeout(180_000)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-ai-name-docs-'))
  const docPath = join(dir, 'name-me.docx')
  await writeFile(docPath, await minimalDocx())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'ai-name-menu-docs',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await page.waitForSelector('.ribbon-tab', { state: 'visible' })
    expect(await fileMenuLabels(launched.app)).toContain('Name with AI…')
    await expectMenuCommandReachesEditor(launched.app, page)
  } finally {
    await closeAndSaveVideo(launched, 'ai-name-menu-docs')
  }
})

test('slides: the File menu names the deck with AI', async () => {
  test.setTimeout(180_000)
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-ai-name-slides-'))
  const deckPath = await makeDeck(dir)
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'ai-name-menu-slides',
    openFile: deckPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://slides/')
    await page.waitForSelector('.ribbon-tab', { state: 'visible' })
    expect(await fileMenuLabels(launched.app)).toContain('Name with AI…')
    await expectMenuCommandReachesEditor(launched.app, page)
  } finally {
    await closeAndSaveVideo(launched, 'ai-name-menu-slides')
  }
})
