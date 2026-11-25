/**
 * A shape read from a file that holds no text keeps what Shape Format changes
 * on it: turned and put behind the text, it is saved that way and comes back
 * that way, next to a shape with a word in it that is turned too.
 */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu, shapeParagraph } from './docs-objects-fixture'

const documentXml = async (file: string): Promise<string> => {
  const zip = await JSZip.loadAsync(readFileSync(file))
  return (await zip.file('word/document.xml')?.async('string')) ?? ''
}

const ready = (page: Page) =>
  page.waitForFunction(
    () => Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
    undefined,
    { timeout: 30_000 },
  )

const control = (page: Page, tip: string) =>
  page.locator(`[data-ribbon-body] button[data-tip="${tip}"]`)

async function pick(page: Page, tip: string, item: string) {
  await control(page, tip).click()
  await page.locator('[data-ribbon-body] [data-rb-panel] button', { hasText: item }).click()
  await expect(page.locator('[data-ribbon-body] [data-rb-panel]')).toHaveCount(0)
}

/** rotation, wrap side and text of every shape the editor holds, in document order */
const shapes = (page: Page): Promise<{ rotDeg: number | null; behind: boolean; text: string }[]> =>
  page.evaluate(() => {
    type Box = { rotDeg?: number; behind?: boolean; paras: { runs: { text: string }[] }[] }
    type Node = { type: { name: string }; attrs: { textboxes?: Box[] | null } }
    const editor = (
      window as unknown as {
        __aidocs: { editor: { state: { doc: { descendants(f: (n: Node) => void): void } } } }
      }
    ).__aidocs.editor
    const out: { rotDeg: number | null; behind: boolean; text: string }[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name !== 'docProtected') return
      for (const box of node.attrs.textboxes ?? [])
        out.push({
          rotDeg: box.rotDeg ?? null,
          behind: !!box.behind,
          text: box.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n'),
        })
    })
    return out
  })

const labelled = (xml: string) =>
  xml.replace(
    '<wps:bodyPr/>',
    '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Shape</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/>',
  )

test('a shape without text saves a new rotation and wrap', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(
    mkdtempSync(join(process.env.TMPDIR ?? tmpdir(), 'genoffice-e2e-textless-shape-')),
  )
  const userDataDir = join(dir, 'user-data')
  mkdirSync(userDataDir)
  const docPath = join(dir, 'shapes.docx')
  writeFileSync(
    docPath,
    await docxWith(
      para('One') +
        shapeParagraph(3, 300, 20, 120, 80) +
        para('Two') +
        para('Three') +
        para('Four') +
        para('Five') +
        para('Six') +
        labelled(shapeParagraph(4, 300, 20, 120, 80)) +
        para('Last'),
    ),
  )
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-textless-shape-save',
    openFile: docPath,
    userDataDir,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await ready(page)
    const boxes = page.locator('.doc-page .doc-textbox')
    await expect(boxes).toHaveCount(2, { timeout: 30_000 })
    expect(await shapes(page)).toEqual([
      { rotDeg: null, behind: false, text: '' },
      { rotDeg: null, behind: false, text: 'Shape' },
    ])
    const tab = page.locator('.ribbon-tab.active')

    await boxes.nth(0).click({ position: { x: 3, y: 3 } })
    await expect(tab).toHaveText('Shape Format')
    await pick(page, 'Rotate', 'Rotate Right 90°')
    await pick(page, 'Wrap Text', 'Behind Text')
    await expect
      .poll(async () => (await shapes(page))[0])
      .toEqual({ rotDeg: 90, behind: true, text: '' })
    await page.screenshot({ path: test.info().outputPath('turned.png') })

    await boxes.nth(1).click({ position: { x: 3, y: 3 } })
    await expect(tab).toHaveText('Shape Format')
    await pick(page, 'Rotate', 'Rotate Left 90°')
    await expect.poll(async () => (await shapes(page))[1].rotDeg).toBe(270)

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          const drawing = (id: number) =>
            new RegExp(
              `<w:drawing>(?:(?!</w:drawing>)[\\s\\S])*<wp:docPr id="${id}"[\\s\\S]*?</w:drawing>`,
            ).exec(xml)?.[0] ?? ''
          const [bare, withText] = [3, 4].map(drawing)
          return [
            /<a:xfrm rot="5400000">/.test(bare),
            /behindDoc="1"/.test(bare) && /<wp:wrapNone\/>/.test(bare),
            // still at the place it was read from, and still without a text body
            /<wp:posOffset>2857500<\/wp:posOffset>/.test(bare) && !/<w:txbxContent/.test(bare),
            /<a:xfrm rot="16200000">/.test(withText) && /behindDoc="0"/.test(withText),
          ]
        },
        { timeout: 30_000 },
      )
      .toEqual([true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-textless-shape-save')
  }

  launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-textless-shape-save',
    openFile: docPath,
    userDataDir,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await ready(page)
    await expect(page.locator('.doc-page .doc-textbox')).toHaveCount(2, { timeout: 30_000 })
    expect(await shapes(page)).toEqual([
      { rotDeg: 90, behind: true, text: '' },
      { rotDeg: 270, behind: false, text: 'Shape' },
    ])
    await page.screenshot({ path: test.info().outputPath('reopened.png') })
  } finally {
    await closeAndSaveVideo(launched, 'docs-textless-shape-save')
    rmSync(dir, { recursive: true, force: true })
  }
})
