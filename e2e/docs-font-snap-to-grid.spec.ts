/**
 * Format > Font > Advanced > Snap to grid when document grid is defined: text
 * taken off the character grid keeps its natural advances while the text
 * around it keeps the grid; the choice is in the saved file and reads back.
 * The expected advances are Word's (probe 2026-11-21, 10.5pt, pitch 12.5pt).
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-font-snap-to-grid'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const OFF = '<w:snapToGrid w:val="0"/>'
const LABEL = 'Snap to grid when document grid is defined'
const KANJI = '\u6c38'
const LATIN = 'abcdefghijkl'

const run = (text: string, rPr = '') => `<w:r><w:rPr>${rPr}</w:rPr><w:t>${text}</w:t></w:r>`
const para = (...runs: string[]) => `<w:p><w:pPr><w:jc w:val="left"/></w:pPr>${runs.join('')}</w:p>`
const four = KANJI.repeat(4)

// 415.3pt of text, 10.5pt characters in cells of 12.5pt
const BODY =
  para(run(KANJI.repeat(12))) +
  para(run(LATIN)) +
  para(run(KANJI.repeat(80))) +
  para(run(four), run(four, OFF), run(four)) +
  para(run(four), run(four, '<w:rStyle w:val="OffGrid"/>'), run(four)) +
  para(run(four), run(four, `${OFF}<w:spacing w:val="40"/>`), run(four)) +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>' +
  '<w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="8192"/></w:sectPr>'

async function docx(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="SimSun"/><w:sz w:val="21"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="character" w:styleId="OffGrid"><w:name w:val="Off Grid"/><w:rPr>${OFF}</w:rPr></w:style></w:styles>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${BODY}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

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

const select = (page: Page, from: number, to: number) =>
  page.evaluate(
    ([from, to]) => {
      const { editor } = (
        window as never as {
          __aidocs: {
            editor: { chain(): { focus(): { setTextSelection(r: object): { run(): boolean } } } }
          }
        }
      ).__aidocs
      editor.chain().focus().setTextSelection({ from, to }).run()
    },
    [from, to],
  )

const paragraphs = async (docPath: string) => {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
  return xml.split('<w:p>').slice(1)
}

/**
 * The nth paragraph on an unzoomed page: the advance of every character of its
 * first line (pt) and how many characters each of its lines holds.
 */
const layoutOf = (page: Page, index: number) =>
  page.evaluate((index) => {
    const paper = document.querySelector('.doc-page')!.getBoundingClientRect()
    // an A4 page is 11906 twips wide
    const k = (11906 / 15 / paper.width) * 0.75
    const p = document.querySelectorAll('.doc-page p')[index]!
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT)
    const range = document.createRange()
    const boxes: DOMRect[] = []
    for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null)
      for (let i = 0; i < node.length; i++) {
        range.setStart(node, i)
        range.setEnd(node, i + 1)
        boxes.push(range.getBoundingClientRect())
      }
    const lines: number[] = []
    let top = Number.NaN
    for (const box of boxes) {
      if (Math.abs(box.top - top) < 2) lines[lines.length - 1]++
      else {
        lines.push(1)
        top = box.top
      }
    }
    const advances: number[] = []
    for (let i = 0; i + 1 < lines[0]; i++)
      advances.push(Math.round((boxes[i + 1].left - boxes[i].left) * k * 100) / 100)
    return { advances, lines }
  }, index)

const close = (to: number[], want: number[]) => {
  expect(to).toHaveLength(want.length)
  to.forEach((value, i) => expect(value, `character ${i}`).toBeCloseTo(want[i], 1))
}
const times = (n: number, value: number) => Array.from({ length: n }, () => value)
const MIXED = [...times(4, 12.5), ...times(4, 10.5), ...times(3, 12.5)]

async function openDocs(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page').first()).toContainText(LATIN, { timeout: 30_000 })
  return { launched, page }
}

async function snapBox(app: ElectronApplication, page: Page): Promise<[Locator, Locator]> {
  await clickMenu(app, ['Format', 'Font…'])
  const dialog = page.locator('.modal.font-dialog')
  await dialog.getByRole('tab', { name: 'Advanced' }).click()
  return [dialog, dialog.getByLabel(LABEL)]
}

async function accept(dialog: Locator) {
  await dialog.locator('button.btn-primary').click()
  await expect(dialog).toBeHidden()
}

test('text taken off the character grid keeps its natural advances, saves and reads back', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-font-snap-')))
  const docPath = join(dir, 'snap.docx')
  writeFileSync(docPath, await docx())
  const first = await openDocs(docPath)
  try {
    const { launched, page } = first
    const { app } = launched
    // on the grid every character advances by the 2pt the pitch adds
    await expect.poll(async () => (await layoutOf(page, 0)).advances[0]).toBeCloseTo(12.5, 1)
    close((await layoutOf(page, 0)).advances, times(11, 12.5))
    expect((await layoutOf(page, 2)).lines).toEqual([33, 33, 14])
    const latinOn = (await layoutOf(page, 1)).advances
    // what the file carries: a run off the grid, one off through its character
    // style, one off with its own 2pt of spacing
    close((await layoutOf(page, 3)).advances, MIXED)
    close((await layoutOf(page, 4)).advances, MIXED)
    close((await layoutOf(page, 5)).advances, times(11, 12.5))
    await page.screenshot({ path: test.info().outputPath('before.png') })

    // four characters in the middle of a line
    await select(page, 5, 9)
    const [dialog, box] = await snapBox(app, page)
    await expect(box).toBeChecked()
    await box.uncheck()
    await dialog.screenshot({ path: test.info().outputPath('advanced-tab.png') })
    await accept(dialog)
    await expect.poll(async () => (await layoutOf(page, 0)).advances[4]).toBeCloseTo(10.5, 1)
    close((await layoutOf(page, 0)).advances, MIXED)

    // Latin letters lose the 2pt, the ones around them keep it
    await select(page, 19, 23)
    await (await snapBox(app, page))[1].uncheck()
    await accept(dialog)
    await expect
      .poll(async () => (await layoutOf(page, 1)).advances[4])
      .toBeCloseTo(latinOn[4] - 2, 1)
    close(
      (await layoutOf(page, 1)).advances,
      latinOn.map((advance, i) => (i >= 4 && i < 8 ? advance - 2 : advance)),
    )

    // a whole paragraph off the grid holds what a line without a grid holds
    await select(page, 29, 109)
    await (await snapBox(app, page))[1].uncheck()
    await accept(dialog)
    await expect.poll(async () => (await layoutOf(page, 2)).lines).toEqual([39, 39, 2])
    close((await layoutOf(page, 2)).advances, times(38, 10.5))
    await page.screenshot({ path: test.info().outputPath('after.png') })

    await saveFromMenu(app)
    await expect
      .poll(async () => (await paragraphs(docPath))[0], { timeout: 30_000 })
      .toContain(`<w:rPr>${OFF}</w:rPr>`)
    const saved = await paragraphs(docPath)
    expect(saved[0].match(/<w:r>/g)).toHaveLength(3)
    expect(saved[0].match(/<w:snapToGrid/g)).toHaveLength(1)
    expect(saved[1]).toContain(
      `<w:r><w:rPr>${OFF}</w:rPr><w:t xml:space="preserve">efgh</w:t></w:r>`,
    )
    expect(saved[2].match(/<w:r>/g)).toHaveLength(1)
    expect(saved[2]).toContain(`<w:rPr>${OFF}</w:rPr>`)
    // the runs the dialog never touched keep their bytes
    expect(saved[3]).toContain(run(four, OFF))
    expect(saved[4]).toContain(run(four, '<w:rStyle w:val="OffGrid"/>'))
    expect(saved[5]).toContain(run(four, `${OFF}<w:spacing w:val="40"/>`))
    writeFileSync(test.info().outputPath('saved.docx'), readFileSync(docPath))
    writeFileSync(
      test.info().outputPath('layout.json'),
      JSON.stringify(await Promise.all([0, 1, 2, 3, 4, 5].map((i) => layoutOf(page, i)))),
    )
  } finally {
    await closeAndSaveVideo(first.launched, VIDEO_DIR)
  }

  const second = await openDocs(docPath)
  try {
    const { launched, page } = second
    const { app } = launched
    await expect.poll(async () => (await layoutOf(page, 0)).advances[4]).toBeCloseTo(10.5, 1)
    close((await layoutOf(page, 0)).advances, MIXED)
    // Times New Roman's own advance of an e at 10.5pt
    expect((await layoutOf(page, 1)).advances[4]).toBeCloseTo(4.66, 1)
    expect((await layoutOf(page, 2)).lines).toEqual([39, 39, 2])

    // the box shows the text under the selection; ticked, the text is back on the grid
    await select(page, 2, 3)
    const [dialog, box] = await snapBox(app, page)
    await expect(box).toBeChecked()
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await select(page, 5, 9)
    await expect((await snapBox(app, page))[1]).not.toBeChecked()
    await box.check()
    await accept(dialog)
    await expect.poll(async () => (await layoutOf(page, 0)).advances[4]).toBeCloseTo(12.5, 1)
    // text off the grid through its character style takes an explicit on
    await select(page, 129, 133)
    await expect((await snapBox(app, page))[1]).not.toBeChecked()
    await box.check()
    await accept(dialog)
    await expect.poll(async () => (await layoutOf(page, 4)).advances[4]).toBeCloseTo(12.5, 1)
    await saveFromMenu(app)
    await expect
      .poll(async () => (await paragraphs(docPath))[0], { timeout: 30_000 })
      .toContain('<w:rPr><w:snapToGrid/></w:rPr>')
    const saved = await paragraphs(docPath)
    expect(saved[0]).not.toContain(OFF)
    expect(saved[4]).toContain('<w:rPr><w:rStyle w:val="OffGrid"/><w:snapToGrid/></w:rPr>')
  } finally {
    await closeAndSaveVideo(second.launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
