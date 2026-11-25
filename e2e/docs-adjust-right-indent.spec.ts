/**
 * Format > Paragraph… > Automatically adjust right indent when document grid
 * is defined. Below compatibilityMode 15 a paragraph with the box on ends its
 * lines short of the right indent by what the character cells leave over in
 * the column (Word probes 2026-11-22: 415.3pt of text, cells of 11pt, 8.3pt
 * over); a Word 2013 document lays out the same with the box on or off.
 */
import { test, expect, type ElectronApplication, type Locator, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { saveFromMenu } from './docs-objects-fixture'

const VIDEO_DIR = 'docs-adjust-right-indent'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const SECT =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>' +
  '<w:cols w:space="425"/><w:docGrid w:type="linesAndChars" w:linePitch="312" w:charSpace="2048"/></w:sectPr>'
const KANJI = '\u6c38'.repeat(120)
const PX_PER_PT = 96 / 72
/** 415.3pt hold 37 cells of 11pt */
const LEFT_OVER_PX = (415.3 - 37 * 11) * PX_PER_PT

/** a label paragraph, then the paragraph under test: the labels find the paragraphs in a Word export too */
const probe = (label: string, pPr: string, text = KANJI) =>
  `<w:p><w:pPr><w:adjustRightInd w:val="0"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:sz w:val="12"/></w:rPr><w:t>#${label}#</w:t></w:r></w:p>` +
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

const BODY =
  probe('P1', '') +
  probe('P2', '<w:adjustRightInd w:val="0"/>') +
  probe('P3', '<w:ind w:right="88"/>') +
  probe('P4', '<w:pStyle w:val="AdjOff"/>') +
  probe('P5', '<w:jc w:val="left"/>', 'i'.repeat(400)) +
  SECT

async function docx(compatibilityMode?: number): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
  )
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="SimSun"/><w:sz w:val="21"/></w:rPr></w:rPrDefault></w:docDefaults>` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>' +
      '<w:style w:type="paragraph" w:customStyle="1" w:styleId="AdjOff"><w:name w:val="AdjOff"/><w:basedOn w:val="Normal"/><w:pPr><w:adjustRightInd w:val="0"/></w:pPr></w:style></w:styles>',
  )
  zip.file(
    'word/settings.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:settings ${W}><w:compat><w:balanceSingleByteDoubleByteWidth/><w:useFELayout/>${
      compatibilityMode
        ? `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${compatibilityMode}"/>`
        : ''
    }</w:compat></w:settings>`,
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${W}><w:body>${BODY}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
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

const part = async (docPath: string, name: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file(name)?.async('string') ?? ''

/** the w:pPr of the paragraph under a label */
const pPrOf = async (docPath: string, label: string) => {
  const xml = await part(docPath, 'word/document.xml')
  const after = xml.split(`#${label}#`)[1] ?? ''
  return /<w:p>(?:<w:pPr>(.*?)<\/w:pPr>)?<w:r>/.exec(after)?.[1] ?? ''
}

async function openDocs(docPath: string) {
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await expect(page.locator('.doc-page').first()).toContainText('#P5#', { timeout: 30_000 })
  return { launched, page }
}

/**
 * The first line of the paragraph under a label, in px of an unzoomed page:
 * how many characters it holds, and how far before the right edge of the text
 * column the paragraph box and the line's last character end.
 */
const lineUnder = (page: Page, label: string) =>
  page.evaluate((label) => {
    const paper = document.querySelector('.doc-page')!.getBoundingClientRect()
    const k = ((11906 / 1440) * 96) / paper.width
    const blocks = [...document.querySelectorAll('.doc-page > p')]
    const at = blocks.findIndex((b) => b.textContent === `#${label}#`)
    const p = blocks[at + 1]!
    // a label paragraph has no indent: its box ends at the column's right edge
    const column = blocks[at]!.getBoundingClientRect().right
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT)
    const node = walker.nextNode() as Text
    const range = document.createRange()
    const box = (i: number) => {
      range.setStart(node, i)
      range.setEnd(node, i + 1)
      return range.getBoundingClientRect()
    }
    const top = box(0).top
    let count = 1
    while (count < node.length && Math.abs(box(count).top - top) < 2) count++
    const round = (n: number) => Math.round(n * k * 10) / 10
    return {
      count,
      boxGap: round(column - p.getBoundingClientRect().right),
      lineGap: round(column - box(count - 1).right),
    }
  }, label)

async function openParagraph(app: ElectronApplication, page: Page, label: string) {
  const labelPara = page.locator('.doc-page > p', { hasText: `#${label}#` })
  await labelPara.locator('xpath=following-sibling::p[1]').click()
  await clickMenu(app, ['Format', 'Paragraph…'])
  const dialog = page.locator('.modal', { hasText: 'Paragraph' })
  await expect(dialog.locator('input[name="adjustRightInd"]')).toBeVisible()
  return dialog
}

async function accept(dialog: Locator) {
  await dialog.locator('button.btn-primary').click()
  await expect(dialog).toBeHidden()
}

const near = (value: number, want: number, slack = 1) =>
  expect(Math.abs(value - want), `${value} is not near ${want}`).toBeLessThan(slack)

test('below compatibilityMode 15 the box moves the right indent onto the character grid', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-adjust-rind-')))
  const docPath = join(dir, 'adjust-right-indent.docx')
  writeFileSync(docPath, await docx())
  let { launched, page } = await openDocs(docPath)
  try {
    const box = (dialog: Locator) => dialog.locator('input[name="adjustRightInd"]')

    // as opened: on by default, off on the paragraph, off through the style
    const p1 = await lineUnder(page, 'P1')
    expect(p1.count).toBe(37)
    near(p1.boxGap, LEFT_OVER_PX)
    near(p1.lineGap, LEFT_OVER_PX, 1.5)
    const p2 = await lineUnder(page, 'P2')
    expect(p2.count).toBe(37)
    near(p2.boxGap, 0)
    near(p2.lineGap, 0, 1.5)
    // 0.4 of a cell of right indent: one character fewer with the box on
    const p3 = await lineUnder(page, 'P3')
    expect(p3.count).toBe(36)
    near(p3.boxGap, LEFT_OVER_PX + 4.4 * PX_PER_PT)
    near((await lineUnder(page, 'P4')).boxGap, 0)
    // Latin text takes the same share: Word fits 128 of these with the box on, 131 with it off
    const p5 = await lineUnder(page, 'P5')
    near(p5.boxGap, LEFT_OVER_PX)
    expect(p5.count).toBe(128)
    await page.screenshot({ path: test.info().outputPath('opened.png') })

    // the box shows each state, and changing it moves the line end
    let dialog = await openParagraph(launched.app, page, 'P1')
    await expect(box(dialog)).toBeChecked()
    await box(dialog).uncheck()
    await page.screenshot({ path: test.info().outputPath('dialog.png') })
    await accept(dialog)
    await expect.poll(async () => (await lineUnder(page, 'P1')).boxGap).toBeLessThan(1)
    near((await lineUnder(page, 'P1')).lineGap, 0, 1.5)

    dialog = await openParagraph(launched.app, page, 'P2')
    await expect(box(dialog)).not.toBeChecked()
    await box(dialog).check()
    await accept(dialog)
    await expect.poll(async () => (await lineUnder(page, 'P2')).boxGap).toBeGreaterThan(10)
    near((await lineUnder(page, 'P2')).lineGap, LEFT_OVER_PX, 1.5)

    dialog = await openParagraph(launched.app, page, 'P4')
    await expect(box(dialog)).not.toBeChecked()
    await box(dialog).check()
    await accept(dialog)
    await expect.poll(async () => (await lineUnder(page, 'P4')).boxGap).toBeGreaterThan(10)

    dialog = await openParagraph(launched.app, page, 'P5')
    await expect(box(dialog)).toBeChecked()
    await box(dialog).uncheck()
    await accept(dialog)
    await expect.poll(async () => (await lineUnder(page, 'P5')).count).toBe(131)

    // one undo brings the last change back (the keys go to the page once it has the focus again)
    await page.locator('.editor-scroll .doc-page').first().focus()
    await page.keyboard.press('ControlOrMeta+z')
    await expect.poll(async () => (await lineUnder(page, 'P5')).count).toBe(128)
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect.poll(async () => (await lineUnder(page, 'P5')).count).toBe(131)
    await page.screenshot({ path: test.info().outputPath('changed.png') })
    const changed = {
      P1: await lineUnder(page, 'P1'),
      P2: await lineUnder(page, 'P2'),
      P3: await lineUnder(page, 'P3'),
      P4: await lineUnder(page, 'P4'),
      P5: await lineUnder(page, 'P5'),
    }

    await saveFromMenu(launched.app)
    await expect
      .poll(async () => pPrOf(docPath, 'P1'), { timeout: 20_000 })
      .toContain('<w:adjustRightInd w:val="0"/>')
    expect(await pPrOf(docPath, 'P2')).toContain('<w:adjustRightInd/>')
    expect(await pPrOf(docPath, 'P3')).not.toContain('adjustRightInd')
    expect(await pPrOf(docPath, 'P4')).toMatch(/<w:pStyle w:val="AdjOff"\/><w:adjustRightInd\/>/)
    expect(await pPrOf(docPath, 'P5')).toContain('<w:adjustRightInd w:val="0"/>')
    expect(await part(docPath, 'word/styles.xml')).toContain(
      '<w:pPr><w:adjustRightInd w:val="0"/></w:pPr>',
    )

    // reopened: the same line ends and the same boxes
    await closeAndSaveVideo(launched, VIDEO_DIR)
    ;({ launched, page } = await openDocs(docPath))
    for (const label of ['P1', 'P2', 'P3', 'P4', 'P5'] as const) {
      const line = await lineUnder(page, label)
      expect(line.count, label).toBe(changed[label].count)
      near(line.boxGap, changed[label].boxGap)
    }
    dialog = await openParagraph(launched.app, page, 'P1')
    await expect(box(dialog)).not.toBeChecked()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    dialog = await openParagraph(launched.app, page, 'P4')
    await expect(box(dialog)).toBeChecked()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a Word 2013 document keeps its line ends whatever the box says, and saves the box', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-adjust-rind-15-')))
  const docPath = join(dir, 'adjust-right-indent-15.docx')
  writeFileSync(docPath, await docx(15))
  const { launched, page } = await openDocs(docPath)
  try {
    const box = (dialog: Locator) => dialog.locator('input[name="adjustRightInd"]')
    const before = {
      P1: await lineUnder(page, 'P1'),
      P2: await lineUnder(page, 'P2'),
      P3: await lineUnder(page, 'P3'),
      P5: await lineUnder(page, 'P5'),
    }
    // every justified line runs to the indent: Word 2013 does not move it onto the grid
    near(before.P1.boxGap, 0)
    near(before.P1.lineGap, 0, 1.5)
    near(before.P2.boxGap, 0)
    near(before.P3.boxGap, 4.4 * PX_PER_PT)
    expect(before.P3.count).toBe(37)
    expect(before.P5.count).toBe(131)

    let dialog = await openParagraph(launched.app, page, 'P1')
    await expect(box(dialog)).toBeChecked()
    await box(dialog).uncheck()
    await accept(dialog)
    dialog = await openParagraph(launched.app, page, 'P2')
    await expect(box(dialog)).not.toBeChecked()
    await box(dialog).check()
    await accept(dialog)
    for (const label of ['P1', 'P2', 'P3', 'P5'] as const)
      expect(await lineUnder(page, label), label).toEqual(before[label])

    await saveFromMenu(launched.app)
    await expect
      .poll(async () => pPrOf(docPath, 'P1'), { timeout: 20_000 })
      .toContain('<w:adjustRightInd w:val="0"/>')
    expect(await pPrOf(docPath, 'P2')).toContain('<w:adjustRightInd/>')
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
