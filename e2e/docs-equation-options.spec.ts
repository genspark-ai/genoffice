/**
 * Format > Equation Options: the document's default justification, margins and
 * n-ary limit placement (settings.xml m:mathPr) move the display equations that
 * do not set their own, are saved, and come back on reopen and in the PDF.
 */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const SECT_PR =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '<w:cols w:space="720"/></w:sectPr>'

const M_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
const run = (text: string) => `<m:r><m:t>${text}</m:t></m:r>`
const mathPara = (inner: string, pr = '') =>
  `<w:p><m:oMathPara ${M_NS}>${pr}${inner}</m:oMathPara></w:p>`
const SUM =
  '<m:oMath><m:nary><m:naryPr><m:chr m:val="\u2211"/></m:naryPr>' +
  `<m:sub>${run('i')}</m:sub><m:sup>${run('n')}</m:sup><m:e>${run('x')}</m:e></m:nary></m:oMath>`

const BODY =
  para('apple') +
  mathPara(`<m:oMath>${run('a+b=c')}</m:oMath>`) +
  mathPara(
    `<m:oMath>${run('d+e=f')}</m:oMath>`,
    '<m:oMathParaPr><m:jc m:val="right"/></m:oMathParaPr>',
  ) +
  mathPara(SUM) +
  mathPara(`<m:oMath>${run('g=h')}</m:oMath><m:oMath>${run('g+h+k+m+p=q')}</m:oMath>`) +
  SECT_PR

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

interface Placed {
  /** each line's distance from the equation paragraph's left and right edges, px */
  lines: Array<{ left: number; right: number }>
  limits: string
}

/** where every display equation's lines sit in their paragraph */
const placed = (page: Page): Promise<Placed[]> =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll('.ProseMirror .doc-protected-formula-display')).map(
      (el) => {
        const box = el.getBoundingClientRect()
        return {
          lines: Array.from(el.querySelectorAll('.doc-formula-math math')).map((math) => {
            const r = math.getBoundingClientRect()
            return { left: Math.round(r.left - box.left), right: Math.round(box.right - r.right) }
          }),
          limits: el.querySelector('munderover')
            ? 'undOvr'
            : el.querySelector('msubsup')
              ? 'subSup'
              : '',
        }
      },
    ),
  )

async function pdfTextLeft(path: string): Promise<number> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(path)) }).promise
  const content = await (await doc.getPage(1)).getTextContent()
  const lefts = content.items
    .filter(
      (it): it is { transform: number[]; str: string } =>
        'transform' in it && 'str' in it && (it as { str: string }).str.trim().length > 0,
    )
    .filter((it) => it.str !== 'apple')
    .map((it) => it.transform[4])
  return Math.min(...lefts)
}

test('Equation Options moves the display equations that follow the document, saves and reopens', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-equation-options-')))
  const docPath = join(dir, 'equations.docx')
  writeFileSync(docPath, await docxWith(BODY))
  const dialogOf = (page: Page) => page.locator('.equation-options-dialog')
  const open = async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-equation-options',
      openFile: docPath,
    })
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'apple' }).first()).toBeVisible({
      timeout: 30_000,
    })
    await expect.poll(async () => (await placed(page)).length).toBe(4)
    await page.locator('.doc-page p', { hasText: 'apple' }).first().click()
    return { launched, page }
  }
  // 1 inch at 96 px
  const INCH = 96
  const moved = async (page: Page) => {
    const [byDefault, own, sum, twoLines] = await placed(page)
    expect(byDefault.lines[0].left).toBe(INCH)
    expect(own.lines[0].right).toBe(0)
    expect(sum.limits).toBe('subSup')
    expect(sum.lines[0].left).toBe(INCH)
    expect(twoLines.lines.map((l) => l.left)).toEqual([INCH, INCH])
  }

  let session = await open()
  try {
    const { launched, page } = session
    const before = await placed(page)
    const [byDefault, own, sum, twoLines] = before
    // centred as a group: one line sits in the middle, two lines share a left edge inside it
    expect(Math.abs(byDefault.lines[0].left - byDefault.lines[0].right)).toBeLessThanOrEqual(1)
    expect(own.lines[0].right).toBe(0)
    expect(sum.limits).toBe('undOvr')
    expect(twoLines.lines).toHaveLength(2)
    expect(twoLines.lines[0].left).toBe(twoLines.lines[1].left)
    expect(Math.abs(twoLines.lines[1].left - twoLines.lines[1].right)).toBeLessThanOrEqual(1)

    await page.screenshot({ path: test.info().outputPath('equations-before.png') })
    await clickMenu(launched.app, ['Format', 'Equation Options…'])
    const dialog = dialogOf(page)
    await expect(dialog).toBeVisible()
    await expect(dialog.locator('select[name="defJc"]')).toHaveValue('centerGroup')
    await expect(dialog.locator('input[name="naryLim"][value="undOvr"]')).toBeChecked()
    await expect(dialog.locator('input[name="intLim"][value="subSup"]')).toBeChecked()
    await expect(dialog.locator('select[name="brkBin"]')).toHaveValue('before')
    await expect(dialog.locator('select[name="mathFont"]')).toHaveValue('Cambria Math')
    await expect(dialog.locator('input[name="smallFrac"]')).not.toBeChecked()
    await expect(dialog.locator('input[name="wrap"]').first()).toBeChecked()

    // each line centred on its own
    await dialog.locator('select[name="defJc"]').selectOption('center')
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(async () => {
        const lines = (await placed(page))[3].lines
        return lines.map((l) => Math.abs(l.left - l.right) <= 1)
      })
      .toEqual([true, true])
    expect((await placed(page))[3].lines[0].left).toBeGreaterThan(twoLines.lines[0].left)
    expect((await placed(page))[1]).toEqual(own)

    // left, one inch in, n-ary limits to the side
    await clickMenu(launched.app, ['Format', 'Equation Options…'])
    await expect(dialog).toBeVisible()
    await dialog.locator('select[name="defJc"]').selectOption('left')
    const leftMargin = dialog.getByLabel('Left margin:', { exact: true })
    await leftMargin.fill('1in')
    await leftMargin.blur()
    await dialog.locator('input[name="naryLim"][value="subSup"]').check()
    await page.screenshot({ path: test.info().outputPath('equation-options.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await placed(page))[0].lines[0].left).toBe(INCH)
    await moved(page)
    // the equation with its own m:jc stayed at the right edge
    expect((await placed(page))[1]).toEqual(own)
    await page.screenshot({ path: test.info().outputPath('equations-left.png') })

    await saveFromMenu(launched.app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          return (await zip.file('word/settings.xml')?.async('string')) ?? ''
        },
        { timeout: 30_000 },
      )
      .toContain(
        `<w:settings ${M_NS} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
          '<m:mathPr><m:lMargin m:val="1440"/><m:defJc m:val="left"/><m:naryLim m:val="subSup"/></m:mathPr>',
      )
    const saved = await JSZip.loadAsync(readFileSync(docPath))
    const documentXml = await saved.file('word/document.xml')!.async('string')
    expect(documentXml.match(/<m:jc /g)).toHaveLength(1)
    expect(documentXml).not.toContain('<m:limLoc')
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-equation-options')
  }

  session = await open()
  try {
    const { launched, page } = session
    await moved(page)
    await clickMenu(launched.app, ['Format', 'Equation Options…'])
    const dialog = dialogOf(page)
    await expect(dialog).toBeVisible()
    await expect(dialog.locator('select[name="defJc"]')).toHaveValue('left')
    await expect(dialog.getByLabel('Left margin:', { exact: true })).toHaveValue(/^(1"|2\.54 cm)$/)
    await expect(dialog.getByLabel('Right margin:', { exact: true })).toHaveValue(/^0(")?( cm)?$/)
    await expect(dialog.locator('input[name="naryLim"][value="subSup"]')).toBeChecked()
    await expect(dialog.locator('input[name="dispDef"]')).toBeChecked()

    // off: the defaults step back and the equations follow their (left-aligned) paragraphs
    await dialog.locator('input[name="dispDef"]').uncheck()
    await expect(dialog.locator('select[name="defJc"]')).toBeDisabled()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await placed(page))[0].lines[0].left).toBe(0)
    expect((await placed(page))[1].lines[0].right).toBe(0)

    await clickMenu(launched.app, ['Format', 'Equation Options…'])
    await dialog.locator('input[name="dispDef"]').check()
    await dialog.locator('button.primary').click()
    await expect.poll(async () => (await placed(page))[0].lines[0].left).toBe(INCH)

    // the exported PDF draws the equations where the page shows them: margin + 1 inch
    const pdfPath = join(dir, 'equations.pdf')
    await launched.app.evaluate(({ dialog: d, shell }, out) => {
      d.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    expect(Math.abs((await pdfTextLeft(pdfPath)) - 144)).toBeLessThanOrEqual(2)
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-equation-options')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a nested fraction keeps its size until Reduce size of nested fractions is ticked, and the choice is saved', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-equation-frac-')))
  const docPath = join(dir, 'fractions.docx')
  const frac = (num: string, den: string) => `<m:f><m:num>${num}</m:num><m:den>${den}</m:den></m:f>`
  writeFileSync(
    docPath,
    await docxWith(
      para('apple') +
        mathPara(`<m:oMath>${frac(run('a'), frac(run('b'), run('c')))}</m:oMath>`) +
        SECT_PR,
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-equation-options-frac',
    openFile: docPath,
  })
  /** font sizes of the outer numerator and of the nested fraction's numerator, px */
  const sizes = (page: Page) =>
    page.evaluate(() => {
      const math = document.querySelector('.ProseMirror .doc-formula-math math')
      const px = (el: Element | null | undefined) =>
        el ? Math.round(parseFloat(getComputedStyle(el).fontSize) * 100) / 100 : 0
      const outer = math?.querySelector('mfrac')
      return {
        outer: px(outer?.firstElementChild),
        nested: px(outer?.querySelector('mfrac')?.firstElementChild),
      }
    })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.ProseMirror .doc-formula-math math mfrac mfrac')).toBeVisible({
      timeout: 30_000,
    })
    const before = await sizes(page)
    expect(before.outer).toBeGreaterThan(0)
    expect(before.nested).toBe(before.outer)
    await page.screenshot({ path: test.info().outputPath('nested-fraction-full.png') })

    await clickMenu(launched.app, ['Format', 'Equation Options…'])
    const dialog = page.locator('.equation-options-dialog')
    await expect(dialog).toBeVisible()
    await dialog.locator('input[name="smallFrac"]').check()
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await sizes(page)).nested).toBeLessThan(before.outer)
    expect((await sizes(page)).outer).toBe(before.outer)
    await page.screenshot({ path: test.info().outputPath('nested-fraction-small.png') })

    await saveFromMenu(launched.app)
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(readFileSync(docPath))
          return (await zip.file('word/settings.xml')?.async('string')) ?? ''
        },
        { timeout: 30_000 },
      )
      .toContain('<m:mathPr><m:smallFrac/></m:mathPr>')
  } finally {
    await closeAndSaveVideo(launched, 'docs-equation-options-frac')
    rmSync(dir, { recursive: true, force: true })
  }
})
