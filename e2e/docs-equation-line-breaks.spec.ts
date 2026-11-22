/**
 * A display equation wider than its column wraps at its top-level operators.
 * Format > Equation Options decides which line the operator is on, how a
 * subtraction is repeated and where the wrapped lines sit (settings.xml
 * m:brkBin / m:brkBinSub / m:wrapIndent / m:wrapRight); the equation in the
 * file stays whole.
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
const MINUS = '−'
const run = (text: string) => `<m:r><m:t>${text}</m:t></m:r>`
const mathPara = (inner: string) => `<w:p><m:oMathPara ${M_NS}>${inner}</m:oMathPara></w:p>`
const terms = (n: number, sign: string) =>
  Array.from({ length: n }, (_, i) => `x${i}y${i}z${i}`).join(sign)

const FITS = `<m:oMath>${run('a+b=c')}</m:oMath>`
const SUM = `<m:oMath>${run(`S=${terms(14, '+')}`)}</m:oMath>`
const DIFF = `<m:oMath>${run(`T=${terms(14, MINUS)}`)}</m:oMath>`
const MANUAL = `<m:oMath>${run('g=h')}</m:oMath><m:oMath>${run('g+h+k+m+p=q')}</m:oMath>`

const BODY =
  para('apple') +
  mathPara(FITS) +
  mathPara(SUM) +
  mathPara(DIFF) +
  mathPara(MANUAL) +
  para('pear') +
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

interface Line {
  text: string
  /** distance from the equation paragraph's left and right edges, px */
  left: number
  right: number
}

interface Equation {
  html: string
  top: number
  bottom: number
  lines: Line[]
}

const equations = (page: Page): Promise<Equation[]> =>
  page.evaluate(() =>
    Array.from(document.querySelectorAll('.ProseMirror .doc-protected-formula-display')).map(
      (el) => {
        const box = el.getBoundingClientRect()
        return {
          html: el.querySelector('.doc-formula-math')?.innerHTML ?? '',
          top: box.top,
          bottom: box.bottom,
          lines: Array.from(el.querySelectorAll('.doc-formula-math math')).map((math) => {
            const r = math.getBoundingClientRect()
            return {
              text: math.textContent ?? '',
              left: Math.round(r.left - box.left),
              right: Math.round(box.right - r.right),
            }
          }),
        }
      },
    ),
  )

/** the equation's lines once it has `count` of them or more */
async function wrapped(page: Page, index: number, count = 2): Promise<Line[]> {
  await expect
    .poll(async () => (await equations(page))[index]?.lines.length)
    .toBeGreaterThanOrEqual(count)
  return (await equations(page))[index].lines
}

const FITS_MATHML =
  '<math display="block"><mrow><mi>a</mi><mo>+</mo><mi>b</mi><mo>=</mo><mi>c</mi></mrow></math>'

async function pdfText(path: string): Promise<{ text: string; maxRight: number; rows: number }> {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(path)) }).promise
  const content = await (await doc.getPage(1)).getTextContent()
  const items = content.items.filter(
    (it): it is { transform: number[]; str: string; width: number } =>
      'transform' in it && 'str' in it && (it as { str: string }).str.trim().length > 0,
  )
  return {
    // math italics back to plain letters
    text: items
      .map((it) => it.str.replace(/\s/g, ''))
      .join('')
      .normalize('NFKC'),
    maxRight: Math.max(...items.map((it) => it.transform[4] + it.width)),
    rows: new Set(items.map((it) => Math.round(it.transform[5]))).size,
  }
}

test('an over-wide display equation wraps at its operators as Equation Options says, and the file keeps it whole', async () => {
  test.setTimeout(300_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-equation-breaks-')))
  const docPath = join(dir, 'wide-equations.docx')
  writeFileSync(docPath, await docxWith(BODY))
  const dialogOf = (page: Page) => page.locator('.equation-options-dialog')
  const open = async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-equation-line-breaks',
      openFile: docPath,
    })
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'apple' }).first()).toBeVisible({
      timeout: 30_000,
    })
    await expect.poll(async () => (await equations(page)).length).toBe(4)
    await page.locator('.doc-page p', { hasText: 'apple' }).first().click()
    return { launched, page }
  }
  const apply = async (
    session: { launched: { app: ElectronApplication }; page: Page },
    change: (dialog: ReturnType<typeof dialogOf>) => Promise<void>,
  ) => {
    await clickMenu(session.launched.app, ['Format', 'Equation Options…'])
    const dialog = dialogOf(session.page)
    await expect(dialog).toBeVisible()
    await change(dialog)
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
  }
  const INCH = 96

  let session = await open()
  try {
    const { launched, page } = session
    const sum = await wrapped(page, 1)
    const [fits, , , manual] = await equations(page)

    // what fits is the markup it always was, centred, and a manual break is two plain lines
    expect(fits.html).toBe(FITS_MATHML)
    expect(Math.abs(fits.lines[0].left - fits.lines[0].right)).toBeLessThanOrEqual(1)
    expect(manual.lines.map((l) => l.text)).toEqual(['g=h', 'g+h+k+m+p=q'])
    expect(manual.html).not.toContain('doc-math-wrap')

    // inside the column; the operator starts the wrapped line, which is indented one inch
    for (const line of sum) {
      expect(line.left).toBeGreaterThanOrEqual(0)
      expect(line.right).toBeGreaterThanOrEqual(0)
    }
    expect(sum.map((l) => l.text).join('')).toBe(`S=${terms(14, '+')}`)
    expect(sum[0].text.endsWith('+')).toBe(false)
    expect(sum[1].text.startsWith('+')).toBe(true)
    expect(sum[1].left - sum[0].left).toBe(INCH)

    // the taller equations push what follows down instead of covering it
    const all = await equations(page)
    for (let i = 1; i < all.length; i++)
      expect(all[i].top).toBeGreaterThanOrEqual(all[i - 1].bottom - 0.5)
    const pearTop = await page
      .locator('.doc-page p', { hasText: 'pear' })
      .first()
      .evaluate((el) => el.getBoundingClientRect().top)
    expect(pearTop).toBeGreaterThanOrEqual(all[3].bottom - 0.5)
    await page.screenshot({ path: test.info().outputPath('wrapped-before.png') })

    await clickMenu(launched.app, ['Format', 'Equation Options…'])
    const dialog = dialogOf(page)
    await expect(dialog).toBeVisible()
    await expect(dialog.locator('select[name="brkBin"]')).toHaveValue('before')
    await expect(dialog.locator('select[name="brkBin"]')).toBeEnabled()
    // only a repeated operator has a second sign to choose
    await expect(dialog.locator('select[name="brkBinSub"]')).toBeDisabled()
    await expect(dialog.locator('input[name="wrap"][value="indent"]')).toBeChecked()
    await expect(dialog.getByRole('textbox', { name: 'Indent wrapped lines by:' })).toHaveValue(
      /^(1"|2\.54 cm)$/,
    )
    await dialog.locator('select[name="brkBin"]').selectOption('after')
    await page.screenshot({ path: test.info().outputPath('equation-options-breaks.png') })
    await dialog.locator('button.primary').click()
    await expect(dialog).toBeHidden()
    await expect.poll(async () => (await wrapped(page, 1))[0].text.endsWith('+')).toBe(true)
    expect((await wrapped(page, 1))[1].text.startsWith('+')).toBe(false)
    expect((await wrapped(page, 1)).map((l) => l.text).join('')).toBe(`S=${terms(14, '+')}`)

    // on both lines; a subtraction ends its line as a minus and starts the next as a plus
    await apply(session, async (d) => {
      await d.locator('select[name="brkBin"]').selectOption('repeat')
      await expect(d.locator('select[name="brkBinSub"]')).toBeEnabled()
      await d.locator('select[name="brkBinSub"]').selectOption('-+')
    })
    await expect.poll(async () => (await wrapped(page, 1))[1].text.startsWith('+')).toBe(true)
    expect((await wrapped(page, 1))[0].text.endsWith('+')).toBe(true)
    const diff = await wrapped(page, 2)
    expect(diff[0].text.endsWith(MINUS)).toBe(true)
    expect(diff[1].text.startsWith('+')).toBe(true)
    // inside the lines the subtractions are untouched
    expect(diff[0].text.slice(0, -1)).not.toContain('+')

    await apply(session, async (d) => {
      const indent = d.getByRole('textbox', { name: 'Indent wrapped lines by:' })
      await indent.fill('0.5in')
      await indent.blur()
    })
    await expect
      .poll(async () => {
        const lines = await wrapped(page, 1)
        return lines[1].left - lines[0].left
      })
      .toBe(INCH / 2)

    await apply(session, async (d) => {
      await d.locator('input[name="wrap"][value="right"]').check()
    })
    await expect.poll(async () => (await wrapped(page, 1))[1].right).toBe(0)
    expect((await wrapped(page, 2))[1].right).toBe(0)
    expect((await equations(page))[0].html).toBe(FITS_MATHML)
    await page.screenshot({ path: test.info().outputPath('wrapped-repeat-right.png') })

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
        '<m:mathPr><m:brkBin m:val="repeat"/><m:brkBinSub m:val="-+"/><m:wrapRight/></m:mathPr>',
      )
    const saved = await JSZip.loadAsync(readFileSync(docPath))
    const documentXml = await saved.file('word/document.xml')!.async('string')
    for (const omml of [FITS, SUM, DIFF, MANUAL]) expect(documentXml).toContain(omml)
    expect(documentXml.match(/<m:oMath[\s>]/g)).toHaveLength(5)
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-equation-line-breaks')
  }

  session = await open()
  try {
    const { launched, page } = session
    await expect.poll(async () => (await wrapped(page, 1))[1].right).toBe(0)
    const diff = await wrapped(page, 2)
    expect(diff[0].text.endsWith(MINUS)).toBe(true)
    expect(diff[1].text.startsWith('+')).toBe(true)

    await apply(session, async (d) => {
      await expect(d.locator('select[name="brkBin"]')).toHaveValue('repeat')
      await expect(d.locator('select[name="brkBinSub"]')).toHaveValue('-+')
      await expect(d.locator('input[name="wrap"][value="right"]')).toBeChecked()
      await expect(d.getByRole('textbox', { name: 'Indent wrapped lines by:' })).toBeDisabled()
      await d.locator('input[name="wrap"][value="indent"]').check()
      await d.locator('select[name="brkBin"]').selectOption('before')
    })
    await expect
      .poll(async () => {
        const lines = await wrapped(page, 1)
        return lines[1].left - lines[0].left
      })
      .toBe(INCH)

    // the exported PDF breaks the equations too: nothing runs past the right margin
    const pdfPath = join(dir, 'wide-equations.pdf')
    await launched.app.evaluate(({ dialog: d, shell }, out) => {
      d.showSaveDialog = async () => ({ canceled: false, filePath: out })
      shell.showItemInFolder = () => {}
    }, pdfPath)
    expect(
      await page.evaluate(() =>
        (window as unknown as { __exportPdf: () => Promise<boolean> }).__exportPdf(),
      ),
    ).toBe(true)
    const pdf = await pdfText(pdfPath)
    expect(pdf.maxRight).toBeLessThanOrEqual(612 - 72 + 2)
    expect(pdf.text).toContain('x13y13z13')
    const screenRows = (await equations(page)).reduce((n, eq) => n + eq.lines.length, 0)
    // apple, pear and every equation line
    expect(pdf.rows).toBe(screenRows + 2)
  } finally {
    await closeAndSaveVideo(session.launched, 'docs-equation-line-breaks')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('pagination counts the lines an equation wraps into', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-equation-pages-')))
  const docPath = join(dir, 'many-wide-equations.docx')
  // 30 one-line equations fit one page; wrapped they need more
  writeFileSync(
    docPath,
    await docxWith(
      para('apple') +
        Array.from({ length: 30 }, () => mathPara(SUM)).join('') +
        para('pear') +
        SECT_PR,
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-equation-line-breaks-pages',
    openFile: docPath,
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-page p', { hasText: 'apple' }).first()).toBeVisible({
      timeout: 30_000,
    })
    await expect.poll(async () => (await equations(page)).length).toBe(30)
    await expect
      .poll(async () => (await equations(page)).every((eq) => eq.lines.length >= 2))
      .toBe(true)
    // how many lines, and so pages, depends on the math font the machine has
    const pageTurns = page.locator('.doc-page .page-gap')
    await expect.poll(() => pageTurns.count(), { timeout: 30_000 }).toBeGreaterThanOrEqual(1)
    const layout = () =>
      page.evaluate(() => {
        const gaps = Array.from(document.querySelectorAll('.doc-page .page-gap')).map((el) =>
          el.getBoundingClientRect(),
        )
        const boxes = Array.from(
          document.querySelectorAll('.ProseMirror .doc-protected-formula-display'),
        ).map((el) => el.getBoundingClientRect())
        const pear = Array.from(document.querySelectorAll('.doc-page p'))
          .find((p) => p.textContent === 'pear')!
          .getBoundingClientRect()
        return {
          straddling: boxes.filter((b) =>
            gaps.some((gap) => b.top < gap.bottom - 0.5 && b.bottom > gap.top + 0.5),
          ).length,
          above: boxes.filter((b) => b.bottom <= gaps[0].top + 0.5).length,
          pearBelow: pear.top >= gaps[gaps.length - 1].bottom - 0.5,
        }
      })
    // no equation is cut by a page turn, and the first page holds only what fits above it
    await expect.poll(async () => (await layout()).straddling).toBe(0)
    const settled = await layout()
    expect(settled.above).toBeGreaterThan(0)
    expect(settled.above).toBeLessThan(30)
    expect(settled.pearBelow).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'docs-equation-line-breaks-pages')
    rmSync(dir, { recursive: true, force: true })
  }
})
