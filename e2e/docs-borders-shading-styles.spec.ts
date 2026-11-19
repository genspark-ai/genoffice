/**
 * Format > Borders and Shading: the line style, colour, width and gap of a
 * paragraph's borders and the pattern of its shading show on the page, are
 * saved as w:pBdr / w:shd and are read back; a border a file brought keeps
 * what was not changed; Apply to > Text writes the character border and fill.
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
  '</w:sectPr>'

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

const documentXml = async (docPath: string) =>
  (await JSZip.loadAsync(readFileSync(docPath))).file('word/document.xml')!.async('string')

const paragraph = (page: Page, text: string) =>
  page.locator('.doc-page p', { hasText: text }).first()

const look = (page: Page, text: string) =>
  paragraph(page, text).evaluate((el) => {
    const s = getComputedStyle(el)
    return {
      top: `${s.borderTopWidth} ${s.borderTopStyle} ${s.borderTopColor}`,
      left: `${s.borderLeftWidth} ${s.borderLeftStyle} ${s.borderLeftColor}`,
      bottom: `${s.borderBottomWidth} ${s.borderBottomStyle} ${s.borderBottomColor}`,
      right: `${s.borderRightWidth} ${s.borderRightStyle} ${s.borderRightColor}`,
      padTop: Math.round(parseFloat(s.paddingTop) * 100) / 100,
      fill: s.backgroundColor,
    }
  })

async function openDialog(app: ElectronApplication, page: Page, text: string) {
  await paragraph(page, text).click()
  await clickMenu(app, ['Format', 'Borders and Shading…'])
  const dialog = page.locator('.borders-shading-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

test('a styled, coloured, wider box and a patterned shading show, save and come back', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-borders-shading-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana', 'cherry'].map(para).join('') + SECT_PR))
  const original = await documentXml(docPath)
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-borders-shading-styles',
    openFile: docPath,
  })
  try {
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(paragraph(page, 'apple')).toBeVisible({ timeout: 30_000 })
    let dialog = await openDialog(launched.app, page, 'apple')
    const tab = (name: string) => page.locator('.bs-tabs button', { hasText: name })
    const setting = (id: string) => dialog.locator(`button[data-setting="${id}"]`)
    const lineStyle = (val: string) => dialog.locator(`button[data-line-style="${val}"]`)

    // no text is selected: there is no text to apply to
    await expect(dialog.locator('select[name="apply-to"] option[value="text"]')).toHaveAttribute(
      'disabled',
      '',
    )
    await expect(setting('none')).toHaveAttribute('aria-pressed', 'true')

    await lineStyle('double').click()
    await dialog.locator('input[name="border-color"]').fill('#c00000')
    await dialog.locator('select[name="border-width"]').selectOption('1.5')
    await setting('box').click()
    for (const side of ['t', 'b', 'l', 'r'])
      await expect(dialog.locator(`button[data-side="${side}"]`)).toHaveAttribute(
        'aria-pressed',
        'true',
      )
    await dialog.locator('.pb-options-toggle').click()
    await expect(dialog.locator('input[name="space-l"]')).toBeEnabled()
    await dialog.locator('input[name="space-t"]').fill('4')

    await tab('Shading').click()
    await dialog.locator('button[data-fill="FFF2CC"]').click()
    await expect(dialog.locator('input[name="pattern-color"]')).toBeDisabled()
    await dialog.locator('select[name="pattern"]').selectOption('pct25')
    await dialog.locator('input[name="pattern-color"]').fill('#0070c0')
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toBeHidden()

    const line = '6px double rgb(192, 0, 0)'
    // 25% of 0070C0 over FFF2CC
    const shown = {
      top: line,
      left: line,
      bottom: line,
      right: line,
      padTop: 5.33,
      fill: 'rgb(191, 210, 201)',
    }
    await expect.poll(() => look(page, 'apple')).toEqual(shown)

    await saveFromMenu(launched.app)
    const side = (name: string, space: number) =>
      `<w:${name} w:val="double" w:sz="12" w:space="${space}" w:color="C00000"/>`
    const pPr =
      `<w:pPr><w:pBdr>${side('top', 4)}${side('left', 0)}${side('bottom', 0)}${side('right', 0)}</w:pBdr>` +
      '<w:shd w:val="pct25" w:color="0070C0" w:fill="FFF2CC"/></w:pPr>'
    await expect
      .poll(async () => (await documentXml(docPath)).includes(pPr), { timeout: 20_000 })
      .toBe(true)
    // the paragraphs that were not touched are the bytes they were
    expect(await documentXml(docPath)).toBe(
      original.replace('<w:p><w:r><w:t>apple', `<w:p>${pPr}<w:r><w:t xml:space="preserve">apple`),
    )

    await closeAndSaveVideo(launched, 'docs-borders-shading-styles')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-borders-shading-styles-reopen',
      openFile: docPath,
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(paragraph(page, 'apple')).toBeVisible({ timeout: 30_000 })
    await expect.poll(() => look(page, 'apple')).toEqual(shown)

    dialog = await openDialog(launched.app, page, 'apple')
    await expect(setting('box')).toHaveAttribute('aria-pressed', 'true')
    await expect(lineStyle('double')).toHaveAttribute('aria-selected', 'true')
    await expect(dialog.locator('input[name="border-color"]')).toHaveValue('#c00000')
    await expect(dialog.locator('select[name="border-width"]')).toHaveValue('1.5')
    await dialog.locator('.pb-options-toggle').click()
    await expect(dialog.locator('input[name="space-t"]')).toHaveValue('4')
    await expect(dialog.locator('input[name="space-b"]')).toHaveValue('0')
    await tab('Shading').click()
    await expect(dialog.locator('button[data-fill="FFF2CC"]')).toHaveClass(/active/)
    await expect(dialog.locator('select[name="pattern"]')).toHaveValue('pct25')
    await expect(dialog.locator('input[name="pattern-color"]')).toHaveValue('#0070c0')
    // OK with nothing changed leaves the document clean and the file as it is
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => look(page, 'apple')).toEqual(shown)
  } finally {
    await closeAndSaveVideo(launched, 'docs-borders-shading-styles')
    rmSync(dir, { recursive: true, force: true })
  }
})

test('Shadow, 3-D and a hatch pattern show on the page, save and come back', async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-borders-fx-')))
  const docPath = join(dir, 'letter.docx')
  writeFileSync(docPath, await docxWith(['apple', 'banana', 'cherry'].map(para).join('') + SECT_PR))
  let launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-borders-shading-fx',
    openFile: docPath,
  })
  const drawn = (page: Page, text: string) =>
    paragraph(page, text).evaluate((el) => {
      const s = getComputedStyle(el)
      return {
        top: `${s.borderTopWidth} ${s.borderTopStyle}`,
        shadow: s.boxShadow,
        hatch:
          s.backgroundImage.startsWith('repeating-linear-gradient(135deg') &&
          s.backgroundImage.includes('rgb(0, 112, 192)')
            ? 'up-diagonal blue'
            : s.backgroundImage,
        fill: s.backgroundColor,
      }
    })
  const shown = {
    apple: {
      top: '2px solid',
      shadow: 'rgb(192, 0, 0) 4px 4px 0px 0px',
      hatch: 'none',
      fill: 'rgba(0, 0, 0, 0)',
    },
    banana: { top: '3px ridge', shadow: 'none', hatch: 'none', fill: 'rgba(0, 0, 0, 0)' },
    cherry: {
      top: '0px none',
      // the shading box widened by no gap: not a drop shadow
      shadow: 'rgb(255, 242, 204) 0px 0px 0px 0px, rgb(255, 242, 204) 0px 0px 0px 0px',
      hatch: 'up-diagonal blue',
      fill: 'rgb(255, 242, 204)',
    },
  }
  const allDrawn = async (page: Page) => ({
    apple: await drawn(page, 'apple'),
    banana: await drawn(page, 'banana'),
    cherry: await drawn(page, 'cherry'),
  })
  try {
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(paragraph(page, 'apple')).toBeVisible({ timeout: 30_000 })
    const setting = (dialog: ReturnType<Page['locator']>, id: string) =>
      dialog.locator(`button[data-setting="${id}"]`)
    const ok = async (dialog: ReturnType<Page['locator']>) => {
      await dialog.locator('button[type="submit"]').click()
      await expect(dialog).toBeHidden()
    }

    let dialog = await openDialog(launched.app, page, 'apple')
    await dialog.locator('input[name="border-color"]').fill('#c00000')
    await dialog.locator('select[name="border-width"]').selectOption('1.5')
    await setting(dialog, 'shadow').click()
    await expect(setting(dialog, 'shadow')).toHaveAttribute('aria-pressed', 'true')
    await page.screenshot({ path: test.info().outputPath('borders-shadow-dialog.png') })
    await ok(dialog)

    dialog = await openDialog(launched.app, page, 'banana')
    await setting(dialog, '3d').click()
    await ok(dialog)

    dialog = await openDialog(launched.app, page, 'cherry')
    await page.locator('.bs-tabs button', { hasText: 'Shading' }).click()
    await dialog.locator('button[data-fill="FFF2CC"]').click()
    await dialog.locator('select[name="pattern"]').selectOption('diagStripe')
    await dialog.locator('input[name="pattern-color"]').fill('#0070c0')
    await page.screenshot({ path: test.info().outputPath('shading-hatch-dialog.png') })
    await ok(dialog)

    await expect.poll(() => allDrawn(page)).toEqual(shown)
    await page.screenshot({ path: test.info().outputPath('borders-fx-page.png') })

    await saveFromMenu(launched.app)
    const box = (attrs: string) =>
      `<w:pBdr>${['top', 'left', 'bottom', 'right'].map((n) => `<w:${n} ${attrs}/>`).join('')}</w:pBdr>`
    const expected = [
      box('w:val="single" w:sz="12" w:space="0" w:color="C00000" w:shadow="1"'),
      box('w:val="single" w:sz="4" w:space="0" w:color="auto" w:frame="1"'),
      '<w:shd w:val="diagStripe" w:color="0070C0" w:fill="FFF2CC"/>',
    ]
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return expected.filter((part) => !xml.includes(part))
        },
        { timeout: 20_000 },
      )
      .toEqual([])

    await closeAndSaveVideo(launched, 'docs-borders-shading-fx')
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-borders-shading-fx-reopen',
      openFile: docPath,
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(paragraph(page, 'apple')).toBeVisible({ timeout: 30_000 })
    await expect.poll(() => allDrawn(page)).toEqual(shown)
    // closing the dialog hands the caret back a moment later: a click made before that
    // lands on the paragraph the dialog had, so each look reopens until it is on its own
    const reopened = (
      text: string,
      check: (dialog: ReturnType<Page['locator']>) => Promise<void>,
    ) =>
      expect(async () => {
        const opened = await openDialog(launched.app, page, text)
        try {
          await check(opened)
        } finally {
          await opened.locator('button', { hasText: 'Cancel' }).click()
          await expect(opened).toBeHidden()
        }
      }).toPass({ timeout: 30_000 })
    await reopened('apple', (d) =>
      expect(setting(d, 'shadow')).toHaveAttribute('aria-pressed', 'true', { timeout: 2_000 }),
    )
    await reopened('banana', (d) =>
      expect(setting(d, '3d')).toHaveAttribute('aria-pressed', 'true', { timeout: 2_000 }),
    )
    await reopened('cherry', async (d) => {
      await page.locator('.bs-tabs button', { hasText: 'Shading' }).click()
      await expect(d.locator('select[name="pattern"]')).toHaveValue('diagStripe', {
        timeout: 2_000,
      })
    })
  } finally {
    await closeAndSaveVideo(launched, 'docs-borders-shading-fx')
    rmSync(dir, { recursive: true, force: true })
  }
})

test("a new colour keeps a file border's style and width, and Apply to Text boxes and fills the selection", async () => {
  test.setTimeout(240_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-borders-shading-')))
  const docPath = join(dir, 'letter.docx')
  const dashed = (color: string) =>
    ['top', 'left', 'bottom', 'right']
      .map((name) => `<w:${name} w:val="dashed" w:sz="18" w:space="1" w:color="${color}"/>`)
      .join('')
  const shd = '<w:shd w:val="pct20" w:color="auto" w:fill="DDEBF7"/>'
  const wave =
    '<w:p><w:pPr><w:pBdr><w:bottom w:val="wave" w:sz="6" w:space="1" w:color="FF0000"/></w:pBdr></w:pPr>' +
    '<w:r><w:t>banana</w:t></w:r></w:p>'
  writeFileSync(
    docPath,
    await docxWith(
      `<w:p><w:pPr><w:pBdr>${dashed('00B050')}</w:pBdr>${shd}</w:pPr><w:r><w:t>apple</w:t></w:r></w:p>` +
        wave +
        para('cherry') +
        SECT_PR,
    ),
  )
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-borders-shading-styles',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    await expect(paragraph(page, 'apple')).toBeVisible({ timeout: 30_000 })
    // the file's dashes draw as dashes, 20% black over the fill behind them
    const before = await look(page, 'apple')
    expect(before.top).toBe('3px dashed rgb(0, 176, 80)')
    expect(before.fill).toBe('rgb(177, 188, 198)')

    let dialog = await openDialog(app, page, 'apple')
    const lineStyle = (val: string) => dialog.locator(`button[data-line-style="${val}"]`)
    await expect(dialog.locator('button[data-setting="box"]')).toHaveAttribute(
      'aria-pressed',
      'true',
    )
    await expect(lineStyle('dashed')).toHaveAttribute('aria-selected', 'true')
    await expect(dialog.locator('select[name="border-width"]')).toHaveValue('2.25')
    await expect(dialog.locator('input[name="border-color"]')).toHaveValue('#00b050')
    await dialog.locator('input[name="border-color"]').fill('#7030a0')
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(async () => (await look(page, 'apple')).top)
      .toBe('3px dashed rgb(112, 48, 160)')
    expect((await look(page, 'apple')).fill).toBe(before.fill)

    // a style the list does not offer is shown as what the file has, and is not replaced
    dialog = await openDialog(app, page, 'banana')
    await expect(lineStyle('wave')).toHaveAttribute('aria-selected', 'true')
    await expect(lineStyle('wave')).toHaveText('Wave')
    await page.keyboard.press('Escape')
    await expect(dialog).toBeHidden()

    // Apply to > Text: the selected word gets the line as its character border, and the fill
    await paragraph(page, 'cherry').dblclick()
    await clickMenu(app, ['Format', 'Borders and Shading…'])
    await expect(dialog).toBeVisible()
    await dialog.locator('select[name="apply-to"]').selectOption('text')
    // a character border is one plain box
    await expect(dialog.locator('button[data-setting="shadow"]')).toBeDisabled()
    await expect(dialog.locator('button[data-setting="3d"]')).toBeDisabled()
    await expect(dialog.locator('button[data-setting="custom"]')).toBeDisabled()
    await expect(dialog.locator('button[data-side="t"]')).toBeDisabled()
    await expect(lineStyle('double')).toHaveCount(0)
    await lineStyle('dashed').click()
    await dialog.locator('input[name="border-color"]').fill('#c00000')
    await dialog.locator('select[name="border-width"]').selectOption('1.5')
    await dialog.locator('button[data-setting="box"]').click()
    await page.locator('.bs-tabs button', { hasText: 'Shading' }).click()
    await expect(dialog.locator('select[name="pattern"]')).toBeDisabled()
    await dialog.locator('button[data-fill="FFF2CC"]').click()
    await dialog.locator('button[type="submit"]').click()
    await expect(dialog).toBeHidden()
    const word = paragraph(page, 'cherry').locator('span', { hasText: 'cherry' }).first()
    await expect
      .poll(() =>
        word.evaluate((el) => {
          const s = getComputedStyle(el)
          return `${s.borderTopStyle} ${s.borderTopColor} ${s.backgroundColor}`
        }),
      )
      .toBe('dashed rgb(192, 0, 0) rgb(255, 242, 204)')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const xml = await documentXml(docPath)
          return [
            xml.includes(`<w:pBdr>${dashed('7030A0')}</w:pBdr>${shd}`),
            xml.includes(wave),
            /<w:bdr w:val="dashed" w:sz="12" w:space="0" w:color="C00000"\/>/.test(xml),
            /<w:rPr>.*<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"\/>.*<\/w:rPr><w:t[^>]*>cherry/.test(
              xml,
            ),
          ]
        },
        { timeout: 20_000 },
      )
      .toEqual([true, true, true, true])
  } finally {
    await closeAndSaveVideo(launched, 'docs-borders-shading-styles')
    rmSync(dir, { recursive: true, force: true })
  }
})
