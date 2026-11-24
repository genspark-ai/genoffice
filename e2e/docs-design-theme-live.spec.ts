import { test, expect } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtempSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'

const STYLES = `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:color w:val="4472C4" w:themeColor="accent1"/><w:sz w:val="32"/></w:rPr></w:style></w:styles>`

const slot = (tag: string, hex: string) => `<a:${tag}><a:srgbClr val="${hex}"/></a:${tag}>`
const fontGroup = (tag: string, face: string) =>
  `<a:${tag}><a:latin typeface="${face}"/><a:ea typeface=""/><a:cs typeface=""/></a:${tag}>`
const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
const three = (xml: string) => xml.repeat(3)
// Word refuses a theme part that has no format scheme
const FMT_SCHEME =
  `<a:fmtScheme name="Office"><a:fillStyleLst>${three(solid)}</a:fillStyleLst>` +
  `<a:lnStyleLst>${three(`<a:ln w="9525">${solid}</a:ln>`)}</a:lnStyleLst>` +
  `<a:effectStyleLst>${three('<a:effectStyle><a:effectLst/></a:effectStyle>')}</a:effectStyleLst>` +
  `<a:bgFillStyleLst>${three(solid)}</a:bgFillStyleLst></a:fmtScheme>`
const theme = (fmtScheme: string) =>
  `<a:theme ${A} name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>${slot('dk2', '44546A')}${slot('lt2', 'E7E6E6')}${slot('accent1', '4472C4')}${slot('accent2', 'ED7D31')}${slot('accent3', 'A5A5A5')}${slot('accent4', 'FFC000')}${slot('accent5', '5B9BD5')}${slot('accent6', '70AD47')}${slot('hlink', '0563C1')}${slot('folHlink', '954F72')}</a:clrScheme><a:fontScheme name="Office">${fontGroup('majorFont', 'Calibri Light')}${fontGroup('minorFont', 'Calibri')}</a:fontScheme>${fmtScheme}</a:themeElements></a:theme>`
const THEME = theme(FMT_SCHEME)

const BODY =
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p>' +
  '<w:p><w:r><w:rPr><w:color w:val="ED7D31" w:themeColor="accent2"/></w:rPr><w:t>tinted</w:t></w:r></w:p>' +
  '<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>literal</w:t></w:r></w:p>' +
  para('plain')

async function themedDocx(themeXml = THEME): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await docxWith(BODY))
  zip.file('word/styles.xml', STYLES)
  zip.file('word/theme/theme1.xml', themeXml)
  const part = async (path: string, edit: (xml: string) => string) =>
    zip.file(path, edit(await zip.file(path)!.async('string')))
  await part('[Content_Types].xml', (xml) =>
    xml.replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
    ),
  )
  await part('word/_rels/document.xml.rels', (xml) =>
    xml.replace(
      '</Relationships>',
      '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('Design > Themes: a pick repaints the page and the style set cards at once, and saves', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-theme-live-')))
  const docPath = join(dir, 'themed.docx')
  writeFileSync(docPath, await themedDocx())
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-design-theme-live',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const line = (text: string) => page.locator('.doc-page :is(h1, p)', { hasText: text }).first()
    const style = (text: string, prop: 'color' | 'fontFamily', inner = false) =>
      line(text).evaluate(
        (el, [prop, inner]) =>
          getComputedStyle((inner && el.querySelector('span')) || el)[prop as 'color'],
        [prop, inner] as const,
      )
    await expect(line('Overview')).toBeVisible({ timeout: 30_000 })
    await line('plain').click()
    await expect.poll(() => style('Overview', 'color')).toBe('rgb(68, 114, 196)')
    await expect.poll(() => style('tinted', 'color', true)).toBe('rgb(237, 125, 49)')

    await page.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
    const cardHeading = page.getByTestId('style-set-current').locator('.style-set-page > span')
    await expect
      .poll(() => cardHeading.nth(1).evaluate((el) => el.style.color))
      .toBe('rgb(68, 114, 196)')
    await page.screenshot({ path: test.info().outputPath('theme-before.png') })

    const themes = page.locator('.rb-big', { has: page.getByText('Themes', { exact: true }) })
    await themes.click()
    const menu = page.locator('[data-rb-panel]')
    await expect(menu.locator('button.active')).toHaveText(/Office/)
    await page.screenshot({ path: test.info().outputPath('theme-menu.png') })
    await menu.locator('button', { hasText: 'Facet' }).click()

    // Facet: first accent 90C226, second 54A021, Trebuchet MS for headings and body
    await expect.poll(() => style('Overview', 'color')).toBe('rgb(144, 194, 38)')
    await expect.poll(() => style('tinted', 'color', true)).toBe('rgb(84, 160, 33)')
    await expect.poll(() => style('literal', 'color', true)).toBe('rgb(255, 0, 0)')
    await expect.poll(() => style('Overview', 'fontFamily')).toContain('Trebuchet MS')
    await expect.poll(() => style('plain', 'fontFamily')).toContain('Trebuchet MS')
    await expect
      .poll(() => cardHeading.nth(1).evaluate((el) => el.style.color))
      .toBe('rgb(144, 194, 38)')
    await themes.click()
    await expect(menu.locator('button.active')).toHaveText(/Facet/)
    await themes.click()
    await expect(menu).toBeHidden()
    await page.screenshot({ path: test.info().outputPath('theme-after.png') })

    // the ribbon reads the body face off the picked theme too
    await line('plain').click()
    await page.locator('.ribbon-tab', { hasText: /^Home$/ }).click()
    await expect(page.locator('input.rb-font-family')).toHaveValue(/Trebuchet MS/)

    await saveFromMenu(app)
    await expect
      .poll(async () => {
        const zip = await JSZip.loadAsync(readFileSync(docPath))
        const theme = (await zip.file('word/theme/theme1.xml')?.async('string')) ?? ''
        const body = (await zip.file('word/document.xml')?.async('string')) ?? ''
        return [
          /<a:accent1>\s*<a:srgbClr val="90C226"/.test(theme),
          /<w:color\b[^>]*w:themeColor="accent2"/.test(body),
          // a whole theme brings its own format scheme in place of the file's
          !theme.includes(FMT_SCHEME) && theme.includes('<a:fmtScheme name="Facet">'),
        ]
      })
      .toEqual([true, true, true])
    // the saved file reads back the way the page already looked
    await expect.poll(() => style('Overview', 'color')).toBe('rgb(144, 194, 38)')
    await expect.poll(() => style('tinted', 'color', true)).toBe('rgb(84, 160, 33)')
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a theme part that came without a format scheme is saved with one once a theme is picked', async () => {
  test.setTimeout(180_000)
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-theme-live-')))
  const docPath = join(dir, 'bare-theme.docx')
  writeFileSync(docPath, await themedDocx(theme('')))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-design-theme-live',
    openFile: docPath,
  })
  try {
    const { app } = launched
    const page = await waitForPageWithUrl(app, '://docs/')
    const heading = page.locator('.doc-page h1', { hasText: 'Overview' }).first()
    await expect(heading).toBeVisible({ timeout: 30_000 })
    await page.locator('.doc-page p', { hasText: 'plain' }).first().click()
    await page.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
    await page.locator('.rb-big', { has: page.getByText('Themes', { exact: true }) }).click()
    await page.locator('[data-rb-panel] button', { hasText: 'Facet' }).click()
    await expect
      .poll(() => heading.evaluate((el) => getComputedStyle(el).color))
      .toBe('rgb(144, 194, 38)')
    await saveFromMenu(app)
    const saved = async () =>
      (await (
        await JSZip.loadAsync(readFileSync(docPath))
      )
        .file('word/theme/theme1.xml')
        ?.async('string')) ?? ''
    await expect.poll(saved).toContain('<a:srgbClr val="90C226"')
    // the three parts the schema asks for, in its order, the format scheme with its four lists
    expect(await saved()).toMatch(
      /<a:themeElements><a:clrScheme [\s\S]*<\/a:clrScheme><a:fontScheme [\s\S]*<\/a:fontScheme><a:fmtScheme name="Facet"><a:fillStyleLst>[\s\S]*<\/a:fillStyleLst><a:lnStyleLst>[\s\S]*<\/a:lnStyleLst><a:effectStyleLst>[\s\S]*<\/a:effectStyleLst><a:bgFillStyleLst>[\s\S]*<\/a:bgFillStyleLst><\/a:fmtScheme><\/a:themeElements>/,
    )
  } finally {
    await closeAndSaveVideo(launched)
    rmSync(dir, { recursive: true, force: true })
  }
})
