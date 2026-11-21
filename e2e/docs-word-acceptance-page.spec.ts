/** Word acceptance tasks, page layout: each walked along Word's path and read back from the saved file. */
import { test, expect, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { para, saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  clickMenu,
  open,
  part,
  ribbon,
  ribbonTab,
  savedBody,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance-page'
const PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

/** opens a ribbon gallery and picks the entry whose name starts with `item` */
const pick = async (page: Page, control: string, item: string) => {
  await ribbon(page, control).click()
  await page.locator('[data-rb-panel] button', { hasText: new RegExp(`^${item}`) }).click()
}

/** the saved header and footer parts, by the type their reference in the body gives them */
async function savedHeadersFooters(docPath: string) {
  const zip = await JSZip.loadAsync(readFileSync(docPath))
  const read = (name: string) => zip.file(name)?.async('string') ?? Promise.resolve('')
  const body = await read('word/document.xml')
  const rels = await read('word/_rels/document.xml.rels')
  const out: Record<string, string> = {}
  for (const ref of body.match(/<w:(header|footer)Reference\b[^>]*\/>/g) ?? []) {
    const kind = /<w:(header|footer)/.exec(ref)![1]
    const type = /w:type="([^"]+)"/.exec(ref)?.[1] ?? ''
    const id = /r:id="([^"]+)"/.exec(ref)?.[1] ?? ''
    const target = new RegExp(
      `Id="${id}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${id}"`,
    ).exec(rels)
    out[`${kind}:${type}`] = await read(`word/${target?.[1] ?? target?.[2]}`)
  }
  return out
}

test('Layout > Margins, Orientation, Size and Columns reach the section properties', async () => {
  test.setTimeout(180_000)
  const s = await open(VIDEO_DIR, ['alpha', 'beta'].map(para).join(''), 'layout.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'alpha' })).toBeVisible({ timeout: 30_000 })
    await caretAtEnd(page, 'alpha')
    await ribbonTab(page, 'Layout')
    await pick(page, 'Margins', 'Narrow')
    await pick(page, 'Orientation', 'Landscape')
    await pick(page, 'Size', 'A4')
    await pick(page, 'Columns', 'Two')
    await page.screenshot({ path: test.info().outputPath('layout.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const sectPr = /<w:sectPr\b[\s\S]*<\/w:sectPr>/.exec(await savedBody(docPath))?.[0] ?? ''
          const attrs = (tag: string) =>
            Object.fromEntries(
              Array.from(
                (new RegExp(`<w:${tag}\\b[^>]*>`).exec(sectPr)?.[0] ?? '').matchAll(
                  /w:(\w+)="([^"]*)"/g,
                ),
                (m) => [m[1], m[2]],
              ),
            )
          const { w, h, orient } = attrs('pgSz')
          const { top, right, bottom, left } = attrs('pgMar')
          return { w, h, orient, top, right, bottom, left, columns: attrs('cols').num }
        },
        { timeout: 30_000 },
      )
      .toEqual({
        w: '16838',
        h: '11906',
        orient: 'landscape',
        top: '720',
        right: '720',
        bottom: '720',
        left: '720',
        columns: '2',
      })
  } finally {
    await s.close()
  }
})

test('Insert > Header takes typed text and a different first page, Insert > Page Numbers centres a number in the footer, Design > Watermark adds a text watermark', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    para('first page') + PAGE_BREAK + para('second page'),
    'header.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'second page' })).toBeAttached({
      timeout: 30_000,
    })
    await caretAtEnd(page, 'first page')

    await ribbonTab(page, 'Insert')
    await pick(page, 'Header', 'Edit Header')
    const editor = page.locator('.page-hf-header .page-hf-editor')
    await expect(editor).toBeFocused()
    await page.keyboard.type('Quarterly report')
    await page
      .locator('[data-ribbon-body] button', { hasText: /^\s*Different First Page\s*$/ })
      .click()
    // ticking it ends the header editing here; Word stays in the header until it is closed
    await expect(editor).toHaveCount(0)
    const headers = page.locator('.page-hf-header')
    await expect(headers.nth(1)).toContainText('Quarterly report')
    await expect(headers.nth(0)).not.toContainText('Quarterly report')

    await clickMenu(app, ['Insert', 'Page Numbers…'])
    const numbers = page.locator('.page-numbers-dialog')
    await expect(numbers.getByLabel('Position')).toHaveValue('footer')
    await numbers.getByLabel('Alignment').selectOption('center')
    await numbers.locator('button.primary').click()
    await expect(numbers).toBeHidden()
    await expect(page.locator('.page-hf-footer').nth(1)).toHaveText('2')

    await ribbonTab(page, 'Design')
    await pick(page, 'Watermark', 'Custom Watermark')
    const watermark = page.locator('.watermark-dialog')
    await watermark.locator('.wm-mode', { hasText: 'Text' }).locator('input').check()
    await watermark.locator('.wm-row input[type=text]').first().fill('INTERNAL')
    await watermark.locator('.btn-primary').click()
    await expect(watermark).toBeHidden()
    await page.screenshot({ path: test.info().outputPath('header-numbers-watermark.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const parts = await savedHeadersFooters(docPath)
          const footer = parts['footer:default'] ?? ''
          return {
            header: /Quarterly report/.test(parts['header:default'] ?? ''),
            firstHeader: /Quarterly report/.test(parts['header:first'] ?? ''),
            titlePg: /<w:titlePg\/>/.test(await savedBody(docPath)),
            pageField: /<w:instrText[^>]*>\s*PAGE\b|w:instr="\s*PAGE\b/.test(footer),
            centred: /<w:jc w:val="center"\/>/.test(footer),
            watermark: Object.entries(parts).some(
              ([type, xml]) => type.startsWith('header:') && /INTERNAL/.test(xml),
            ),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({
        header: true,
        firstHeader: false,
        titlePg: true,
        pageField: true,
        centred: true,
        watermark: true,
      })
    expect(await part(docPath, 'word/document.xml')).toContain('second page')
  } finally {
    await s.close()
  }
})
