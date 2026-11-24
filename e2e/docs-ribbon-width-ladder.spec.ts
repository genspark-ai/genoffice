import { test, expect, type Page } from '@playwright/test'
import { mkdirSync, mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'

/**
 * Word takes each ribbon group down a ladder as the window narrows:
 * full-height buttons, a stack of small captioned controls, icons alone,
 * one menu button for the group.
 */

const SHOTS = process.env.RIBBON_LADDER_SHOTS

async function minimalDocx(): Promise<Buffer> {
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
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  )
  const style = (id: string, name: string) =>
    `<w:style w:type="paragraph"${id === 'Normal' ? ' w:default="1"' : ''} w:styleId="${id}"><w:name w:val="${name}"/><w:qFormat/></w:style>`
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${style('Normal', 'Normal')}${[1, 2, 3, 4].map((n) => style(`Heading${n}`, `heading ${n}`)).join('')}${style('Title', 'Title')}${style('Subtitle', 'Subtitle')}</w:styles>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function showTab(page: Page, name: string): Promise<void> {
  const tab = page.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).first()
  // clicking the active tab collapses the ribbon, so only click an inactive one
  if (!/\bactive\b/.test((await tab.getAttribute('class')) ?? '')) await tab.click()
  await expect(page.locator('[data-ribbon-body] button').first()).toBeVisible()
}

type Form = 'buttons' | 'stack' | 'icons' | 'menu'

/** the form of the stack that holds the control with this caption */
const formOf = (page: Page, caption: string): Promise<Form> =>
  page.evaluate((text) => {
    const label = [...document.querySelectorAll('[data-ribbon-body] .rb-small-label')].find(
      (el) => el.textContent?.trim() === text,
    )
    const stack = label?.closest('.rb-col')
    if (!stack) throw new Error(`no stack holds ${text}`)
    if (stack.closest('.rb-group-menu')) return 'menu'
    if (stack.classList.contains('rb-col-wide')) return 'buttons'
    return stack.classList.contains('rb-col-tight') ? 'icons' : 'stack'
  }, caption)

const overflow = (page: Page): Promise<number> =>
  page.evaluate(() => {
    const body = document.querySelector('[data-ribbon-body]')!
    return body.scrollWidth - body.clientWidth
  })

test.describe('Docs ribbon follows the window width', () => {
  let dir: string
  test.beforeEach(() => {
    dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-ribbon-ladder-')))
  })
  test.afterEach(() => {
    rmSync(dir, { recursive: true, force: true })
  })

  test('groups go from buttons to stacks to icons, and the style cards stand in for the Styles button', async () => {
    test.setTimeout(180_000)
    const docPath = join(dir, 'ladder.docx')
    writeFileSync(docPath, await minimalDocx())
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-ribbon-width-ladder',
      openFile: docPath,
    })
    const { app } = launched
    try {
      const page = await waitForPageWithUrl(app, '://docs/')
      await page.waitForFunction(
        () => Boolean((window as never as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
        undefined,
        { timeout: 30_000 },
      )
      await page.locator('.doc-page p', { hasText: 'Body' }).first().click()
      const resize = async (width: number) => {
        await app.evaluate(({ BrowserWindow }, w) => {
          for (const win of BrowserWindow.getAllWindows()) win.setContentSize(w, 800)
        }, width)
        // a screen narrower than the width asked for keeps the window at its own
        await expect
          .poll(() => page.evaluate(() => window.innerWidth), { timeout: 5_000 })
          .toBe(width)
          .catch(() => {})
      }
      const shot = async (name: string) => {
        if (!SHOTS) return
        mkdirSync(SHOTS, { recursive: true })
        await page
          .locator('.ribbon')
          .first()
          .screenshot({ path: join(SHOTS, `${name}.png`) })
      }
      const ladder: Record<string, Form[]> = { pages: [], links: [], text: [], proofing: [] }
      const widths = [1700, 1500, 1300, 1100, 900]
      const cardCounts: number[] = []
      for (const width of widths) {
        await resize(width)
        await showTab(page, 'Insert')
        await expect
          .poll(() => overflow(page), { message: `Insert at ${width}` })
          .toBeLessThanOrEqual(0)
        ladder.pages.push(await formOf(page, 'Blank Page'))
        ladder.links.push(await formOf(page, 'Bookmark'))
        ladder.text.push(await formOf(page, 'Date & Time'))
        await shot(`insert-${width}`)
        await showTab(page, 'Review')
        ladder.proofing.push(await formOf(page, 'Word Count'))
        await shot(`review-${width}`)
        await showTab(page, 'Home')
        // the cards and the Styles button are one or the other
        const styles = () =>
          page.evaluate(() => ({
            cards: document.querySelectorAll('.style-gallery-inline .style-card').length,
            button: document.querySelectorAll('.rb-styles-gallery.rb-big').length,
            opener: document.querySelectorAll('.rb-styles-gallery.rb-styles-more').length,
          }))
        // the count settles a render after the resize
        await expect
          .poll(
            async () => {
              const { cards, button, opener } = await styles()
              return cards === 0
                ? button === 1 && opener === 0
                : cards >= 3 && button === 0 && opener === 1
            },
            { message: `Home at ${width}: ${JSON.stringify(await styles())}` },
          )
          .toBe(true)
        cardCounts.push((await styles()).cards)
        await shot(`home-${width}`)
      }
      if (SHOTS)
        writeFileSync(
          join(SHOTS, 'ladder.json'),
          JSON.stringify({ widths, ladder, cardCounts }, null, 2),
        )
      // a form never comes back as the window narrows
      const step = { buttons: 0, stack: 1, icons: 2, menu: 3 }
      for (const forms of Object.values(ladder)) {
        const steps = forms.map((form) => step[form])
        expect(steps).toEqual([...steps].sort((a, b) => a - b))
      }
      expect(ladder.pages[0]).toBe('buttons')
      expect(ladder.links[0]).toBe('buttons')
      expect(ladder.pages.at(-1)).not.toBe('buttons')
      // Word folds Links before Pages
      for (let i = 0; i < widths.length; i += 1) {
        expect(step[ladder.links[i]]).toBeGreaterThanOrEqual(step[ladder.pages[i]])
      }

      expect(cardCounts.at(-1)).toBe(0)

      // off macOS Home also carries the Editing group, so the cards need a wider window there
      const cards = page.locator('.style-gallery-inline .style-card')
      let shown = 0
      for (const width of [1700, 2000, 2400]) {
        await resize(width)
        await showTab(page, 'Home')
        await expect
          .poll(() => cards.count(), { timeout: 3_000 })
          .toBeGreaterThan(0)
          .catch(() => {})
        shown = await cards.count()
        if (shown > 0) break
      }
      expect(shown).toBeGreaterThanOrEqual(3)
      // the opener at the end of the cards drops the whole gallery
      await page.locator('.rb-styles-gallery.rb-styles-more').click()
      await expect(page.locator('.style-gallery-menu .style-card').first()).toBeVisible()
    } finally {
      await closeAndSaveVideo(launched, 'docs-ribbon-width-ladder')
    }
  })
})
