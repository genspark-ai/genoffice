/** Word acceptance tasks, text formatting: each walked along Word's path and read back from the saved file. */
import { test, expect } from '@playwright/test'
import { para, saveFromMenu } from './docs-objects-fixture'
import {
  open,
  pickStyle,
  ribbon,
  ribbonTab,
  savedParagraph,
  selectParagraph,
  textBox,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance-text'

test('Home > Bold, Italic, Underline and their shortcuts, Font and Font Size', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    ['alpha', 'beta', 'gamma', 'delta', 'epsilon'].map(para).join(''),
    'characters.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'alpha' })).toBeVisible({ timeout: 30_000 })
    await ribbonTab(page, 'Home')

    await selectParagraph(page, 'alpha')
    await ribbon(page, 'Bold').click()
    await selectParagraph(page, 'beta')
    await ribbon(page, 'Italic').click()
    await selectParagraph(page, 'gamma')
    await ribbon(page, 'Underline').click()

    await selectParagraph(page, 'delta')
    await page.keyboard.press('ControlOrMeta+b')
    await page.keyboard.press('ControlOrMeta+i')
    await page.keyboard.press('ControlOrMeta+u')

    await selectParagraph(page, 'epsilon')
    const font = page.locator('[data-ribbon-body] input[aria-label="Font"]')
    await font.click()
    await page.keyboard.type('Georgia')
    await page.keyboard.press('Enter')
    await expect(font).toHaveValue('Georgia')
    const size = page.locator('[data-ribbon-body] input[aria-label="Font Size"]')
    await size.click()
    await page.keyboard.type('18')
    await page.keyboard.press('Enter')
    await expect(size).toHaveValue('18')
    await expect(
      page.locator('.doc-page p', { hasText: 'epsilon' }).locator('span').first(),
    ).toHaveCSS('font-size', '24px')
    await page.screenshot({ path: test.info().outputPath('characters.png') })

    await saveFromMenu(app)
    const marks = (p: string) => ({
      bold: /<w:b\/>/.test(p),
      italic: /<w:i\/>/.test(p),
      underline: /<w:u w:val="single"\/>/.test(p),
    })
    await expect
      .poll(
        async () => {
          const [alpha, beta, gamma, delta, epsilon] = await Promise.all(
            ['alpha', 'beta', 'gamma', 'delta', 'epsilon'].map((text) =>
              savedParagraph(docPath, text),
            ),
          )
          return {
            alpha: marks(alpha),
            beta: marks(beta),
            gamma: marks(gamma),
            delta: marks(delta),
            font: /<w:rFonts\b[^>]*w:ascii="Georgia"/.test(epsilon),
            size: /<w:sz w:val="36"\/>/.test(epsilon),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({
        alpha: { bold: true, italic: false, underline: false },
        beta: { bold: false, italic: true, underline: false },
        gamma: { bold: false, italic: false, underline: true },
        delta: { bold: true, italic: true, underline: true },
        font: true,
        size: true,
      })
  } finally {
    await s.close()
  }
})

test('Home > Clear All Formatting returns a formatted heading to the body style', async () => {
  test.setTimeout(180_000)
  const s = await open(VIDEO_DIR, ['alpha', 'beta'].map(para).join(''), 'clear.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'alpha' })).toBeVisible({ timeout: 30_000 })
    await ribbonTab(page, 'Home')

    await selectParagraph(page, 'alpha')
    await pickStyle(page, 'Heading 1')
    await selectParagraph(page, 'alpha')
    await page.keyboard.press('ControlOrMeta+i')
    await ribbon(page, 'Center').click()
    await saveFromMenu(app)
    await expect
      .poll(() => savedParagraph(docPath, 'alpha'), { timeout: 30_000 })
      .toMatch(/<w:pStyle w:val="Heading1"\/>[\s\S]*<w:jc w:val="center"\/>[\s\S]*<w:i\/>/)

    await selectParagraph(page, 'alpha')
    await ribbon(page, 'Clear Formatting').click()
    await page.screenshot({ path: test.info().outputPath('cleared.png') })
    await page.locator('.doc-page p', { hasText: 'beta' }).click()
    await page.keyboard.press('End')
    await page.keyboard.type('!')
    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const alpha = await savedParagraph(docPath, 'alpha')
          return {
            saved: (await savedParagraph(docPath, 'beta!')) !== '',
            heading: /<w:pStyle w:val="Heading1"\/>/.test(alpha),
            centred: /<w:jc w:val="center"\/>/.test(alpha),
            italic: /<w:i\/>/.test(alpha),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ saved: true, heading: false, centred: false, italic: false })
  } finally {
    await s.close()
  }
})

test('Home > Format Painter carries a paragraph’s formatting to another paragraph', async () => {
  test.setTimeout(180_000)
  const s = await open(VIDEO_DIR, ['alpha', 'beta', 'gamma'].map(para).join(''), 'painter.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'alpha' })).toBeVisible({ timeout: 30_000 })
    await ribbonTab(page, 'Home')

    await selectParagraph(page, 'alpha')
    await page.keyboard.press('ControlOrMeta+b')
    await ribbon(page, 'Center').click()

    // the whole paragraph is the source, so its paragraph formatting comes along
    await selectParagraph(page, 'alpha')
    const painter = ribbon(page, 'Format Painter')
    await painter.click()
    await expect(painter).toHaveAttribute('aria-pressed', 'true')
    const target = await textBox(page, 'gamma')
    await page.mouse.move(target.x + 1, target.y + target.height / 2)
    await page.mouse.down()
    await page.mouse.move(target.x + target.width - 1, target.y + target.height / 2, { steps: 5 })
    await page.mouse.up()
    await expect(painter).toHaveAttribute('aria-pressed', 'false')
    await page.screenshot({ path: test.info().outputPath('painted.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const [beta, gamma] = await Promise.all(
            ['beta', 'gamma'].map((text) => savedParagraph(docPath, text)),
          )
          return {
            gammaBold: /<w:b\/>/.test(gamma),
            gammaCentred: /<w:jc w:val="center"\/>/.test(gamma),
            betaTouched: /<w:b\/>|<w:jc /.test(beta),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ gammaBold: true, gammaCentred: true, betaTouched: false })
  } finally {
    await s.close()
  }
})
