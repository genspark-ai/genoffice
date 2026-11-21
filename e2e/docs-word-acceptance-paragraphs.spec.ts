/** Word acceptance tasks, paragraphs and lists: each walked along Word's path and read back from the saved file. */
import { test, expect } from '@playwright/test'
import { para, saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  clickMenu,
  open,
  part,
  ribbon,
  ribbonTab,
  savedBody,
  savedParagraph,
  selectParagraph,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance-paragraphs'

/** level and number format of the list the saved paragraph holding `text` is in */
async function savedListItem(docPath: string, text: string) {
  const paragraph = await savedParagraph(docPath, text)
  const numId = /<w:numId w:val="(\d+)"\/>/.exec(paragraph)?.[1]
  if (!numId || numId === '0') return null
  const level = /<w:ilvl w:val="(\d)"\/>/.exec(paragraph)?.[1] ?? '0'
  const numbering = await part(docPath, 'word/numbering.xml')
  const abstractId = new RegExp(
    `<w:num\\b[^>]*w:numId="${numId}"[^>]*>\\s*<w:abstractNumId w:val="(\\d+)"/>`,
  ).exec(numbering)?.[1]
  const abstract =
    new RegExp(
      `<w:abstractNum\\b[^>]*w:abstractNumId="${abstractId}"[\\s\\S]*?</w:abstractNum>`,
    ).exec(numbering)?.[0] ?? ''
  const lvl = new RegExp(`<w:lvl\\b[^>]*w:ilvl="${level}"[\\s\\S]*?</w:lvl>`).exec(abstract)?.[0]
  return { numId, level, format: /<w:numFmt w:val="(\w+)"\/>/.exec(lvl ?? '')?.[1] }
}

test('Format > Paragraph sets a first-line indent and spacing, Format > Tabs a right tab stop, Home > Sort orders paragraphs', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    ['intro', 'tabbed', 'pear', 'apple', 'mango'].map(para).join(''),
    'paragraphs.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'intro' })).toBeVisible({ timeout: 30_000 })

    await selectParagraph(page, 'intro')
    await clickMenu(app, ['Format', 'Paragraph…'])
    const paragraph = page.locator('.modal', { hasText: 'Paragraph' })
    await paragraph.locator('button.gs-dd-btn[aria-label="Special"]').click()
    await page.locator('.gs-dd-pop [role="option"]', { hasText: /^First line$/ }).click()
    await paragraph.getByLabel('By', { exact: true }).fill('0.5 in')
    await paragraph.getByLabel('Space before').fill('12')
    await paragraph.getByLabel('Space after').fill('6')
    await paragraph.getByLabel('Space after').blur()
    await page.screenshot({ path: test.info().outputPath('paragraph-dialog.png') })
    await paragraph.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(paragraph).toBeHidden()
    await expect(page.locator('.doc-page p', { hasText: 'intro' })).toHaveCSS('text-indent', '48px')

    await selectParagraph(page, 'tabbed')
    await clickMenu(app, ['Format', 'Tabs…'])
    const tabs = page.locator('.tabs-dialog')
    await tabs.getByLabel('Tab stop position').fill('4 in')
    await tabs.getByLabel('Right', { exact: true }).check()
    await tabs.getByRole('button', { name: 'Set', exact: true }).click()
    await expect(tabs.locator('.tabs-dialog-list button')).toHaveCount(1)
    await tabs.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(tabs).toBeHidden()

    await selectParagraph(page, 'pear', 'mango')
    await ribbonTab(page, 'Home')
    await ribbon(page, 'Sort').click()
    const sort = page.locator('.modal', { hasText: 'Sort' })
    await expect(sort).toBeVisible()
    await sort.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(sort).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const intro = await savedParagraph(docPath, 'intro')
          const body = await savedBody(docPath)
          return {
            firstLine: /<w:ind\b[^>]*w:firstLine="720"/.test(intro),
            before: /<w:spacing\b[^>]*w:before="240"/.test(intro),
            after: /<w:spacing\b[^>]*w:after="120"/.test(intro),
            tab: /<w:tabs>\s*<w:tab w:val="right" w:pos="5760"\/>\s*<\/w:tabs>/.test(
              await savedParagraph(docPath, 'tabbed'),
            ),
            order: ['apple', 'mango', 'pear']
              .map((text) => body.indexOf(`>${text}<`))
              .every((at, i, all) => at > 0 && (i === 0 || at > all[i - 1])),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ firstLine: true, before: true, after: true, tab: true, order: true })
  } finally {
    await s.close()
  }
})

test('Home > Bullets and Numbering: Enter continues the list, Enter on an empty item leaves it, Tab and Shift+Tab change the level', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    ['one', 'two', 'three', 'four', 'tail'].map(para).join(''),
    'lists.docx',
  )
  try {
    const { app, page, docPath } = s
    const items = page.locator('.doc-page .doc-li')
    await expect(page.locator('.doc-page p', { hasText: 'one' })).toBeVisible({ timeout: 30_000 })
    await ribbonTab(page, 'Home')

    await selectParagraph(page, 'one')
    await ribbon(page, 'Bullets').click()
    await expect(items).toHaveCount(1)
    await caretAtEnd(page, 'one')
    await page.keyboard.press('Enter')
    await page.keyboard.type('added')
    await expect(items).toHaveCount(2)
    await page.keyboard.press('Enter')
    await expect(items).toHaveCount(3)
    await page.keyboard.press('Enter')
    await expect(items).toHaveCount(2)
    await page.keyboard.type('plain')

    await selectParagraph(page, 'two', 'four')
    await ribbon(page, 'Numbering').click()
    await expect(items).toHaveCount(5)
    for (const text of ['three', 'four']) {
      await selectParagraph(page, text)
      await page.keyboard.press('ArrowLeft')
      await page.keyboard.press('Tab')
    }
    await page.screenshot({ path: test.info().outputPath('lists.png') })
    await saveFromMenu(app)
    await expect
      .poll(
        async () => ({
          one: (await savedListItem(docPath, 'one'))?.format,
          added: (await savedListItem(docPath, 'added'))?.format,
          sameBullets:
            (await savedListItem(docPath, 'one'))?.numId ===
            (await savedListItem(docPath, 'added'))?.numId,
          plain: await savedListItem(docPath, 'plain'),
          two: await savedListItem(docPath, 'two').then(
            (item) => item && [item.level, item.format],
          ),
          three: (await savedListItem(docPath, 'three'))?.level,
          four: (await savedListItem(docPath, 'four'))?.level,
        }),
        { timeout: 30_000 },
      )
      .toEqual({
        one: 'bullet',
        added: 'bullet',
        sameBullets: true,
        plain: null,
        two: ['0', 'decimal'],
        three: '1',
        four: '1',
      })

    await selectParagraph(page, 'four')
    await page.keyboard.press('ArrowLeft')
    await page.keyboard.press('Shift+Tab')
    await saveFromMenu(app)
    await expect
      .poll(
        async () => [
          (await savedListItem(docPath, 'three'))?.level,
          (await savedListItem(docPath, 'four'))?.level,
        ],
        { timeout: 30_000 },
      )
      .toEqual(['1', '0'])
  } finally {
    await s.close()
  }
})
