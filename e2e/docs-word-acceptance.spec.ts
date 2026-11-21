/**
 * Word acceptance tasks that no other spec walks: each is done the way a Word
 * user does it (ribbon tab, menu path or shortcut) and checked in the saved
 * file. The rest of the acceptance list is in the docs-word-acceptance-*
 * specs, one per group of the list.
 */
import { test, expect } from '@playwright/test'
import { mkdtempSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl } from './helpers'
import { docxWith, para, saveFromMenu } from './docs-objects-fixture'
import {
  SECT_PR,
  clickMenu,
  part,
  ribbon,
  ribbonTab,
  savedParagraph,
  selectParagraph,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance'

async function open(body: string, name: string) {
  const dir = realpathSync(mkdtempSync(join(tmpdir(), 'genoffice-e2e-acceptance-')))
  const docPath = join(dir, name)
  writeFileSync(docPath, await docxWith(body + SECT_PR))
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: VIDEO_DIR,
    openFile: docPath,
  })
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  return { dir, docPath, launched, page }
}

test('text colour, highlight, clear formatting, alignment, line spacing and the Paragraph dialog', async () => {
  test.setTimeout(180_000)
  const { dir, docPath, launched, page } = await open(
    ['alpha', 'beta', 'gamma', 'delta'].map(para).join(''),
    'text.docx',
  )
  try {
    const { app } = launched
    await expect(page.locator('.doc-page p', { hasText: 'alpha' })).toBeVisible({ timeout: 30_000 })
    await ribbonTab(page, 'Home')

    // Home > Font Color and Text Highlight Color put their current colours on the selection
    await selectParagraph(page, 'alpha')
    await ribbon(page, 'Font Color').click()
    await ribbon(page, 'Text Highlight Color').click()

    // Home > Clear Formatting takes direct formatting off again
    await selectParagraph(page, 'beta')
    await page.keyboard.press('ControlOrMeta+b')
    await expect(page.locator('.doc-page p', { hasText: 'beta' }).locator('strong')).toHaveCount(1)
    await ribbon(page, 'Clear Formatting').click()
    await expect(page.locator('.doc-page p', { hasText: 'beta' }).locator('strong')).toHaveCount(0)

    // Home > Center and Justify
    await ribbon(page, 'Center').click()
    await selectParagraph(page, 'gamma')
    await ribbon(page, 'Justify').click()

    // Home > Line and Paragraph Spacing > 1.5
    await ribbon(page, 'Line and Paragraph Spacing').click()
    await page.locator('.spacing-menu button', { hasText: /^1\.5$/ }).click()

    // Format > Paragraph…: space before and after
    await selectParagraph(page, 'delta')
    await clickMenu(app, ['Format', 'Paragraph…'])
    const dialog = page.locator('.modal', { hasText: 'Paragraph' })
    await expect(dialog).toBeVisible()
    await dialog.getByLabel('Space before').fill('12')
    await dialog.getByLabel('Space after').fill('6')
    await dialog.getByLabel('Space after').blur()
    await page.screenshot({ path: test.info().outputPath('paragraph-dialog.png') })
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const [alpha, beta, gamma, delta] = await Promise.all(
            ['alpha', 'beta', 'gamma', 'delta'].map((text) => savedParagraph(docPath, text)),
          )
          return {
            colour: /<w:color w:val="[0-9A-Fa-f]{6}"/.test(alpha),
            highlight: /<w:highlight w:val="\w+"/.test(alpha),
            betaBold: /<w:b\/>/.test(beta),
            betaCentred: /<w:jc w:val="center"\/>/.test(beta),
            gammaJustified: /<w:jc w:val="both"\/>/.test(gamma),
            gammaLine: /<w:spacing\b[^>]*w:line="360"/.test(gamma),
            deltaBefore: /<w:spacing\b[^>]*w:before="240"/.test(delta),
            deltaAfter: /<w:spacing\b[^>]*w:after="120"/.test(delta),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({
        colour: true,
        highlight: true,
        betaBold: false,
        betaCentred: true,
        gammaJustified: true,
        gammaLine: true,
        deltaBefore: true,
        deltaAfter: true,
      })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' +
  ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
    .join('') +
  '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4675"/><w:gridCol w:w="4675"/></w:tblGrid>' +
  [
    ['north', 'south'],
    ['east', 'west'],
  ]
    .map(
      (row) =>
        `<w:tr>${row
          .map(
            (text) => `<w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>${para(text)}</w:tc>`,
          )
          .join('')}</w:tr>`,
    )
    .join('') +
  '</w:tbl>'

test('a multilevel list, a new style, merged cells and an equation', async () => {
  test.setTimeout(180_000)
  const { dir, docPath, launched, page } = await open(
    ['one', 'two', 'three'].map(para).join('') + TABLE + para('last'),
    'structure.docx',
  )
  try {
    const { app } = launched
    // a list item is not drawn as a p
    const paragraph = (text: string) =>
      page.locator('.doc-page :is(p, .doc-li)', { hasText: text }).first()
    await expect(paragraph('one')).toBeVisible({ timeout: 30_000 })
    await ribbonTab(page, 'Home')

    // Home > Multilevel List: two items, the second pushed down a level with Tab
    await selectParagraph(page, 'one', 'two')
    await ribbon(page, 'Multilevel list').click()
    await page.locator('.list-gallery-multi .list-gallery-card').nth(1).click()
    // the caret at the start of the second item (a click can land on the item above it
    // while the list is still being laid out)
    await selectParagraph(page, 'two')
    await page.keyboard.press('ArrowLeft')
    await page.keyboard.press('Tab')

    // Home > Styles Pane > New Style…: a bold paragraph style for the paragraph with the caret
    await paragraph('three').click()
    await ribbon(page, 'Styles Pane').click()
    await page.locator('.styles-pane-new', { hasText: 'New Style' }).click()
    const style = page.locator('.style-dialog')
    await style.getByLabel('Name').fill('Callout')
    await style.getByLabel('Bold').check()
    await page.screenshot({ path: test.info().outputPath('new-style.png') })
    await style.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(style).toBeHidden()
    // the new style is on the paragraph already, as in Word
    await expect(page.locator('.styles-pane-current')).toContainText('Callout')

    // Table Layout > Merge Cells on the first row
    // whole cells select with a click and a Shift+click; a drag is a text selection
    await page.locator('.styles-pane .nav-pane-close').click()
    // the table's tabs are there only while the caret is in the table
    await page.locator('.doc-page td', { hasText: 'north' }).click()
    await ribbonTab(page, 'Table Layout')
    await expect(async () => {
      await page.locator('.doc-page td', { hasText: 'north' }).click()
      await page.locator('.doc-page td', { hasText: 'south' }).click({ modifiers: ['Shift'] })
      await expect(ribbon(page, 'Merge Cells')).toBeEnabled({ timeout: 2_000 })
    }).toPass({ timeout: 20_000 })
    await ribbon(page, 'Merge Cells').click()
    await expect(page.locator('.doc-page td[colspan="2"]')).toHaveCount(1)

    // Insert > Equation
    await paragraph('last').click()
    await page.keyboard.press('End')
    await ribbonTab(page, 'Insert')
    await ribbon(page, 'Insert an equation').click()
    await page.screenshot({ path: test.info().outputPath('equation.png') })
    await page.locator('[data-rb-panel] button', { hasText: 'Pythagorean Theorem' }).click()
    await expect(page.locator('.doc-page math').first()).toBeVisible()

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const [one, two, three] = await Promise.all(
            ['one', 'two', 'three'].map((text) => savedParagraph(docPath, text)),
          )
          const xml = await part(docPath, 'word/document.xml')
          const styles = await part(docPath, 'word/styles.xml')
          const styleId =
            /<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>\s*<w:name w:val="Callout"\/>/.exec(
              styles,
            )?.[1]
          return {
            oneLevel: /<w:ilvl w:val="(\d)"\/>/.exec(one)?.[1],
            twoLevel: /<w:ilvl w:val="(\d)"\/>/.exec(two)?.[1],
            sameList:
              /<w:numId w:val="(\d+)"\/>/.exec(one)?.[1] ===
              /<w:numId w:val="(\d+)"\/>/.exec(two)?.[1],
            styled: !!styleId && three.includes(`<w:pStyle w:val="${styleId}"/>`),
            merged: /<w:gridSpan w:val="2"\/>/.test(xml),
            equation: /<m:oMath[ >]/.test(xml),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({
        oneLevel: '0',
        twoLevel: '1',
        sameList: true,
        styled: true,
        merged: true,
        equation: true,
      })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})

test('a comment is added, replied to and resolved from the Review tab', async () => {
  test.setTimeout(180_000)
  const { dir, docPath, launched, page } = await open(
    ['alpha', 'beta'].map(para).join(''),
    'review.docx',
  )
  try {
    const { app } = launched
    const beta = page.locator('.doc-page p', { hasText: 'beta' }).first()
    await expect(beta).toBeVisible({ timeout: 30_000 })
    await selectParagraph(page, 'beta')
    await ribbonTab(page, 'Review')
    await ribbon(page, 'New Comment').click()
    const pane = page.locator('.comments-pane')
    await pane.locator('.comment-compose textarea').fill('check this figure')
    await pane.locator('.comment-compose-actions button.primary').click()
    const thread = pane.locator('.comment-thread').first()
    await expect(thread).toContainText('check this figure')
    await thread.locator('.comment-reply-input').fill('checked, it is right')
    await thread.locator('.comment-reply-input').press('Enter')
    await expect(thread).toContainText('checked, it is right')
    await page.screenshot({ path: test.info().outputPath('comment-thread.png') })
    await thread.locator('.comment-resolve').click()

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const comments = await part(docPath, 'word/comments.xml')
          const extended = await part(docPath, 'word/commentsExtended.xml')
          return {
            comment: comments.includes('check this figure'),
            reply: comments.includes('checked, it is right'),
            anchored: /<w:commentRangeStart\b/.test(await part(docPath, 'word/document.xml')),
            replyLinked: /w15:paraIdParent="/.test(extended),
            resolved: /w15:done="1"/.test(extended),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ comment: true, reply: true, anchored: true, replyLinked: true, resolved: true })
  } finally {
    await closeAndSaveVideo(launched, VIDEO_DIR)
    rmSync(dir, { recursive: true, force: true })
  }
})
