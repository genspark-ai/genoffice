/** Word acceptance tasks, review: each walked along Word's path and read back from the saved file. */
import { test, expect, type ElectronApplication, type Page } from '@playwright/test'
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { para, saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  clickMenu,
  open,
  pickFile,
  pickSavePath,
  ribbon,
  ribbonTab,
  savedBody,
  savedParagraph,
  savedText,
  selectWord,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance-review'

/** Review > Accept or Reject: the change at the caret, then on to the next one */
const review = async (page: Page, action: 'Accept' | 'Reject') => {
  await ribbon(page, action).click()
  await page.locator('[data-rb-panel] button', { hasText: `${action} and Move to Next` }).click()
}

test('Review > Track Changes records an insertion and a deletion; Accept and Reject settle them one at a time', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    ['The quick fox', 'It costs thirty dollars', 'closing'].map(para).join(''),
    'tracked.docx',
  )
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'closing' })).toBeVisible({
      timeout: 30_000,
    })
    const inserted = page.locator('.doc-page .doc-ins')
    const deleted = page.locator('.doc-page .doc-del')

    await ribbonTab(page, 'Review')
    await ribbon(page, 'Track Changes').click()
    await expect(page.locator('.status-track')).toHaveText('Track Changes: On')
    await caretAtEnd(page, 'The quick fox')
    await page.keyboard.type(' jumps')
    await selectWord(page, 'thirty ')
    await page.keyboard.press('Backspace')
    // typing that runs over a second boundary is recorded as two insertions side by side
    await expect.poll(async () => (await inserted.allTextContents()).join('')).toBe(' jumps')
    await expect(deleted).toHaveText('thirty')
    await page.screenshot({ path: test.info().outputPath('tracked.png') })
    await saveFromMenu(app)
    await expect
      .poll(
        async () => [
          // across the insertions, as on the page: typed over a second boundary there are two
          (
            (await savedParagraph(docPath, 'The quick fox')).match(/<w:ins\b[\s\S]*?<\/w:ins>/g) ??
            []
          )
            .join('')
            .replace(/<[^>]+>/g, '')
            .includes('jumps'),
          /<w:del\b[^>]*>(?:(?!<\/w:del>)[\s\S])*<w:delText[^>]*>thirty /.test(
            await savedBody(docPath),
          ),
        ],
        { timeout: 30_000 },
      )
      .toEqual([true, true])

    await caretAtEnd(page, 'The quick fox')
    await review(page, 'Accept')
    await expect(inserted).toHaveCount(0)
    await expect(deleted).toHaveCount(1)
    await selectWord(page, 'thirty')
    await page.keyboard.press('ArrowLeft')
    await review(page, 'Reject')
    await expect(deleted).toHaveCount(0)
    await expect(page.locator('.doc-page p', { hasText: 'It costs' })).toHaveText(
      'It costs thirty dollars',
    )

    await saveFromMenu(app)
    await expect
      .poll(
        async () => ({
          marks: /<w:(ins|del)\b/.test(await savedBody(docPath)),
          text: await savedText(docPath),
        }),
        { timeout: 30_000 },
      )
      .toEqual({
        marks: false,
        text: ['The quick fox jumps', 'It costs thirty dollars', 'closing'],
      })
  } finally {
    await s.close()
  }
})

/** the tab a comparison opened: the docs page, other than `from`, that shows tracked changes */
async function comparedPage(app: ElectronApplication, from: Page): Promise<Page> {
  let found: Page | undefined
  await expect
    .poll(
      async () => {
        for (const candidate of app.windows()) {
          if (candidate === from) continue
          const href = await candidate.evaluate(() => window.location.href).catch(() => '')
          if (!href.includes('://docs/')) continue
          const marks = await candidate
            .locator('.doc-page .doc-ins')
            .count()
            .catch(() => 0)
          if (marks > 0) found = candidate
        }
        return !!found
      },
      { timeout: 60_000 },
    )
    .toBe(true)
  return found!
}

const ORIGINAL = [
  'The lease runs for one year.',
  'Rent is thirty dollars.',
  'Signed by both parties.',
]
const REVISED = [
  'The lease runs for one year.',
  'Rent is sixty dollars.',
  'Signed by both parties.',
  'A new closing clause.',
]

test('Review > Compare makes a document that carries the differences as tracked changes', async () => {
  test.setTimeout(180_000)
  const s = await open(VIDEO_DIR, ORIGINAL.map(para).join(''), 'original.docx', {
    'revised.docx': REVISED.map(para).join(''),
  })
  try {
    const { app, page, dir } = s
    await expect(page.locator('.doc-page p', { hasText: 'Signed' })).toBeVisible({
      timeout: 30_000,
    })
    await pickFile(app, join(dir, 'revised.docx'))
    await ribbonTab(page, 'Review')
    await ribbon(page, 'Compare').click()
    await page.locator('[data-rb-panel] button', { hasText: 'Compare Documents' }).click()
    const dialog = page.locator('.cmp-dialog')
    await expect(dialog.locator('select[name="compare-original"]')).toHaveValue('current')
    await dialog.locator('button[name="browse-revised"]').click()
    await expect(dialog.locator('select[name="compare-revised"] option:checked')).toHaveText(
      'revised.docx',
    )
    await page.screenshot({ path: test.info().outputPath('compare-dialog.png') })
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()

    // the comparison opens as a document of its own; the one it was run from is untouched
    const result = await comparedPage(app, page)
    await expect(result.locator('.doc-page .doc-ins', { hasText: 'sixty' })).toHaveCount(1)
    await expect(result.locator('.doc-page .doc-del', { hasText: 'thirty' })).toHaveCount(1)
    await expect(result.locator('.doc-page .doc-ins', { hasText: 'closing clause' })).toHaveCount(1)
    await expect(page.locator('.doc-page .doc-ins, .doc-page .doc-del')).toHaveCount(0)
    await result.screenshot({ path: test.info().outputPath('compare.png') })

    const comparedPath = join(dir, 'compared.docx')
    await pickSavePath(app, comparedPath)
    await clickMenu(app, ['File', 'Save As…'])
    await expect.poll(() => existsSync(comparedPath), { timeout: 30_000 }).toBe(true)
    const body = await savedBody(comparedPath)
    expect(body).toMatch(/<w:ins\b[^>]*>(?:(?!<\/w:ins>).)*sixty/)
    expect(body).toMatch(/<w:del\b[^>]*>(?:(?!<\/w:del>).)*thirty/)
  } finally {
    await s.close()
  }
})

test('Tools > Spelling and Grammar corrects a misspelled word', async () => {
  const s = await open(VIDEO_DIR, para('We will recieve the parcel.'), 'spelling.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'recieve' })).toBeVisible({
      timeout: 30_000,
    })
    // the dictionary loads after the window does
    await expect
      .poll(
        () =>
          page.evaluate(() =>
            (
              window as unknown as { desktop: { spellIsMisspelled(word: string): boolean } }
            ).desktop.spellIsMisspelled('recieve'),
          ),
        { timeout: 30_000 },
      )
      .toBe(true)
    await clickMenu(app, ['Tools', 'Spelling and Grammar', 'Spelling and Grammar…'])
    const spelling = page.locator('.spelling-dialog')
    await expect(spelling.locator('.spelling-context b')).toHaveText('recieve')
    await spelling.getByRole('option', { name: 'receive', exact: true }).click()
    await page.screenshot({ path: test.info().outputPath('spelling.png') })
    await spelling.getByRole('button', { name: 'Change', exact: true }).click()
    await expect(spelling.locator('.spelling-done')).toBeVisible()
    await spelling.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(spelling).toBeHidden()
    await saveFromMenu(app)
    await expect
      .poll(() => savedBody(docPath), { timeout: 30_000 })
      .toContain('We will receive the parcel.')
  } finally {
    await s.close()
  }
})

test('Tools > Word Count counts the words the status bar shows', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    ['one two three', 'four five', 'six seven eight nine'].map(para).join(''),
    'count.docx',
  )
  try {
    const { app, page } = s
    await expect(page.locator('.doc-page p', { hasText: 'six' })).toBeVisible({ timeout: 30_000 })
    await expect(page.locator('.status-wordcount')).toHaveText('9 words')
    await clickMenu(app, ['Tools', 'Word Count…'])
    const count = page.locator('.word-count-dialog')
    await expect(count.locator('tr[data-stat="words"] td').last()).toHaveText('9')
    await expect(count.locator('tr[data-stat="paragraphs"] td').last()).toHaveText('3')
    await page.screenshot({ path: test.info().outputPath('word-count.png') })
    await count.getByRole('button', { name: 'Close', exact: true }).click()
    await expect(count).toBeHidden()

    // the count follows the text
    const middle = page.locator('.doc-page p', { hasText: 'four five' })
    await middle.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' ten eleven')
    await expect(middle).toHaveText('four five ten eleven')
    await expect(page.locator('.status-wordcount')).toHaveText('11 words')
    await clickMenu(app, ['Tools', 'Word Count…'])
    await expect(count.locator('tr[data-stat="words"] td').last()).toHaveText('11')
  } finally {
    await s.close()
  }
})
