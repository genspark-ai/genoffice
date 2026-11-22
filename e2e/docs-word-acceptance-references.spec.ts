/** Word acceptance tasks, references: each walked along Word's path and read back from the saved file. */
import { test, expect } from '@playwright/test'
import { para, pictureRun, saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  clickMenu,
  open,
  part,
  ribbon,
  ribbonTab,
  savedBody,
  savedParagraph,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance-references'

test('References > Insert Footnote adds a footnote with its text', async () => {
  test.setTimeout(180_000)
  const s = await open(VIDEO_DIR, ['a claim', 'more text'].map(para).join(''), 'footnote.docx')
  try {
    const { app, page, docPath } = s
    await expect(page.locator('.doc-page p', { hasText: 'a claim' })).toBeVisible({
      timeout: 30_000,
    })
    await caretAtEnd(page, 'a claim')
    await ribbonTab(page, 'References')
    await ribbon(page, 'Insert Footnote').click()
    // the text is asked for in a box; Word puts the caret in the note at the foot of the page
    const prompt = page.locator('.modal', { hasText: 'Insert Footnote' })
    await prompt.locator('textarea').fill('The source of the claim.')
    await prompt.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(prompt).toBeHidden()
    await expect(page.locator('.page-wrap', { hasText: 'The source of the claim.' })).toBeVisible()
    await page.screenshot({ path: test.info().outputPath('footnote.png') })

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const id = /<w:footnoteReference\b[^>]*w:id="(\d+)"/.exec(
            await savedParagraph(docPath, 'a claim'),
          )?.[1]
          const note =
            new RegExp(`<w:footnote\\b[^>]*w:id="${id}"[\\s\\S]*?</w:footnote>`).exec(
              await part(docPath, 'word/footnotes.xml'),
            )?.[0] ?? ''
          return note.includes('The source of the claim.')
        },
        { timeout: 30_000 },
      )
      .toBe(true)
  } finally {
    await s.close()
  }
})

test('Insert > Caption numbers a figure and Insert > Cross-reference points at it', async () => {
  test.setTimeout(180_000)
  const s = await open(
    VIDEO_DIR,
    para('The chart follows.') +
      `<w:p>${pictureRun(1, 'left', 0, 120, 80)}</w:p>` +
      para('As shown in'),
    'caption.docx',
  )
  try {
    const { app, page, docPath } = s
    const picture = page.locator('.doc-page img').first()
    await expect(picture).toBeVisible({ timeout: 30_000 })

    await picture.click()
    await clickMenu(app, ['Insert', 'Caption…'])
    const caption = page.locator('.modal', { hasText: 'Insert Caption' })
    await expect(caption.locator('select[name="caption-label"]')).toHaveValue('Figure')
    await caption.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(caption).toBeHidden()
    await expect(
      page.locator('.doc-page .doc-protected-field', { hasText: 'Figure 1' }),
    ).toBeVisible()

    await caretAtEnd(page, 'As shown in')
    await page.keyboard.type(' ')
    await clickMenu(app, ['Insert', 'Cross-reference…'])
    const crossRef = page.locator('.modal', { hasText: 'Cross-reference' })
    await crossRef.locator('select[name="ref-type"]').selectOption({ label: 'Figure' })
    await crossRef.locator('.bookmark-row', { hasText: 'Figure 1' }).click()
    await page.screenshot({ path: test.info().outputPath('cross-reference.png') })
    await crossRef.locator('button[name="insert"]').click()
    await expect(page.locator('.doc-page .doc-ref-field')).toHaveText(['Figure 1'])
    if (await crossRef.isVisible()) await page.keyboard.press('Escape')

    // a caption put in above renumbers the one below it, and the reference follows
    await caretAtEnd(page, 'The chart follows.')
    await clickMenu(app, ['Insert', 'Caption…'])
    await caption.locator('input[name="caption-text"]').fill('Earlier')
    await caption.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(caption).toBeHidden()
    await expect(page.locator('.doc-page .doc-protected-field-text')).toHaveText([
      'Figure 1 Earlier',
      'Figure 2',
    ])
    await expect(page.locator('.doc-page .doc-ref-field')).toHaveText(['Figure 2'])

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const body = await savedBody(docPath)
          const target = /REF (_Ref\d+)/.exec(await savedParagraph(docPath, 'As shown in '))?.[1]
          // the picture's caption: the second of the two by now
          const captionParagraph =
            (body.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).filter((p) => /SEQ Figure/.test(p))[1] ??
            ''
          return {
            sequence: /SEQ Figure \\\* ARABIC/.test(captionParagraph),
            referenceToCaption:
              !!target &&
              new RegExp(`<w:bookmarkStart\\b[^>]*w:name="${target}"`).test(captionParagraph),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ sequence: true, referenceToCaption: true })
    // the numbers are saved in document order as each field's result, with nothing left for Word to update
    const body = await savedBody(docPath)
    expect(
      [...body.matchAll(/SEQ Figure[^<]*<\/w:instrText>[\s\S]*?<w:t[^>]*>([^<]*)<\/w:t>/g)].map(
        (m) => m[1],
      ),
    ).toEqual(['1', '2'])
    expect(body).toMatch(/ REF _Ref\d+ [^<]*<\/w:instrText>[\s\S]*?<w:t[^>]*>Figure 2<\/w:t>/)
    expect(body).not.toContain('w:dirty')
  } finally {
    await s.close()
  }
})
