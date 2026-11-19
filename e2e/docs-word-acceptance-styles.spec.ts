/** Word acceptance tasks, styles and structure: each walked along Word's path and read back from the saved file. */
import { test, expect } from '@playwright/test'
import { para, saveFromMenu } from './docs-objects-fixture'
import {
  caretAtEnd,
  open,
  pickStyle,
  ribbon,
  ribbonTab,
  savedBody,
  selectParagraph,
} from './docs-word-acceptance-kit'

const VIDEO_DIR = 'docs-word-acceptance-styles'

test('Home > Styles applies Heading 1 and Heading 2, References > Table of Contents lists them and Update Table follows a changed heading, View > Navigation Pane jumps to a heading', async () => {
  test.setTimeout(180_000)
  const filler = Array.from({ length: 70 }, (_, i) => para(`filler line ${i + 1}`)).join('')
  const s = await open(
    VIDEO_DIR,
    para('Contents go here') +
      para('Chapter one') +
      para('body text') +
      filler +
      para('Section two') +
      para('closing text'),
    'structure.docx',
  )
  try {
    const { app, page, docPath } = s
    const block = (text: string) =>
      page.locator('.doc-page :is(p, h1, h2)', { hasText: text }).first()
    await expect(block('Chapter one')).toBeVisible({ timeout: 30_000 })
    await ribbonTab(page, 'Home')

    await selectParagraph(page, 'Chapter one')
    await pickStyle(page, 'Heading 1')
    await selectParagraph(page, 'Section two')
    await pickStyle(page, 'Heading 2')
    await expect(page.locator('.doc-page h1')).toHaveText('Chapter one')
    await expect(page.locator('.doc-page h2')).toHaveText('Section two')

    await caretAtEnd(page, 'Contents go here')
    await ribbonTab(page, 'References')
    await ribbon(page, 'Table of Contents').click()
    const entries = page.locator('.doc-page .doc-toc-title')
    await expect(entries).toHaveText(['Chapter one', 'Section two'])

    await caretAtEnd(page, 'Section two')
    await page.keyboard.type(' revised')
    await expect(entries).toHaveText(['Chapter one', 'Section two'])
    await ribbon(page, 'Update Table').click()
    await expect(entries).toHaveText(['Chapter one', 'Section two revised'])
    await block('Contents go here').scrollIntoViewIfNeeded()
    await page.screenshot({ path: test.info().outputPath('contents.png') })

    await ribbonTab(page, 'View')
    await ribbon(page, 'Navigation Pane').click()
    const rows = page.locator('.nav-pane .nav-row')
    await expect(rows).toHaveText(['Chapter one', 'Section two revised'])
    await expect(page.locator('.doc-page h2')).not.toBeInViewport()
    await rows.nth(1).click()
    await expect(page.locator('.doc-page h2')).toBeInViewport()
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as never as { __aidocs: { editor: any } }).__aidocs.editor.state.selection.$from
              .parent.textContent,
        ),
      )
      .toBe('Section two revised')

    await saveFromMenu(app)
    await expect
      .poll(
        async () => {
          const body = await savedBody(docPath)
          const field =
            /<w:fldChar w:fldCharType="begin"[\s\S]*?<w:fldChar w:fldCharType="end"\/>/.exec(
              body,
            )?.[0] ?? ''
          // the table of contents repeats the heading text, so the style decides which paragraph is meant
          const styled = (style: string, text: string) =>
            (body.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? []).some(
              (p) => p.includes(`<w:pStyle w:val="${style}"/>`) && p.includes(`>${text}<`),
            )
          return {
            heading1: styled('Heading1', 'Chapter one'),
            heading2: styled('Heading2', 'Section two revised'),
            tocField: /<w:instrText[^>]*>\s*TOC\b/.test(field),
            tocEntries: ['Chapter one', 'Section two revised'].map((text) => field.includes(text)),
          }
        },
        { timeout: 30_000 },
      )
      .toEqual({ heading1: true, heading2: true, tocField: true, tocEntries: [true, true] })
  } finally {
    await s.close()
  }
})
