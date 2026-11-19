import { test, expect } from '@playwright/test'
import JSZip from 'jszip'
import { withCommentEditor } from './docs-comment-fixture'

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

async function commentDocument(lines: number, replyLines = 0, wrapped = false): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
  )
  const root = wrapped
    ? `<w:t>${'Wrapped review words '.repeat(lines)}</w:t>`
    : Array.from(
        { length: lines },
        (_, i) => `${i ? '<w:br/>' : ''}<w:t>Review line ${i + 1}</w:t>`,
      ).join('')
  const reply = replyLines
    ? `<w:comment w:id="1" w:author="Reply reviewer"><w:p w14:paraId="00000002"><w:r>${Array.from({ length: replyLines }, (_, i) => `${i ? '<w:br/>' : ''}<w:t>Reply line ${i + 1}</w:t>`).join('')}</w:r></w:p></w:comment>`
    : ''
  zip.file(
    'word/comments.xml',
    `<w:comments ${W} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:comment w:id="0" w:author="Reviewer"><w:p w14:paraId="00000001"><w:r>${root}</w:r></w:p></w:comment>${reply}</w:comments>`,
  )
  if (replyLines)
    zip.file(
      'word/commentsExtended.xml',
      '<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:commentEx w15:paraId="00000001"/><w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000001"/></w15:commentsEx>',
    )
  zip.file(
    'word/document.xml',
    `<w:document ${W}><w:body><w:p><w:r><w:t>Before </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>commented text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

test('margin comments show a five-line summary only beyond fifteen visual lines and expand without losing text', async () => {
  test.setTimeout(120_000)
  for (const lines of [2, 15, 16, 100]) {
    await withCommentEditor(await commentDocument(lines), async (page) => {
      const card = page.locator('.comment-bubble[data-comment-id="0"]')
      await expect(card).toBeVisible()
      const text = card.locator('.comment-entry .comment-text')
      const visibleLines = () =>
        text.evaluate(
          (el) => el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
        )
      await expect.poll(visibleLines).toBeCloseTo(lines > 15 ? 5 : lines, 0)
      await expect(text).toContainText(`Review line ${lines}`)
      if (lines > 15) {
        await card.getByRole('button', { name: 'Expand', exact: true }).click()
        await expect(card.getByRole('button', { name: 'Collapse', exact: true })).toBeVisible()
        if (lines === 100) {
          await expect
            .poll(() => card.evaluate((el) => el.scrollHeight > el.clientHeight))
            .toBe(true)
          await card.evaluate((el) => {
            el.scrollTop = el.scrollHeight
          })
          await expect.poll(() => card.evaluate((el) => el.scrollTop)).toBeGreaterThan(0)
        } else {
          await expect.poll(visibleLines).toBeCloseTo(lines, 0)
        }
        await card.getByRole('button', { name: 'Collapse', exact: true }).click()
        await expect.poll(visibleLines).toBeCloseTo(5, 0)
      } else {
        await expect(card.getByRole('button', { name: 'Expand', exact: true })).toHaveCount(0)
      }
    })
  }
})

test('wrapped prose uses visual lines and replies follow the main comment preview', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(await commentDocument(100, 0, true), async (page) => {
    const card = page.locator('.comment-bubble[data-comment-id="0"]')
    const text = card.locator('.comment-entry .comment-text')
    await expect(card).toBeVisible()
    await expect
      .poll(() =>
        text.evaluate(
          (el) => el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
        ),
      )
      .toBeCloseTo(5, 0)
    expect(await text.textContent()).not.toContain('\n')
    await card.getByRole('button', { name: 'Expand', exact: true }).click()
    await expect
      .poll(() =>
        text.evaluate(
          (el) => el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
        ),
      )
      .toBeGreaterThan(15)
  })
  await withCommentEditor(await commentDocument(16, 8), async (page) => {
    const card = page.locator('.comment-bubble[data-comment-id="0"]')
    await expect(card.locator('.comment-reply')).toHaveCount(1)
    for (const selector of ['.comment-entry .comment-text', '.comment-reply .comment-text']) {
      const text = card.locator(selector)
      await expect
        .poll(() =>
          text.evaluate(
            (el) => el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
          ),
        )
        .toBeCloseTo(5, 0)
    }
    await expect(card.locator('.comment-reply .comment-text')).toContainText('Reply line 8')
    await card.getByRole('button', { name: 'Expand', exact: true }).click()
    await expect
      .poll(() =>
        card
          .locator('.comment-reply .comment-text')
          .evaluate(
            (el) => el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
          ),
      )
      .toBeCloseTo(8, 0)
  })
})
