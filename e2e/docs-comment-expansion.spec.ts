import { test, expect } from '@playwright/test'
import { capacityDocument } from './docs-comment-capacity-fixture'
import { withCommentEditor } from './docs-comment-fixture'

test('temporary yielding uses the least compressed five- or three-line stage and leaves spacious pages unchanged', async () => {
  test.setTimeout(120_000)
  for (const [height, expandedLines, restoredLines] of [
    [12000, 10, 10],
    [6750, 5, 10],
    [5850, 3, 5],
  ]) {
    await withCommentEditor(await capacityDocument({ height }), async (page) => {
      const first = page.locator('.comment-bubble[data-comment-id="0"]')
      const second = page.locator('.comment-bubble[data-comment-id="1"]')
      const lines = () =>
        second
          .locator('.comment-entry .comment-text')
          .evaluate(
            (el) => el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
          )
      await expect.poll(lines).toBeCloseTo(restoredLines, 0)
      await first.click({ position: { x: 10, y: 10 } })
      await expect(first).toHaveClass(/comment-expanded/)
      await expect.poll(lines).toBeCloseTo(expandedLines, 0)
      await first.getByRole('button', { name: 'Collapse', exact: true }).click()
      await expect.poll(lines).toBeCloseTo(restoredLines, 0)
    })
  }
})

test('expansion borrows page space only as needed and collapse restores normal previews', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(await capacityDocument({ height: 4500 }), async (page) => {
    const region = page.locator('.comment-page')
    const first = region.locator('.comment-bubble[data-comment-id="0"]')
    const second = region.locator('.comment-bubble[data-comment-id="1"]')
    const visibleLines = () =>
      second
        .locator('.comment-entry .comment-text')
        .evaluate(
          (el) => el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
        )
    await expect.poll(visibleLines).toBeCloseTo(5, 0)
    await first.getByRole('button', { name: 'Expand', exact: true }).click()
    await expect(first).toHaveClass(/comment-expanded/)
    await expect.poll(visibleLines).toBeCloseTo(1, 0)
    await expect.poll(() => region.evaluate((el) => el.scrollHeight > el.clientHeight)).toBe(false)
    await expect.poll(() => first.evaluate((el) => el.scrollHeight > el.clientHeight)).toBe(true)
    const bounds = await region.evaluate((el) => {
      const region = el.getBoundingClientRect()
      return [...el.querySelectorAll('.comment-bubble')].map((card) => {
        const r = card.getBoundingClientRect()
        return { top: r.top - region.top, bottom: r.bottom - region.top, limit: region.height }
      })
    })
    expect(bounds[0].bottom).toBeLessThanOrEqual(bounds[1].top)
    expect(bounds[1].bottom).toBeLessThanOrEqual(bounds[1].limit + 1)
    await first.getByRole('button', { name: 'Collapse', exact: true }).click()
    await expect(first).not.toHaveClass(/comment-expanded/)
    await expect.poll(visibleLines).toBeCloseTo(5, 0)
    await expect.poll(() => region.evaluate((el) => el.scrollHeight > el.clientHeight)).toBe(false)
  })
})

test('dense pages keep a readable expanded discussion and preserve both scroll levels and draft through zoom', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(
    await capacityDocument({ height: 4500, count: 8, rootLines: 100, replyLines: 8 }),
    async (page) => {
      const region = page.locator('.comment-page')
      const card = region.locator('.comment-bubble[data-comment-id="7"]')
      await expect(region.locator('.comment-bubble')).toHaveCount(8)
      await region.evaluate((el) => {
        el.scrollTop = el.scrollHeight
      })
      await card.getByRole('button', { name: 'Expand', exact: true }).click()
      await expect(card).toHaveClass(/comment-expanded/)
      await expect.poll(() => region.evaluate((el) => el.scrollHeight > el.clientHeight)).toBe(true)
      await expect.poll(() => card.evaluate((el) => el.scrollHeight > el.clientHeight)).toBe(true)
      const bounds = await card.evaluate((el) => {
        const region = el.closest('.comment-page')!.getBoundingClientRect()
        const card = el.getBoundingClientRect()
        return {
          top: card.top - region.top,
          bottom: card.bottom - region.top,
          limit: region.height,
          height: card.height,
        }
      })
      expect(bounds.height).toBeGreaterThan(60)
      expect(bounds.top).toBeGreaterThanOrEqual(-1)
      expect(bounds.bottom).toBeLessThanOrEqual(bounds.limit + 1)
      const input = card.locator('.comment-reply-input')
      await input.fill('Keep this unsubmitted reply')
      const previous = await card.evaluate((el) => ({
        inner: el.scrollTop,
        outer: el.closest('.comment-page')!.scrollTop,
      }))
      expect(previous.inner).toBeGreaterThan(0)
      expect(previous.outer).toBeGreaterThan(0)
      await page.locator('.zoom-slider').press('ArrowRight')
      await expect(input).toHaveValue('Keep this unsubmitted reply')
      await expect
        .poll(() => card.evaluate((el) => el.scrollTop))
        .toBeGreaterThan(previous.inner - 2)
      await expect.poll(() => region.evaluate((el) => el.scrollTop)).toBeGreaterThan(0)
      await expect(card).toHaveClass(/comment-expanded/)
      await card.getByRole('button', { name: 'Collapse', exact: true }).click()
      await expect(card).not.toHaveClass(/comment-expanded/)
      await expect(input).toHaveValue('Keep this unsubmitted reply')
    },
  )
})

test('inline comment edits retain their selection through repagination and can still be committed', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(await capacityDocument({ height: 6750 }), async (page) => {
    const card = page.locator('.comment-bubble[data-comment-id="0"]')
    await expect(card).toBeVisible()
    await card.click({ position: { x: 10, y: 10 } })
    await card.locator('.comment-entry .comment-more').click()
    await card.locator('.comment-entry .comment-card-edit').click()
    const input = card.locator('.comment-edit-input')
    await input.fill('Unsubmitted correction with its selection')
    await input.evaluate((el) => (el as HTMLTextAreaElement).setSelectionRange(4, 11))
    await page.locator('.zoom-slider').press('ArrowRight')
    await expect(input).toHaveValue('Unsubmitted correction with its selection')
    await expect
      .poll(() =>
        input.evaluate((el) => [
          (el as HTMLTextAreaElement).selectionStart,
          (el as HTMLTextAreaElement).selectionEnd,
        ]),
      )
      .toEqual([4, 11])
    await input.press('Control+Enter')
    await expect(card.locator('.comment-entry .comment-text')).toHaveText(
      'Unsubmitted correction with its selection',
    )
  })
})

for (const kind of ['reply', 'edit'] as const) {
  test(`reflow restores focused ${kind} without scrolling the document to a provisional card position`, async () => {
    test.setTimeout(120_000)
    await withCommentEditor(
      await capacityDocument({ height: 12000, secondPage: true }),
      async (page) => {
        const card = page.locator('.comment-bubble[data-comment-id="2"]')
        await card.click({ position: { x: 10, y: 10 } })
        if (kind === 'edit') {
          await card.locator('.comment-entry .comment-more').click()
          await card.locator('.comment-entry .comment-card-edit').click()
        }
        const input = card.locator(
          kind === 'reply' ? '.comment-reply-input' : '.comment-edit-input',
        )
        await input.fill('Retain focus without moving the document')
        const previous = await page.locator('.editor-scroll').evaluate((el) => el.scrollTop)
        const oldInput = (await input.elementHandle())!
        await page.setViewportSize({ width: 1450, height: 1000 })
        await expect.poll(() => oldInput.evaluate((el) => el.isConnected)).toBe(false)
        await expect(input).toBeFocused()
        await expect(input).toHaveValue('Retain focus without moving the document')
        await expect
          .poll(() => page.locator('.editor-scroll').evaluate((el) => el.scrollTop))
          .toBeCloseTo(previous, 0)
      },
    )
  })
}
