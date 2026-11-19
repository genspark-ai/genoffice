import { test, expect } from '@playwright/test'
import { capacityDocument } from './docs-comment-capacity-fixture'
import { withCommentEditor } from './docs-comment-fixture'

test('page capacity selects every compression tier from real discussion heights', async () => {
  test.setTimeout(180_000)
  for (const [height, lines] of [
    [12000, 10],
    [4500, 5],
    [3600, 3],
    [2700, 1],
  ]) {
    await withCommentEditor(await capacityDocument({ height }), async (page) => {
      const cards = page.locator('.comment-bubble:not(.rev-bubble)')
      await expect(cards).toHaveCount(2)
      for (const card of await cards.all()) {
        await expect
          .poll(() =>
            card
              .locator('.comment-entry .comment-text')
              .evaluate(
                (el) =>
                  el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
              ),
          )
          .toBeCloseTo(lines, 0)
      }
      const bounds = await cards.evaluateAll((elements) =>
        elements.map((el) => {
          const r = el.getBoundingClientRect()
          return { top: r.top, bottom: r.bottom }
        }),
      )
      expect(bounds[0].bottom).toBeLessThanOrEqual(bounds[1].top)
    })
  }
})

test('reply heights drive compression while short roots and adjacent spacious pages retain natural content', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(
    await capacityDocument({ height: 6750, replyLines: 10, rootLines: 2 }),
    async (page) => {
      const cards = page.locator('.comment-bubble:not(.rev-bubble)')
      await expect(cards).toHaveCount(2)
      for (const card of await cards.all()) {
        await expect(card.locator('.comment-reply')).toHaveCount(1)
        for (const [selector, lines] of [
          ['.comment-entry .comment-text', 2],
          ['.comment-reply .comment-text', 5],
        ] as const) {
          await expect
            .poll(() =>
              card
                .locator(selector)
                .evaluate(
                  (el) =>
                    el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
                ),
            )
            .toBeCloseTo(lines, 0)
        }
      }
    },
  )
  await withCommentEditor(
    await capacityDocument({ height: 4500, secondPage: true }),
    async (page) => {
      await expect(page.locator('.comment-page')).toHaveCount(2)
      const first = page.locator('.comment-page').nth(0),
        second = page.locator('.comment-page').nth(1)
      for (const text of await first.locator('.comment-entry .comment-text').all())
        await expect
          .poll(() =>
            text.evaluate(
              (el) =>
                el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
            ),
          )
          .toBeCloseTo(5, 0)
      await expect
        .poll(() =>
          second
            .locator('.comment-entry .comment-text')
            .evaluate(
              (el) =>
                el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
            ),
        )
        .toBeCloseTo(10, 0)
      const geometry = await page.locator('.comment-page').evaluateAll((elements) =>
        elements.map((el) => {
          const r = el.getBoundingClientRect()
          return {
            left: r.left,
            top: r.top,
            bottom: r.bottom,
            cards: [...el.querySelectorAll('.comment-bubble')].map((card) => {
              const c = card.getBoundingClientRect()
              return { top: c.top, bottom: c.bottom }
            }),
          }
        }),
      )
      expect(geometry[1].left).toBeGreaterThan(geometry[0].left)
      for (const region of geometry)
        for (const card of region.cards) {
          expect(card.top).toBeGreaterThanOrEqual(region.top - 1)
          expect(card.bottom).toBeLessThanOrEqual(region.bottom + 1)
        }
    },
  )
})

test('minimum previews scroll only their page and keep the connector tied to visible cards', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(
    await capacityDocument({ height: 4500, count: 8, secondPage: true }),
    async (page) => {
      const region = page.locator('.comment-page').first()
      await expect(region.locator('.comment-bubble')).toHaveCount(8)
      await expect.poll(() => region.evaluate((el) => el.scrollHeight > el.clientHeight)).toBe(true)
      const card = region.locator('.comment-bubble[data-comment-id="7"]')
      const path = page.locator('.comment-leader[data-comment-id="7"]')
      const point = () =>
        path.evaluate((el) => {
          const path = el as SVGPathElement,
            m = path.getScreenCTM()!,
            start = path.getPointAtLength(0),
            end = path.getPointAtLength(path.getTotalLength())
          return {
            start: new DOMPoint(start.x, start.y).matrixTransform(m).toJSON(),
            end: new DOMPoint(end.x, end.y).matrixTransform(m).toJSON(),
          }
        })
      const before = await point()
      await region.evaluate((el) => {
        el.scrollTop = el.scrollHeight
      })
      await expect.poll(() => region.evaluate((el) => el.scrollTop)).toBeGreaterThan(0)
      await expect
        .poll(async () => Math.abs((await point()).end.y - (await card.boundingBox())!.y - 9))
        .toBeLessThan(2)
      const after = await point(),
        anchor = await page.locator('.doc-comment[data-comment-ids="7"]').boundingBox()
      expect(after.start.x).toBeCloseTo(before.start.x, 1)
      expect(after.start.y).toBeCloseTo(before.start.y, 1)
      expect(after.start.x).toBeCloseTo(anchor!.x + anchor!.width, 0)
      expect(after.start.y).toBeGreaterThanOrEqual(anchor!.y)
      expect(after.start.y).toBeLessThanOrEqual(anchor!.y + anchor!.height)
      const cardBounds = await card.boundingBox(),
        regionBounds = await region.boundingBox()
      expect(cardBounds!.y).toBeGreaterThanOrEqual(regionBounds!.y - 1)
      expect(cardBounds!.y + cardBounds!.height).toBeLessThanOrEqual(
        regionBounds!.y + regionBounds!.height + 1,
      )
      await expect
        .poll(() =>
          page
            .locator('.comment-page')
            .nth(1)
            .evaluate((el) => el.scrollTop),
        )
        .toBe(0)
    },
  )
})

test('page-bottom anchor clusters move upward without unnecessarily truncating previews', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(await capacityDocument({ height: 12000, bottom: true }), async (page) => {
    const cards = page.locator('.comment-bubble:not(.rev-bubble)')
    await expect(cards).toHaveCount(2)
    for (const text of await cards.locator('.comment-entry .comment-text').all())
      await expect
        .poll(() =>
          text.evaluate(
            (el) => el.getBoundingClientRect().height / parseFloat(getComputedStyle(el).lineHeight),
          ),
        )
        .toBeCloseTo(10, 0)
    const bounds = await cards.evaluateAll((elements) =>
        elements.map((el) => {
          const r = el.getBoundingClientRect()
          return { top: r.top, bottom: r.bottom }
        }),
      ),
      region = (await page.locator('.comment-page').boundingBox())!,
      anchor = (await page.locator('.doc-comment').first().boundingBox())!
    expect(bounds[0].top).toBeLessThan(anchor.y)
    expect(bounds[0].bottom).toBeLessThanOrEqual(bounds[1].top)
    expect(bounds[1].bottom).toBeLessThanOrEqual(region.y + region.height + 1)
  })
})

test('a wide page followed by a narrow page keeps independent card boundaries', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(
    await capacityDocument({ height: 4500, secondPage: true, widths: [12000, 9000] }),
    async (page) => {
      await expect(page.locator('.page-sheet')).toHaveCount(2)
      await expect(page.locator('.comment-page')).toHaveCount(2)
      const geometry = await page.evaluate(() => {
        const rect = (el: Element) => {
          const r = el.getBoundingClientRect()
          return { right: r.right, left: r.left, top: r.top, bottom: r.bottom }
        }
        return {
          sheets: [...document.querySelectorAll('.page-sheet')].map(rect),
          regions: [...document.querySelectorAll('.comment-page')].map(rect),
          cards: [...document.querySelectorAll('.comment-bubble')].map(rect),
        }
      })
      expect(geometry.regions[1].left).toBeLessThan(geometry.regions[0].left)
      for (let i = 0; i < 2; i++)
        expect(Math.abs(geometry.regions[i].left - geometry.sheets[i].right - 14)).toBeLessThan(1)
      expect(geometry.cards[1].bottom).toBeLessThanOrEqual(geometry.sheets[0].bottom + 1)
      expect(geometry.cards[2].bottom).toBeLessThanOrEqual(geometry.sheets[1].bottom + 1)
    },
  )
})
