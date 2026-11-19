import { test, expect } from '@playwright/test'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { withCommentEditor } from './docs-comment-fixture'

test('a popup connector remains visible through the card edge without reserving a column', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(await mixedPages(), async (page) => {
    await expect(page.locator('.comment-bubble')).toHaveCount(3)
    await page.locator('.ribbon-mode-split > button[aria-pressed]').first().click()
    await expect(page.locator('.comment-bubble')).toHaveCount(0)
    await page.locator('.comment-balloon[data-comment-id="1"]').click()
    await expect(page.locator('.comment-popup[data-comment-id="1"]')).toBeVisible()
    await expect(page.locator('.page-wrap')).not.toHaveClass(/has-markup-area/)
    const bounds = await page.locator('.comment-leader[data-comment-id="1"]').evaluate((el) => {
      const path = el as SVGPathElement
      const end = path.getPointAtLength(path.getTotalLength()).matrixTransform(path.getScreenCTM()!)
      const id = /#([^)]*)/.exec(path.getAttribute('clip-path')!)![1]
      const clip = document.getElementById(id)!.querySelector('rect')!
      const x = Number(clip.getAttribute('x'))
      const width = Number(clip.getAttribute('width'))
      const left = new DOMPoint(x, 0).matrixTransform(path.getScreenCTM()!).x
      const right = new DOMPoint(x + width, 0).matrixTransform(path.getScreenCTM()!).x
      return { end: end.x, left, right }
    })
    expect(bounds.end).toBeGreaterThanOrEqual(bounds.left)
    expect(bounds.end).toBeLessThanOrEqual(bounds.right)
  })
})

const comments =
  '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  ['1', '2', '3']
    .map(
      (id) =>
        `<w:comment w:id="${id}" w:author="Reviewer"><w:p><w:r><w:t>Review ${id}</w:t></w:r></w:p></w:comment>`,
    )
    .join('') +
  '</w:comments>'

async function mixedPages(): Promise<Buffer> {
  return Buffer.from(
    await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>Before </w:t></w:r><w:commentRangeStart w:id="1"/><w:r><w:t>narrow </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>target</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>' +
        Array.from(
          { length: 18 },
          (_, i) =>
            `<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>Calibration line ${i}: regular text before the endpoint.</w:t></w:r></w:p>`,
        ).join('') +
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:pPr><w:commentRangeStart w:id="3"/><w:r><w:t>Cross range begins</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>Cross range ends</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/><w:t xml:space="preserve"> untouched suffix</w:t></w:r></w:p>' +
        '<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>wide target</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>',
      sectPrXml:
        '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>',
      extraRels:
        '<Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>',
      extraParts: [
        {
          path: 'word/comments.xml',
          xml: comments,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
        },
      ],
    }),
  )
}

test('comments follow real range endpoints and the width of their endpoint page', async () => {
  test.setTimeout(120_000)
  await withCommentEditor(await mixedPages(), async (page) => {
    await expect(page.locator('.comment-bubble')).toHaveCount(3)
    await expect(page.locator('.page-sheet')).toHaveCount(2)
    const geometry = () =>
      page.evaluate(() => {
        const rect = (el: Element) => {
          const r = el.getBoundingClientRect()
          return { left: r.left, right: r.right, top: r.top, bottom: r.bottom }
        }
        const cards = Object.fromEntries(
          [...document.querySelectorAll<HTMLElement>('.comment-bubble')].map((el) => [
            el.dataset.commentId!,
            rect(el),
          ]),
        )
        const sheets = [...document.querySelectorAll('.page-sheet')].map(rect)
        const span = [...document.querySelectorAll('.doc-comment')]
          .filter((el) => el.getAttribute('data-comment-ids') === '1')
          .at(-1)!
        const target = [...span.getClientRects()].at(-1)!
        const path = document.querySelector<SVGPathElement>('.comment-leader')!
        const point = path.getPointAtLength(0).matrixTransform(path.getScreenCTM()!)
        const paragraph = [...document.querySelectorAll('.ProseMirror p')].find((p) =>
          p.textContent?.includes('Cross range ends'),
        )!
        const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
        let text: Node | null
        let end: DOMRect | undefined
        while ((text = walker.nextNode())) {
          const index = text.textContent!.indexOf('Cross range ends')
          if (index < 0) continue
          const range = document.createRange()
          range.setStart(text, index)
          range.setEnd(text, index + 'Cross range ends'.length)
          end = [...range.getClientRects()].at(-1)
          break
        }
        const crossPath = document.querySelectorAll<SVGPathElement>('.comment-leader')[1]
        const crossPoint = crossPath.getPointAtLength(0).matrixTransform(crossPath.getScreenCTM()!)
        return {
          cards,
          sheets,
          target: { right: target.right, bottom: target.bottom },
          point: { x: point.x, y: point.y },
          cross: { x: crossPoint.x, y: crossPoint.y, right: end!.right, bottom: end!.bottom },
        }
      })
    await expect
      .poll(async () => {
        const g = await geometry()
        return Math.abs(g.cards['1'].left - g.sheets[0].right - 14)
      })
      .toBeLessThan(1)
    const g = await geometry()
    expect(Math.abs(g.cards['2'].left - g.sheets[1].right - 14)).toBeLessThan(1)
    expect(g.cards['3'].top).toBeGreaterThan(g.sheets[1].top)
    expect(g.cards['3'].bottom).toBeLessThan(g.sheets[1].bottom)
    expect(Math.abs(g.point.x - g.target.right)).toBeLessThan(1)
    expect(Math.abs(g.point.y - (g.target.bottom - 1))).toBeLessThan(1)
    expect(Math.abs(g.cross.x - g.cross.right)).toBeLessThan(1)
    expect(Math.abs(g.cross.y - (g.cross.bottom - 1))).toBeLessThan(1)
    await page.locator('.comment-bubble[data-comment-id="1"]').click()
    await expect
      .poll(async () => {
        const next = await geometry()
        return Math.abs(next.cards['1'].left - next.sheets[0].right - 14)
      })
      .toBeLessThan(1)
    for (const where of ['before', 'after']) {
      await page.evaluate((where) => {
        const paragraph = [...document.querySelectorAll('.ProseMirror p')].find((p) =>
          p.textContent?.includes('Cross range ends'),
        )!
        const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT)
        const texts: Node[] = []
        let node: Node | null
        while ((node = walker.nextNode())) if (node.textContent) texts.push(node)
        const target = where === 'before' ? texts[0] : texts.at(-1)!
        ;(paragraph.closest('.ProseMirror') as HTMLElement).focus()
        const selection = window.getSelection()!
        selection.removeAllRanges()
        const range = document.createRange()
        range.setStart(target, where === 'before' ? 0 : target.textContent!.length)
        range.collapse(true)
        selection.addRange(range)
        document.dispatchEvent(new Event('selectionchange'))
      }, where)
      await page.keyboard.type(where === 'before' ? 'Prefix ' : ' after')
      await expect
        .poll(async () => {
          const next = await geometry()
          return Math.abs(next.cross.x - next.cross.right)
        })
        .toBeLessThan(1)
    }
    const slider = page.locator('.zoom-slider')
    await slider.press('ArrowLeft')
    await slider.press('ArrowLeft')
    await expect(slider).toHaveValue('80')
    await expect
      .poll(async () => {
        const next = await geometry()
        return Math.max(
          Math.abs(next.cross.x - next.cross.right),
          Math.abs(next.cross.y - (next.cross.bottom - 0.8)),
          Math.abs(next.point.x - next.target.right),
          Math.abs(next.point.y - (next.target.bottom - 0.8)),
        )
      })
      .toBeLessThan(1)
    const zoomed = await geometry()
    for (const [id, index] of [
      ['1', 0],
      ['2', 1],
      ['3', 1],
    ] as const) {
      expect(Math.abs(zoomed.cards[id].left - zoomed.sheets[index].right - 14 * 0.8)).toBeLessThan(
        1,
      )
      expect(zoomed.cards[id].top).toBeGreaterThanOrEqual(zoomed.sheets[index].top - 1)
      expect(zoomed.cards[id].bottom).toBeLessThanOrEqual(zoomed.sheets[index].bottom + 1)
    }
  })
})
