import { expect, it } from 'vitest'
import { createBlankPptx, openPptx, savePptx, type TextElement } from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

const offset = { x: 914400, y: 914400, cx: 4000000, cy: 1600000 }
const payload = (spacing: Record<string, number> = {}) => [
  {
    runs: [
      {
        text: 'Spacing example',
        bold: true,
        italic: true,
        fontSize: 18,
        color: '#123456',
        baseline: 12,
      },
    ],
    ...spacing,
  },
]
async function created(paragraphs: unknown[]) {
  const opened = await openPptx(await createBlankPptx())
  const before = new Map(opened.archive.entries)
  const r = runTxn(opened, {
    ops: [{ op: 'addElement', target: { slide: 0 }, kind: 'textbox', offset, paragraphs }],
  })
  expect(r.applied).toBe(true)
  const id = r.records![0]!.created![0]!
  return { opened, id, before }
}
async function reopened(opened: Awaited<ReturnType<typeof openPptx>>, id: string) {
  const again = await openPptx(await savePptx(opened))
  const element = again.deck.slides[0]!.elements.find(
    (item) => item.id === id || item.type === 'text',
  ) as TextElement
  return { again, p: element.text!.paragraphs[0]! }
}
it.each([
  { lineSpacingPct: 132, spaceBeforePt: 7, spaceAfterPt: 8 },
  { lineSpacingPct: 0, spaceBeforePt: 0, spaceAfterPt: 0 },
])('persists public creation spacing %j through save/reopen', async (spacing) => {
  const { opened, id } = await created(payload(spacing))
  const { p } = await reopened(opened, id)
  expect(p.lineHeight).toBe(spacing.lineSpacingPct)
  expect(p.spaceBefore).toBe(spacing.spaceBeforePt)
  expect(p.spaceAfter).toBe(spacing.spaceAfterPt)
})
it('keeps omitted spacing absent and preserves literal run formatting', async () => {
  const paragraphs = payload()
  const snapshot = JSON.stringify(paragraphs)
  const { opened, id } = await created(paragraphs)
  const { p } = await reopened(opened, id)
  expect(p.lineHeight).toBeUndefined()
  expect(p.spaceBefore).toBeUndefined()
  expect(p.spaceAfter).toBeUndefined()
  expect(p.runs[0]).toMatchObject({
    text: 'Spacing example',
    bold: true,
    italic: true,
    fontSize: 18,
    baseline: 12,
  })
  expect(JSON.stringify(paragraphs)).toBe(snapshot)
})
it('keeps setText spacing on the same public payload', async () => {
  const { opened, id } = await created(payload())
  expect(
    runTxn(opened, {
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el: id },
          paragraphs: payload({ lineSpacingPct: 132, spaceBeforePt: 7, spaceAfterPt: 8 }),
        },
      ],
    }).applied,
  ).toBe(true)
  const { p } = await reopened(opened, id)
  expect([p.lineHeight, p.spaceBefore, p.spaceAfter]).toEqual([132, 7, 8])
})

it('keeps table-cell paragraph editing spacing and run formatting', async () => {
  const opened = await openPptx(await createBlankPptx())
  const added = runTxn(opened, {
    ops: [{ op: 'addTable', target: { slide: 0 }, rows: 1, cols: 1, offset }],
  })
  expect(added.applied).toBe(true)
  const id = added.records![0]!.created![0]!
  expect(
    runTxn(opened, {
      ops: [
        {
          op: 'setTableCell',
          target: { slide: 0, el: id },
          row: 0,
          col: 0,
          paragraphs: payload({ lineSpacingPct: 132, spaceBeforePt: 7, spaceAfterPt: 8 }),
        },
      ],
    }).applied,
  ).toBe(true)
  const saved = await openPptx(await savePptx(opened))
  const table = saved.deck.slides[0]!.elements.find((item) => item.type === 'table') as any
  const p = table.rows[0][0].text.paragraphs[0]
  expect([p.lineHeight, p.spaceBefore, p.spaceAfter]).toEqual([132, 7, 8])
  expect(p.runs[0].bold).toBe(true)
})
it('preserves unrelated package payloads when creating a spaced textbox', async () => {
  const { opened, id, before } = await created(
    payload({ lineSpacingPct: 132, spaceBeforePt: 7, spaceAfterPt: 8 }),
  )
  const { again } = await reopened(opened, id)
  expect([...again.archive.entries.keys()].sort()).toEqual([...before.keys()].sort())
  for (const [name, bytes] of before)
    if (name !== 'ppt/slides/slide1.xml')
      expect(again.archive.entries.get(name), name).toEqual(bytes)
})
