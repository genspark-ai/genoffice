/**
 * per_op isolation: an op that fails after it started changing the deck leaves
 * nothing behind, while the ops around it keep their result.
 */
import { describe, it, expect } from 'vitest'
import {
  addElement,
  createBlankPptx,
  openPptx,
  savePptx,
  type OpenedPptx,
  type TextElement,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import { GuidedError, lookup, register, type Op } from '../src/ops/registry'
import '../src/ops/index'

const MASTER = 'ppt/slideMasters/slideMaster1.xml'
const SLIDE_RELS = 'ppt/slides/_rels/slide1.xml.rels'
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

register({
  name: 'failAfter',
  validate(op, ctx) {
    const inner = op.inner as Op
    lookup(inner.op).validate(inner, ctx)
  },
  apply(op, ctx) {
    const inner = op.inner as Op
    lookup(inner.op).apply(inner, ctx)
    throw new GuidedError('op "failAfter": failed after the change.')
  },
})

const failAfter = (inner: Op): Op => ({ op: 'failAfter', target: inner.target, inner })

async function deckWithText(): Promise<{ opened: OpenedPptx; id: string }> {
  const opened = await openPptx(await createBlankPptx())
  const el = addElement(opened.deck.slides[0]!, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 1828800, cy: 914400 },
    paragraphs: [{ runs: [{ text: 'Original text' }] }],
  })
  return { opened, id: el.id }
}

const linkedText = (id: string, text: string): Op => ({
  op: 'setText',
  target: { slide: 0, el: id },
  paragraphs: [
    {
      runs: [{ text, link: { kind: 'url', url: `https://example.com/${text}` } }],
      bullet: 'blip',
      bulletImage: { base64: PNG, ext: 'png' },
    },
  ],
})

const marker = (text: string): Op => ({
  op: 'addElement',
  target: { slide: 0 },
  kind: 'textbox',
  offset: { x: 0, y: 1828800, cx: 1828800, cy: 914400 },
  paragraphs: [{ runs: [{ text }] }],
})

function slideTexts(opened: OpenedPptx): string[] {
  return opened.deck.slides[0]!.elements.map((el) =>
    ((el as TextElement).text?.paragraphs ?? []).flatMap((p) => p.runs.map((r) => r.text)).join(''),
  )
}

describe('per_op rollback of a failed op', () => {
  it('drops the text, link relationship and media a failed op wrote', async () => {
    const { opened, id } = await deckWithText()
    const relsBefore = opened.archive.readText(SLIDE_RELS)
    const mediaBefore = [...opened.archive.entries.keys()].filter((p) => p.startsWith('ppt/media/'))

    const r = runTxn(opened, {
      isolation: 'per_op',
      ops: [failAfter(linkedText(id, 'rejected')), marker('Successful marker')],
    })

    expect(r.applied).toBe(true)
    expect(r.records).toHaveLength(1)
    expect(r.failures?.map((f) => f.index)).toEqual([0])

    const saved = await openPptx(await savePptx(opened))
    expect(slideTexts(saved)).toEqual(['Original text', 'Successful marker'])
    expect(saved.archive.readText(SLIDE_RELS)).toBe(relsBefore)
    expect(saved.archive.readText(SLIDE_RELS)).not.toContain('example.com')
    expect([...saved.archive.entries.keys()].filter((p) => p.startsWith('ppt/media/'))).toEqual(
      mediaBefore,
    )
  })

  it('keeps the successes on both sides of a failure', async () => {
    const { opened, id } = await deckWithText()
    const r = runTxn(opened, {
      isolation: 'per_op',
      ops: [marker('First'), failAfter(marker('Rejected marker')), linkedText(id, 'accepted')],
    })

    expect(r.records).toHaveLength(2)
    expect(r.failures?.map((f) => f.index)).toEqual([1])

    const saved = await openPptx(await savePptx(opened))
    expect(slideTexts(saved)).toEqual(['accepted', 'First'])
    expect(saved.archive.readText(SLIDE_RELS)).toContain('example.com/accepted')
  })

  it('reports nothing applied when every op fails, and leaves the deck as it was', async () => {
    const { opened, id } = await deckWithText()
    const before = new Map(opened.archive.entries)
    const r = runTxn(opened, {
      isolation: 'per_op',
      ops: [failAfter(linkedText(id, 'rejected')), failAfter(marker('Never'))],
    })

    expect(r.applied).toBe(false)
    expect(r.records).toEqual([])
    expect(r.failures?.map((f) => f.index)).toEqual([0, 1])
    expect(slideTexts(opened)).toEqual(['Original text'])
    expect(new Map(opened.archive.entries)).toEqual(before)
  })

  it('writes a master part with the successful edits only', async () => {
    const opened = await openPptx(await createBlankPptx())
    const box = (text: string): Op => ({
      op: 'addElement',
      target: { part: MASTER },
      kind: 'textbox',
      offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
      paragraphs: [{ runs: [{ text }] }],
    })
    const r = runTxn(opened, {
      isolation: 'per_op',
      ops: [box('Kept box'), failAfter(box('Rejected box')), box('Later box')],
    })

    expect(r.records).toHaveLength(2)
    const xml = opened.archive.readText(MASTER)!
    expect(xml).toContain('Kept box')
    expect(xml).toContain('Later box')
    expect(xml).not.toContain('Rejected box')
  })

  it('does not write a part only a failed op touched', async () => {
    const opened = await openPptx(await createBlankPptx())
    const masterBefore = opened.archive.readText(MASTER)
    const r = runTxn(opened, {
      isolation: 'per_op',
      ops: [
        failAfter({
          op: 'addElement',
          target: { part: MASTER },
          kind: 'textbox',
          offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
          paragraphs: [{ runs: [{ text: 'Rejected box' }] }],
        }),
        marker('Successful marker'),
      ],
    })

    expect(r.records).toHaveLength(1)
    expect(opened.archive.readText(MASTER)).toBe(masterBefore)
    expect(slideTexts(opened)).toEqual(['Successful marker'])
  })

  it('still restores the whole batch in atomic mode', async () => {
    const { opened, id } = await deckWithText()
    const r = runTxn(opened, {
      ops: [marker('First'), failAfter(linkedText(id, 'rejected'))],
    })

    expect(r.applied).toBe(false)
    expect(r.failures?.map((f) => f.index)).toEqual([1])
    expect(slideTexts(opened)).toEqual(['Original text'])
  })
})

describe('picture bullet format', () => {
  const badBullet = { bullet: 'blip', bulletImage: { base64: 'AA==', ext: 'not-real' } }

  it('is rejected by setText before anything is applied', async () => {
    const { opened, id } = await deckWithText()
    const op: Op = {
      op: 'setText',
      target: { slide: 0, el: id },
      paragraphs: [{ runs: [{ text: 'Rejected replacement' }], ...badBullet }],
    }

    const plan = runTxn(opened, { ops: [op], dryRun: true })
    expect(plan.failures?.[0]?.error).toMatch(/unsupported picture bullet format "\.not-real"/)

    const r = runTxn(opened, { isolation: 'per_op', ops: [op, marker('Successful marker')] })
    expect(r.failures?.map((f) => f.index)).toEqual([0])
    expect(slideTexts(opened)).toEqual(['Original text', 'Successful marker'])
  })

  it('is rejected by setParagraphFormat before anything is applied', async () => {
    const { opened, id } = await deckWithText()
    const plan = runTxn(opened, {
      ops: [{ op: 'setParagraphFormat', target: { slide: 0, el: id }, format: badBullet }],
      dryRun: true,
    })
    expect(plan.failures?.[0]?.error).toMatch(/unsupported picture bullet format "\.not-real"/)
  })
})
