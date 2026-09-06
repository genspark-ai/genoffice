/**
 * Every ```json example in prompts/ops/*.md must validate against the op
 * registry: the docs are what the model copies from, so a stale example is a
 * guaranteed failing tool call. Placeholder ids (e_TEXT, e_TABLE, SECTION_ID,
 * ...) are swapped for real ids from a fixture deck, then each example runs
 * through runTxn's dry run (plan/validate only, nothing mutates).
 */
import { describe, it, expect, beforeAll } from 'vitest'
import {
  addChart,
  addElement,
  addPicture,
  addSection,
  addSlideComment,
  addTable,
  createBlankPptx,
  elementDurableId,
  groupElements,
  insertBlankSlide,
  openPptx,
  slideDurableId,
  type GroupElement,
  type OpenedPptx,
  type SlideElement,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/main/ops'
import { OP_DOCS } from '../src/shared/op-docs'

const PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const box = (x: number, y: number) => ({ x, y, cx: 1828800, cy: 914400 })

let opened: OpenedPptx
let ids: Record<string, string>

beforeAll(async () => {
  opened = await openPptx(await createBlankPptx())
  const slide = () => opened.deck.slides[0]!
  // Raw-XML inserts (table / chart / group) reparse the slide and renumber the
  // parse-time ids, so build everything first and resolve ids from the final
  // slide by element type; the docs use durable e_* ids, so do the tests.
  addElement(slide(), {
    kind: 'textbox',
    offset: box(0, 0),
    paragraphs: [{ runs: [{ text: 'Title' }] }],
  })
  addElement(slide(), {
    kind: 'roundRect',
    offset: box(0, 1000000),
    paragraphs: [{ runs: [{ text: 'Card' }] }],
  })
  addElement(slide(), { kind: 'lineArrow', offset: box(0, 2000000) })
  expect(
    addPicture(opened, slide(), {
      bytes: new Uint8Array(Buffer.from(PNG_B64, 'base64')),
      ext: 'png',
      offset: box(0, 3000000),
    }),
  ).toBeTruthy()
  expect(addTable(opened, 0, { rows: 2, cols: 2, offset: box(2000000, 0) })).toBeTruthy()
  expect(
    addChart(opened, 0, {
      kind: 'bar',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
      offset: box(2000000, 1000000),
    }),
  ).toBeTruthy()
  const a = addElement(slide(), { kind: 'rect', offset: box(4000000, 0) }).id
  const b = addElement(slide(), { kind: 'ellipse', offset: box(4000000, 1000000) }).id
  expect(groupElements(opened, 0, [a, b])).toBeTruthy()
  expect(insertBlankSlide(opened, 0)).toBeTruthy()
  const section = addSection(opened, 0, 'Intro')![0]!.id
  expect(addSlideComment(opened, 0, { author: 'Reviewer', text: 'Looks good' })).toBeTruthy()

  const els = slide().elements
  const pick = (pred: (e: SlideElement) => boolean, what: string): string => {
    const el = els.find(pred)
    if (!el)
      throw new Error(
        `fixture: no ${what} on slide 0 (have ${els.map((e) => `${e.id}:${e.type}`).join(', ')})`,
      )
    return elementDurableId(el) ?? el.id
  }
  const group = els.find((e) => e.type === 'group') as GroupElement | undefined
  ids = {
    e_TEXT: pick((e) => e.type === 'text', 'text box'),
    e_SHAPE: pick((e) => e.type === 'shape' && e.id.startsWith('sp_'), 'shape'),
    e_LINE: pick((e) => e.id.startsWith('cxn_'), 'connector'),
    e_PICTURE: pick((e) => e.type === 'picture', 'picture'),
    e_TABLE: pick((e) => e.type === 'table', 'table'),
    e_CHART: pick((e) => e.type === 'chart', 'chart'),
    e_GROUP: pick((e) => e.type === 'group', 'group'),
    e_CHILD: group?.children[0]?.id ?? '',
    SECTION_ID: section,
  }
  // The fixture contract the docs rely on (see prompts/ops/_format.md)
  expect(opened.deck.slides).toHaveLength(2)
  expect(slideDurableId(opened.deck.slides[0]!)).toBe('s_1')
  for (const id of Object.values(ids)) expect(id).toBeTruthy()
})

function substitute(example: string): string {
  return example.replace(/\b(e_[A-Z]+|SECTION_ID)\b/g, (m) => {
    const real = ids[m]
    if (!real) throw new Error(`unknown placeholder ${m}; add it to the fixture or _format.md`)
    return real
  })
}

const callable = Object.entries(OP_DOCS).filter(
  ([, doc]) => doc.aiCallable !== false && !doc.pending,
)

describe('op doc examples', () => {
  it('every AI-callable op documents at least one example', () => {
    const missing = callable.filter(([, doc]) => doc.examples.length === 0).map(([n]) => n)
    expect(missing).toEqual([])
  })

  for (const [name, doc] of callable) {
    for (const [i, example] of doc.examples.entries()) {
      it(`${name} example #${i + 1} passes validation`, () => {
        const op = JSON.parse(substitute(example)) as { op: string }
        expect(op.op).toBe(name)
        const r = runTxn(opened, { ops: [op], dryRun: true })
        expect(r.failures ?? []).toEqual([])
        expect(r.plan).toHaveLength(1)
      })
    }
  }

  it('hidden ops carry no runnable examples (the model never sees them)', () => {
    const hiddenWithExamples = Object.entries(OP_DOCS)
      .filter(([, d]) => d.aiCallable === false && d.examples.length > 0)
      .map(([n]) => n)
    expect(hiddenWithExamples).toEqual([])
  })
})
