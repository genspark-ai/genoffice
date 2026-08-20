/**
 * Canonical op layer (main/ops): registry validation with guided errors,
 * transaction executor semantics (atomic rollback / per_op / dry-run) — all
 * against a real in-memory deck (createBlankPptx + engine mutations), no mocks.
 */
import { describe, it, expect, beforeEach } from 'vitest'
import {
  addElement,
  createBlankPptx,
  openPptx,
  parseMasterPart,
  patchSlideXml,
  type OpenedPptx,
  type TextElement,
} from '@genoffice/pptx-engine'
import { runTxn, opNames, elementDurableId } from '../src/main/ops'
import { mapScriptOps } from '../src/main/ops/script-map'

let opened: OpenedPptx
let titleId: string
let cardId: string

beforeEach(async () => {
  opened = await openPptx(await createBlankPptx())
  const slide = opened.deck.slides[0]!
  titleId = addElement(slide, {
    kind: 'textbox',
    offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
    paragraphs: [{ runs: [{ text: 'Title' }] }],
  }).id
  cardId = addElement(slide, {
    kind: 'roundRect',
    offset: { x: 0, y: 500000, cx: 914400, cy: 457200 },
    fillColor: '#FFFFFF',
  }).id
})

const els = () => opened.deck.slides[0]!.elements

describe('op validation (guided errors)', () => {
  it('unknown op name lists the supported vocabulary', () => {
    const r = runTxn(opened, { ops: [{ op: 'sparkle' }] })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('unknown op "sparkle"')
    for (const name of opNames()) expect(r.failures![0]!.error).toContain(name)
  })

  it('unknown element lists the ids that do exist', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'deleteElement', target: { slide: 0, el: 'ghost' } }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('no element "ghost" on slide 0')
    expect(r.failures![0]!.error).toContain(titleId)
    expect(r.failures![0]!.error).toContain(cardId)
    expect(els()).toHaveLength(2) // nothing applied
  })

  it('slide index out of range is guided', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'deleteElement', target: { slide: 9, el: titleId } }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('out of range (0-0)')
  })

  it('type-restricted ops name the allowed types', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setFill', target: { slide: 0, el: titleId }, fill: 12 as unknown as string }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('needs "fill"')
  })
})

describe('atomic transactions', () => {
  it('applies a mixed batch and journals before/after per op', () => {
    const r = runTxn(opened, {
      ops: [
        { op: 'setFill', target: { slide: 0, el: cardId }, fill: '#112233' },
        {
          op: 'setStroke',
          target: { slide: 0, el: cardId },
          stroke: { color: '#445566', widthEmu: 12700 },
        },
        { op: 'deleteElement', target: { slide: 0, el: titleId } },
      ],
    })
    expect(r.applied).toBe(true)
    expect(r.records).toHaveLength(3)
    expect(r.records![2]!.before).toEqual({ type: 'text' })
    expect(els()).toHaveLength(1)
    const card = els()[0] as TextElement
    expect(card.fill).toEqual({ type: 'solid', color: '#112233' })
    expect(card.stroke?.width).toBe(12700)
  })

  it('plan-time failure anywhere means nothing is applied', () => {
    const r = runTxn(opened, {
      ops: [
        { op: 'setFill', target: { slide: 0, el: cardId }, fill: '#112233' },
        { op: 'deleteElement', target: { slide: 0, el: 'ghost' } },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('Nothing was applied (atomic)')
    expect((els()[1] as TextElement).fill).toEqual({ type: 'solid', color: '#FFFFFF' })
  })

  it('apply-time failure rolls the deck back to the pre-transaction state', () => {
    // Both deletes validate against the pre-txn state; the second fails at apply time
    const r = runTxn(opened, {
      ops: [
        { op: 'deleteElement', target: { slide: 0, el: titleId } },
        { op: 'deleteElement', target: { slide: 0, el: titleId } },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.index).toBe(1)
    expect(els()).toHaveLength(2) // first delete rolled back
    expect(els().some((x) => x.id === titleId)).toBe(true)
  })
})

describe('per_op isolation and dry-run', () => {
  it('per_op keeps successes and reports failures by index', () => {
    const r = runTxn(opened, {
      isolation: 'per_op',
      ops: [
        { op: 'deleteElement', target: { slide: 0, el: titleId } },
        { op: 'deleteElement', target: { slide: 0, el: titleId } },
        { op: 'setFill', target: { slide: 0, el: cardId }, fill: '#0000FF' },
      ],
    })
    expect(r.applied).toBe(true)
    expect(r.records).toHaveLength(2)
    expect(r.failures).toHaveLength(1)
    expect(r.failures![0]!.index).toBe(1)
    expect(els()).toHaveLength(1)
  })

  it('dry-run returns the plan and touches nothing', () => {
    const r = runTxn(opened, {
      dryRun: true,
      ops: [
        { op: 'deleteElement', target: { slide: 0, el: titleId } },
        { op: 'deleteElement', target: { slide: 0, el: 'ghost' } },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.dryRun).toBe(true)
    expect(r.plan).toEqual([`[0] deleteElement s0/${titleId}`])
    expect(r.failures).toHaveLength(1)
    expect(els()).toHaveLength(2)
  })
})

describe('cross-family ops on a real deck', () => {
  it('addElement mints an id; setTransform/setTextAnchor/reorder then act on it', () => {
    const added = runTxn(opened, {
      ops: [
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'ellipse',
          offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
          fill: '#00FF00',
        },
      ],
    })
    expect(added.applied).toBe(true)
    const newId = added.records![0]!.created![0]!
    expect(els().some((x) => x.id === newId)).toBe(true)
    const edit = runTxn(opened, {
      ops: [
        {
          op: 'setTransform',
          target: { slide: 0, el: newId },
          box: { x: 914400, y: 914400, cx: 457200, cy: 457200 },
          rotDeg: 45,
        },
        { op: 'reorderElement', target: { slide: 0, el: newId }, dir: 'back' },
      ],
    })
    expect(edit.applied).toBe(true)
    const el = els().find((x) => x.id === newId)!
    expect(el.transform.offset.x).toBe(914400)
    expect(el.transform.rot).toBe(45 * 60000)
    expect(els()[0]!.id).toBe(newId) // sent to back
  })

  it('slide lifecycle: duplicate, hide, find-replace, delete — one vocabulary', () => {
    // Plan-time validation runs against the pre-transaction state, so ops that
    // address a slide minted earlier in the same batch go in a follow-up txn.
    const dup = runTxn(opened, { ops: [{ op: 'duplicateSlide', target: { slide: 0 } }] })
    expect(dup.applied).toBe(true)
    expect(opened.deck.slides).toHaveLength(2)
    const r = runTxn(opened, {
      ops: [
        { op: 'setHidden', target: { slide: 1 }, hidden: true },
        { op: 'findReplace', find: 'Title', replace: 'Heading' },
        { op: 'setNotes', target: { slide: 0 }, text: 'speaker notes' },
      ],
    })
    expect(r.applied).toBe(true)
    expect((r.records![1]!.after as { count: number }).count).toBe(2) // both copies replaced
    const del = runTxn(opened, { ops: [{ op: 'deleteSlide', target: { slide: 1 } }] })
    expect(del.applied).toBe(true)
    expect(opened.deck.slides).toHaveLength(1)
  })

  it('atomic rollback spans families: a failing slide op undoes an element op', () => {
    const r = runTxn(opened, {
      ops: [
        { op: 'setFill', target: { slide: 0, el: cardId }, fill: '#FF0000' },
        { op: 'deleteSlide', target: { slide: 7 } },
      ],
    })
    expect(r.applied).toBe(false)
    expect((els()[1] as TextElement).fill).toEqual({ type: 'solid', color: '#FFFFFF' })
  })
})

describe('group-addressed ops', () => {
  it('setFill with an unknown group lists available groups', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setFill', target: { slide: 0, el: 'c1' }, group: 'g9', fill: '#123456' }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('no group "g9"')
  })
})

describe('script-map: edit-script primitives → one op transaction', () => {
  const req = (boxes: unknown[], edits: unknown[]) =>
    ({ slideIndex: 0, fitWidthPx: 1280, boxes, edits }) as Parameters<typeof mapScriptOps>[1]

  it('maps px boxes and ordered edits onto ops; runTxn applies them on the real deck', () => {
    const ops = mapScriptOps(
      opened,
      req(
        [{ id: cardId, x: 100, y: 50, w: 200, h: 80, rotation: 15 }],
        [
          { kind: 'text', id: titleId, paragraphs: [{ runs: [{ text: 'Scripted' }] }] },
          { kind: 'style', id: titleId, style: { fontSize: 30, bold: true, align: 'center' } },
          { kind: 'fill', id: cardId, fill: '#ABCDEF' },
          { kind: 'stroke', id: cardId, stroke: { color: '#112233', widthPt: 2 } },
        ],
      ),
    )
    // style splits into run-level setFont + paragraph-level setParagraphFormat
    expect(ops.map((o) => o.op)).toEqual([
      'setTransform',
      'setText',
      'setFont',
      'setParagraphFormat',
      'setFill',
      'setStroke',
    ])
    const r = runTxn(opened, { ops })
    expect(r.applied).toBe(true)
    const card = els().find((x) => x.id === cardId)!
    // deck is 1280px wide at fitWidthPx 1280 → scale 1 → px * 9525 EMU
    expect(card.transform.offset).toEqual({ x: 952500, y: 476250, cx: 1905000, cy: 762000 })
    expect(card.transform.rot).toBe(15 * 60000)
    expect((card as TextElement).fill).toEqual({ type: 'solid', color: '#ABCDEF' })
    expect((card as TextElement).stroke?.width).toBe(2 * 12700)
    const title = els().find((x) => x.id === titleId) as TextElement
    const run = title.text!.paragraphs[0]!.runs[0]!
    expect(run.text).toBe('Scripted')
    expect(run.fontSize).toBe(30)
    expect(run.bold).toBe(true)
    expect(title.text!.paragraphs[0]!.align).toBe('center')
  })

  it('a bad primitive fails the whole mapped transaction atomically', () => {
    const ops = mapScriptOps(
      opened,
      req(
        [{ id: cardId, x: 10, y: 10, w: 50, h: 50, rotation: 0 }],
        [{ kind: 'fill', id: 'ghost', fill: '#FF0000' }],
      ),
    )
    const r = runTxn(opened, { ops })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('no element "ghost"')
    // The transform earlier in the batch was not applied (plan-time failure)
    expect(els().find((x) => x.id === cardId)!.transform.offset.x).toBe(0)
  })
})

describe('script-map group children: apply-time conversion', () => {
  it('a group and its child moved by the same delta keep the child in place locally (no double shift)', () => {
    const slide0 = opened.deck.slides[0]!
    const a = addElement(slide0, {
      kind: 'rect',
      offset: { x: 914400, y: 914400, cx: 914400, cy: 914400 },
      fillColor: '#111111',
    })
    const b = addElement(slide0, {
      kind: 'rect',
      offset: { x: 2286000, y: 914400, cx: 914400, cy: 914400 },
      fillColor: '#222222',
    })
    const g = runTxn(opened, {
      ops: [{ op: 'groupElements', target: { slide: 0 }, els: [a.id, b.id] }],
    })
    expect(g.applied).toBe(true)
    const groupId = g.records![0]!.created![0]!
    const grp = els().find((x) => x.id === groupId)! as unknown as {
      transform: { offset: { x: number; y: number; cx: number; cy: number } }
      children: Array<{
        id: string
        transform: { offset: { x: number; y: number; cx: number; cy: number } }
      }>
    }
    const child = grp.children[0]!
    const localBefore = { ...child.transform.offset }
    const gBefore = { ...grp.transform.offset }
    // Freshly grouped: chOff equals the group's offset and the scale is 1, so the
    // child's document-space box equals its stored offset
    const px = (emu: number) => emu / 9525
    const delta = 10 // px
    const ops = mapScriptOps(opened, {
      slideIndex: 0,
      fitWidthPx: 1280,
      boxes: [
        {
          id: groupId,
          x: px(gBefore.x) + delta,
          y: px(gBefore.y),
          w: px(gBefore.cx),
          h: px(gBefore.cy),
          rotation: 0,
        },
        {
          id: child.id,
          groupId,
          x: px(localBefore.x) + delta,
          y: px(localBefore.y),
          w: px(localBefore.cx),
          h: px(localBefore.cy),
          rotation: 0,
        },
      ],
      edits: [],
    })
    // Top-level group move is ordered before its children
    expect(ops[0]!.target!.el).toBe(groupId)
    const r = runTxn(opened, { ops })
    expect(r.applied).toBe(true)
    // The group moved once; the child stayed put in group-local space — the abs
    // box converted against the group's post-move state, not the stale one
    expect(grp.transform.offset.x).toBe(gBefore.x + delta * 9525)
    expect(child.transform.offset).toEqual(localBefore)
  })
})

describe('part addressing (master/layout chrome)', () => {
  const masterPath = 'ppt/slideMasters/slideMaster1.xml'
  let chromeId: string

  beforeEach(() => {
    const master = parseMasterPart(opened.archive, masterPath)!
    addElement(master, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
      paragraphs: [{ runs: [{ text: 'Chrome' }] }],
    })
    opened.archive.entries.set(masterPath, Buffer.from(patchSlideXml(master), 'utf8'))
    // Parse-time ids are re-minted on every parse; parts are re-parsed per
    // transaction, so chrome elements are addressed by durable id
    chromeId = elementDurableId(parseMasterPart(opened.archive, masterPath)!.elements[0]!)!
  })

  const reparse = () => parseMasterPart(opened.archive, masterPath)!

  it('setText/setFill/setTransform accept target.part and flush to the entry', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setText',
          target: { part: masterPath, el: chromeId },
          paragraphs: [{ runs: [{ text: 'Edited' }] }],
        },
        { op: 'setFill', target: { part: masterPath, el: chromeId }, fill: '#112233' },
        {
          op: 'setTransform',
          target: { part: masterPath, el: chromeId },
          box: { x: 10, y: 20, cx: 30000, cy: 40000 },
          rotDeg: 0,
        },
      ],
    })
    expect(r.applied).toBe(true)
    const el = reparse().elements.find((x) => elementDurableId(x) === chromeId) as TextElement
    expect(el.text!.paragraphs[0]!.runs[0]!.text).toBe('Edited')
    expect(el.fill).toMatchObject({ type: 'solid', color: '#112233' })
    expect(el.transform.offset).toMatchObject({ x: 10, y: 20 })
  })

  it('deleteElement accepts target.part', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'deleteElement', target: { part: masterPath, el: chromeId } }],
    })
    expect(r.applied).toBe(true)
    expect(reparse().elements.find((x) => elementDurableId(x) === chromeId)).toBeUndefined()
  })

  it('unknown part is guided with the available part list', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setFill',
          target: { part: 'ppt/slideMasters/nope.xml', el: chromeId },
          fill: '#000000',
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('no master/layout part')
    expect(r.failures![0]!.error).toContain(masterPath)
  })

  it('ops without part support reject part targets', () => {
    const r = runTxn(opened, { ops: [{ op: 'duplicateSlide', target: { part: masterPath } }] })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('cannot target a master/layout part')
  })

  it('seeded live parts are mutated in place — ids stay stable for the caller', () => {
    const live = parseMasterPart(opened.archive, masterPath)!
    const liveEl = live.elements[0]!
    const liveId = liveEl.id
    const r = runTxn(opened, {
      parts: new Map([[masterPath, live]]),
      ops: [{ op: 'setFill', target: { part: masterPath, el: liveId }, fill: '#ABCDEF' }],
    })
    expect(r.applied).toBe(true)
    // The seeded object itself carries the change and keeps its parse-time id
    expect(liveEl.id).toBe(liveId)
    expect((liveEl as TextElement).fill).toMatchObject({ type: 'solid', color: '#ABCDEF' })
    const baked = Buffer.from(opened.archive.entries.get(masterPath) as Uint8Array).toString('utf8')
    expect(baked.toUpperCase()).toContain('ABCDEF')
  })

  it('atomic rollback leaves the part entry untouched', () => {
    const before = opened.archive.entries.get(masterPath)
    const r = runTxn(opened, {
      ops: [
        { op: 'setFill', target: { part: masterPath, el: chromeId }, fill: '#FF0000' },
        { op: 'deleteElement', target: { slide: 0, el: 'ghost' } },
      ],
    })
    expect(r.applied).toBe(false)
    expect(opened.archive.entries.get(masterPath)).toBe(before)
    const el = reparse().elements.find((x) => elementDurableId(x) === chromeId) as TextElement
    expect(el.fill).not.toMatchObject({ type: 'solid', color: '#FF0000' })
  })
})

describe('applyTheme op', () => {
  it('background fallback does not clobber the explicit color remap', () => {
    const r0 = runTxn(opened, {
      ops: [{ op: 'setFill', target: { slide: 0, el: cardId }, fill: '#4472C4' }], // old accent1
    })
    expect(r0.applied).toBe(true)
    const r = runTxn(opened, {
      ops: [{ op: 'applyTheme', name: 'Dark', colors: { accent1: '#00FF11', lt1: '#112233' } }],
    })
    expect(r.applied).toBe(true)
    const slideXml = Buffer.from(
      opened.archive.entries.get(opened.deck.slides[0]!.path) as Uint8Array,
    ).toString('utf8')
    // The remap rebinds the explicit old-accent fill to the scheme slot; the
    // in-op model refresh must not let a later materialize write the stale
    // pre-remap bytes back over it
    expect(slideXml.toUpperCase()).not.toContain('4472C4')
    // The refreshed model resolves the new theme (master bg -> lt1), so the
    // fallback correctly leaves inherited backgrounds alone
    expect(opened.deck.slides[0]!.background).toMatchObject({ color: '#112233' })
  })

  it('patches the theme part and reports counts', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'applyTheme', name: 'Test', colors: { accent1: '#FF0000', lt1: '#112233' } }],
    })
    expect(r.applied).toBe(true)
    const after = r.records![0]!.after as { patched: number }
    expect(after.patched).toBeGreaterThan(0)
    const themeXml = Buffer.from(
      opened.archive.entries.get('ppt/theme/theme1.xml') as Uint8Array,
    ).toString('utf8')
    expect(themeXml.toUpperCase()).toContain('FF0000')
  })

  it('validates the colors record with guided errors', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'applyTheme', name: 'Bad', colors: { accent1: 'red' } }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('colors.accent1')
  })
})
