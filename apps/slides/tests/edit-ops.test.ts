/**
 * Canonical op layer (main/ops): registry validation with guided errors,
 * transaction executor semantics (atomic rollback / per_op / dry-run) — all
 * against a real in-memory deck (createBlankPptx + engine mutations), no mocks.
 */
import { describe, it, expect, beforeEach } from 'vitest'
import {
  addElement,
  createBlankPptx,
  getSlideAnimations,
  extractMergeSlideSource,
  openPptx,
  parseMasterPart,
  patchSlideXml,
  savePptx,
  type OpenedPptx,
  type SlideElement,
  type TextElement,
} from '@genoffice/pptx-engine'
import { runTxn, opNames, elementDurableId, slideDurableId } from '../src/main/ops'
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

  // Color/number payloads are pasted into XML attribute values — validation is
  // the only thing between a CSS color name and a schema-invalid srgbClr @val.
  it('setFill rejects non-hex color strings', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setFill', target: { slide: 0, el: cardId }, fill: 'red' }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('hex color')
  })

  it('setFill rejects a single-stop gradient', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setFill',
          target: { slide: 0, el: cardId },
          fill: { stops: [{ pos: 0, color: '#FF0000' }] },
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('at least two')
  })

  it('setStroke rejects an empty patch (NaN width would reach the XML)', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setStroke', target: { slide: 0, el: cardId }, stroke: {} }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('hex color')
  })

  it('setStroke rejects a non-finite width', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setStroke',
          target: { slide: 0, el: cardId },
          stroke: { color: '#000000', widthEmu: Number.NaN },
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('finite number')
  })

  it('setFont rejects a non-hex color', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setFont', target: { slide: 0, el: titleId }, font: { color: 'red' } }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('hex color')
  })

  it('setParagraphFormat rejects non-finite numerics', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setParagraphFormat',
          target: { slide: 0, el: titleId },
          format: { lineSpacingPct: Number.NaN },
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('finite number')
  })

  it('setBackground rejects CSS color names', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setBackground', target: { slide: 0 }, kind: 'solid', color: 'blue' }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('hex color')
  })

  it('setPictureOpacity rejects NaN (it would survive the engine clamp)', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setPictureOpacity', target: { slide: 0, el: cardId }, opacity: Number.NaN }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('finite number')
  })

  it('setPictureSrcRect rejects fractions outside 0..1', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setPictureSrcRect',
          target: { slide: 0, el: cardId },
          srcRect: { l: 2, t: 0, r: 0, b: 0 },
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('fraction 0..1')
  })

  it('setPictureSrcRect rejects opposing crops that erase the image', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setPictureSrcRect',
          target: { slide: 0, el: cardId },
          srcRect: { l: 0.6, t: 0, r: 0.6, b: 0 },
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('visible strip')
  })

  it('setAnimations resolves durable element refs and writes the timing target', () => {
    const durable = elementDurableId(els()[0]!)!
    expect(durable).toMatch(/^e_[0-9a-f]{8}$/)
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setAnimations',
          target: { slide: 0 },
          items: [
            { sourceId: durable, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
          ],
        },
      ],
    })
    expect(r.applied).toBe(true)
    const anims = getSlideAnimations(opened.deck.slides[0]!)
    expect(anims).toHaveLength(1)
    expect(anims[0]!.effect).toBe('fade')
  })

  it('setAnimations rejects an unresolvable sourceId instead of silently wiping', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setAnimations',
          target: { slide: 0 },
          items: [
            {
              sourceId: 'e_ghost',
              effect: 'fade',
              trigger: 'onClick',
              durationMs: 500,
              delayMs: 0,
            },
          ],
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('no element "e_ghost"')
  })

  it('setAnimations rejects unknown effect / trigger / non-finite timings', () => {
    const base = { sourceId: titleId, durationMs: 500, delayMs: 0 }
    for (const [patch, needle] of [
      [{ effect: 'sparkle', trigger: 'onClick' }, 'effect'],
      [{ effect: 'fade', trigger: 'sometime' }, 'trigger'],
      [{ effect: 'fade', trigger: 'onClick', durationMs: Number.NaN }, 'durationMs'],
      [{ effect: 'fade', trigger: 'onClick', delayMs: -1 }, 'delayMs'],
    ] as const) {
      const r = runTxn(opened, {
        ops: [{ op: 'setAnimations', target: { slide: 0 }, items: [{ ...base, ...patch }] }],
      })
      expect(r.applied).toBe(false)
      expect(r.failures![0]!.error).toContain(needle)
    }
  })

  it('setTransition rejects unmapped kinds (they would serialize as literal text)', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setTransition', target: { slide: 0 }, kind: 'flip' }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('one of')
  })

  it('setAdvanceTime rejects negative/NaN ms (advTm is unsigned)', () => {
    for (const ms of [-5, Number.NaN]) {
      const r = runTxn(opened, {
        ops: [{ op: 'setAdvanceTime', target: { slide: 0 }, ms }],
      })
      expect(r.applied).toBe(false)
      expect(r.failures![0]!.error).toContain('non-negative')
    }
  })

  it('setSlideSize rejects non-positive dimensions', () => {
    const r = runTxn(opened, { ops: [{ op: 'setSlideSize', cx: 0, cy: 6858000 }] })
    expect(r.applied).toBe(false)
  })

  it('addChart rejects empty categories/series with the actual reason', () => {
    const offset = { x: 0, y: 0, cx: 4572000, cy: 3200400 }
    const r1 = runTxn(opened, {
      ops: [
        {
          op: 'addChart',
          target: { slide: 0 },
          kind: 'bar',
          categories: [],
          series: [{ name: 'S', values: [1] }],
          offset,
        },
      ],
    })
    expect(r1.applied).toBe(false)
    expect(r1.failures![0]!.error).toContain('at least one category')
    const r2 = runTxn(opened, {
      ops: [
        {
          op: 'addChart',
          target: { slide: 0 },
          kind: 'bar',
          categories: ['A'],
          series: [],
          offset,
        },
      ],
    })
    expect(r2.applied).toBe(false)
    expect(r2.failures![0]!.error).toContain('at least one {name, values[]}')
  })

  it('insertSlidePptx rejects a source missing rels/media/layoutChain (TypeError otherwise)', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'insertSlidePptx',
          at: 0,
          source: { srcSlidePath: 'ppt/slides/slide1.xml', slideXml: '<p:sld/>' },
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('source.rels')
  })

  it('setConnectorEndpoints refuses non-connectors and unresolvable anchors, applies to connectors', () => {
    const onShape = runTxn(opened, {
      ops: [
        {
          op: 'setConnectorEndpoints',
          target: { slide: 0, el: cardId },
          p1: { x: 0, y: 0 },
          p2: { x: 914400, y: 914400 },
        },
      ],
    })
    expect(onShape.applied).toBe(false)
    expect(onShape.failures![0]!.error).toContain('not a connector')

    const cxn = addElement(opened.deck.slides[0]!, {
      kind: 'line',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
    })
    const badAnchor = runTxn(opened, {
      ops: [
        {
          op: 'setConnectorEndpoints',
          target: { slide: 0, el: cxn.id },
          p1: { x: 0, y: 0 },
          p2: { x: 914400, y: 914400 },
          start: { targetId: 'e_ghost', idx: 1 },
        },
      ],
    })
    expect(badAnchor.applied).toBe(false)
    expect(badAnchor.failures![0]!.error).toContain('does not resolve')

    // per_op does not roll back a failed apply: the bad anchor must reject
    // before the box mutation, or the connector stays moved on failure
    const beforeBox = { ...els().find((x) => x.id === cxn.id)!.transform.offset }
    const perOp = runTxn(opened, {
      isolation: 'per_op',
      ops: [
        {
          op: 'setConnectorEndpoints',
          target: { slide: 0, el: cxn.id },
          p1: { x: 457200, y: 457200 },
          p2: { x: 1828800, y: 1828800 },
          end: { targetId: 'e_ghost', idx: 0 },
        },
      ],
    })
    expect(perOp.applied).toBe(false)
    expect(els().find((x) => x.id === cxn.id)!.transform.offset).toEqual(beforeBox)

    const ok = runTxn(opened, {
      ops: [
        {
          op: 'setConnectorEndpoints',
          target: { slide: 0, el: cxn.id },
          p1: { x: 914400, y: 0 },
          p2: { x: 0, y: 914400 },
        },
      ],
    })
    expect(ok.applied).toBe(true)
    const live = els().find((x) => x.id === cxn.id)!
    expect(live.transform.offset).toEqual({ x: 0, y: 0, cx: 914400, cy: 914400 })
    expect(live.transform.flipH).toBe(true)
  })

  it('setPictureSrcRect accepts null per the documented contract', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setPictureSrcRect', target: { slide: 0, el: cardId }, srcRect: null }],
    })
    // null passes validation; the failure is the shape not being a picture
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('not a croppable picture')
  })
})

describe('media bytes over JSON (base64 / data URL)', () => {
  const PNG_B64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  const offset = { x: 0, y: 0, cx: 914400, cy: 914400 }

  it('addPicture accepts a base64 string (the AI tool channel cannot carry binary)', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'addPicture', target: { slide: 0 }, bytes: PNG_B64, ext: 'png', offset }],
    })
    expect(r.applied).toBe(true)
    expect(r.records![0]!.created).toHaveLength(1)
    expect(els().some((x) => x.type === 'picture')).toBe(true)
  })

  it('addPicture derives a missing ext from a data URL and replacePicture round-trips', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'addPicture',
          target: { slide: 0 },
          bytes: `data:image/png;base64,${PNG_B64}`,
          offset,
        },
      ],
    })
    expect(r.applied).toBe(true)
    const pic = els().find((x) => x.type === 'picture')!
    const r2 = runTxn(opened, {
      ops: [
        {
          op: 'replacePicture',
          target: { slide: 0, el: pic.id },
          bytes: `data:image/png;base64,${PNG_B64}`,
        },
      ],
    })
    expect(r2.applied).toBe(true)
  })

  it('setImageFill accepts data-URL bytes on a shape', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setImageFill',
          target: { slide: 0, el: cardId },
          source: { bytes: `data:image/png;base64,${PNG_B64}` },
        },
      ],
    })
    expect(r.applied).toBe(true)
    expect((els().find((x) => x.id === cardId) as TextElement).fill?.type).toBe('image')
  })

  it('garbage byte strings are rejected with a guided error', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'addPicture', target: { slide: 0 }, bytes: 'not base64!!', ext: 'png', offset }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('base64')
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

describe('setText on a shape that has no text body', () => {
  /** Strip the <p:txBody> the insert helper writes: AI/converter output and shapes
   *  whose text was deleted elsewhere reach us with no text body at all. */
  const stripTxBody = (xml: string) => xml.replace(/<p:txBody>[\s\S]*<\/p:txBody>/, '')

  /** An autoshape with no text body, standing alone on the slide. */
  function bodylessShape(): string {
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
    })
    el.anchor.originalXml = stripTxBody(el.anchor.originalXml)
    delete el.text
    return el.id
  }

  /** The parser sets these from <p:ph> / <p:cNvSpPr txBox="1">; the op reads the model. */
  const flagged = (id: string, patch: Partial<TextElement>) => {
    Object.assign(els().find((x) => x.id === id) as TextElement, patch)
    return id
  }

  /** The regenerated <p:sp> carrying `text`, so assertions can name one shape's bytes. */
  const spWithText = (text: string) =>
    patchSlideXml(opened.deck.slides[0]!)
      .match(/<p:sp>[\s\S]*?<\/p:sp>/g)!
      .find((sp) => sp.includes(`<a:t>${text}</a:t>`))!

  const type = (el: string, text: string, group?: string) =>
    runTxn(opened, {
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el },
          paragraphs: [{ runs: [{ text }] }],
          ...(group ? { group } : {}),
        },
      ],
    })

  it('creates the body PowerPoint would have created', () => {
    const id = bodylessShape()
    expect(type(id, 'Typed').applied).toBe(true)
    const el = els().find((x) => x.id === id) as TextElement
    expect(el.text?.paragraphs[0]?.runs[0]?.text).toBe('Typed')
    // Centered like PowerPoint, with the parser's bodyPr insets so the live render
    // and the reopened file agree
    expect(el.text?.anchor).toBe('middle')
    expect(el.text?.paragraphs[0]?.align).toBe('center')
    expect(el.text?.insets).toEqual({ l: 91440, t: 45720, r: 91440, b: 45720 })
    // and both halves of the centering reach the bytes
    expect(spWithText('Typed')).toContain('anchor="ctr"')
    expect(spWithText('Typed')).toContain('algn="ctr"')
  })

  // Centering is an autoshape default. Across 421 PowerPoint-authored decks ~90% of
  // placeholders carry neither anchor nor algn — they inherit from the layout/master,
  // and baking the attributes in would override the template.
  it('a placeholder inherits instead of centering', () => {
    const id = flagged(bodylessShape(), { placeholder: 'body' })
    expect(type(id, 'In a placeholder').applied).toBe(true)
    const el = els().find((x) => x.id === id) as TextElement
    expect(el.text?.anchor).toBeUndefined()
    expect(el.text?.paragraphs[0]?.align).toBeUndefined()
    const sp = spWithText('In a placeholder')
    expect(sp).not.toContain('anchor=')
    expect(sp).not.toContain('algn=')
  })

  it('a text box stays top-left', () => {
    const id = flagged(bodylessShape(), { txBox: true })
    expect(type(id, 'In a text box').applied).toBe(true)
    const el = els().find((x) => x.id === id) as TextElement
    expect(el.text?.anchor).toBeUndefined()
    expect(el.text?.paragraphs[0]?.align).toBeUndefined()
    expect(spWithText('In a text box')).not.toContain('anchor=')
  })

  it('the text survives a save → reopen round trip', async () => {
    expect(type(bodylessShape(), 'Persisted').applied).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    const found = reopened.deck.slides[0]!.elements.find((e) =>
      (e as TextElement).text?.paragraphs.some((p) => p.runs.some((r) => r.text === 'Persisted')),
    ) as TextElement | undefined
    expect(found?.text?.anchor).toBe('middle')
  })

  it('an explicit alignment from the caller wins over the centering default', () => {
    const id = bodylessShape()
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el: id },
          paragraphs: [{ runs: [{ text: 'Left' }], align: 'left' }],
        },
      ],
    })
    expect(r.applied).toBe(true)
    expect((els().find((x) => x.id === id) as TextElement).text?.paragraphs[0]?.align).toBe('left')
  })

  it('a group child gains a body too', () => {
    const childId = bodylessShape()
    const grouped = runTxn(opened, {
      ops: [{ op: 'groupElements', target: { slide: 0 }, els: [childId, cardId] }],
    })
    expect(grouped.applied).toBe(true)
    const grp = els().find((x) => x.type === 'group')!
    // Group children carry an empty byte anchor — the stripped child lives inside the
    // group's single blob, which is the path the op has to look the bytes up through
    const child = (grp as unknown as { children: TextElement[] }).children.find((c) => !c.text)!
    expect(type(child.id, 'In group', grp.id).applied).toBe(true)
    expect(patchSlideXml(opened.deck.slides[0]!)).toContain('In group')
  })

  it('refuses a connector — CT_Connector has no txBody child', () => {
    const lineId = addElement(opened.deck.slides[0]!, {
      kind: 'line',
      offset: { x: 0, y: 0, cx: 914400, cy: 0 },
    }).id
    const r = type(lineId, 'on a line')
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('connector')
    expect(patchSlideXml(opened.deck.slides[0]!)).not.toContain('on a line')
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

  it('records stamp the acted-on slide durable id, immune to later index shifts', () => {
    const dup = runTxn(opened, { ops: [{ op: 'duplicateSlide', target: { slide: 0 } }] })
    expect(dup.applied).toBe(true)
    const r = runTxn(opened, {
      ops: [
        { op: 'setNotes', target: { slide: 1 }, text: 'edited' },
        { op: 'deleteSlide', target: { slide: 0 } },
      ],
    })
    expect(r.applied).toBe(true)
    // Numeric index 1 is gone after the delete; the stamp still finds the slide.
    const stamped = r.records![0]!.slideId
    expect(stamped).toBeTruthy()
    expect(opened.deck.slides.findIndex((s) => slideDurableId(s) === stamped)).toBe(0)
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

describe('insertSlidePptx (generated-page landing)', () => {
  const oneSlideSource = async (text: string) => {
    const src = await openPptx(await createBlankPptx())
    addElement(src.deck.slides[0]!, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
      paragraphs: [{ runs: [{ text }] }],
    })
    const source = await extractMergeSlideSource(await savePptx(src))
    expect(source).not.toBeNull()
    return source!
  }

  it('lands a page at the end and reports its durable id in created', async () => {
    const source = await oneSlideSource('Landed')
    const r = runTxn(opened, { ops: [{ op: 'insertSlidePptx', source }] })
    expect(r.applied).toBe(true)
    expect(opened.deck.slides).toHaveLength(2)
    expect(r.records![0]!.created![0]).toBe(slideDurableId(opened.deck.slides[1]!))
  })

  it('replace mode swaps the page at `at` without changing the count', async () => {
    const source = await oneSlideSource('Replacement')
    const r = runTxn(opened, {
      ops: [{ op: 'insertSlidePptx', source, at: 0, replace: true }],
    })
    expect(r.applied).toBe(true)
    expect(opened.deck.slides).toHaveLength(1)
    expect(slideDurableId(opened.deck.slides[0]!)).toBe(r.records![0]!.created![0])
  })

  it('rejects replace without a valid at', async () => {
    const source = await oneSlideSource('Rejected')
    const r = runTxn(opened, { ops: [{ op: 'insertSlidePptx', source, replace: true }] })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('replace')
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

describe('XML-invalid characters are rejected at plan time', () => {
  const vt = String.fromCharCode(0x0b)

  it('setText with a vertical tab fails guided, nothing applied', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el: titleId },
          paragraphs: [{ runs: [{ text: `PDF${vt}line` }] }],
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('paragraphs[0].runs[0].text')
    expect(r.failures![0]!.error).toContain('\\u000B')
    const title = els().find((e) => e.id === titleId) as TextElement
    expect(title.text!.paragraphs[0]!.runs[0]!.text).toBe('Title')
  })

  it('tab and newline are legal XML and pass', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el: titleId },
          paragraphs: [{ runs: [{ text: 'a\tb\nc' }] }],
        },
      ],
    })
    expect(r.applied).toBe(true)
  })

  it('covers every string field, not just text payloads', () => {
    const r = runTxn(opened, {
      ops: [{ op: 'setNotes', target: { slide: 0 }, text: `n${String.fromCharCode(0)}` }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('"text"')
  })
})

describe('insert-time options (genpptx parity ops)', () => {
  it('addTable rejects a length-mismatched colWidthsEmu with a guided error', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'addTable',
          target: { slide: 0 },
          rows: 2,
          cols: 3,
          offset: { x: 0, y: 0, cx: 3000000, cy: 1000000 },
          colWidthsEmu: [1000000, 2000000],
        },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('exactly 3 positive EMU values')
  })

  it('addTable applies explicit widths/heights and cell merges', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'addTable',
          target: { slide: 0 },
          rows: 2,
          cols: 2,
          offset: { x: 0, y: 0, cx: 3000000, cy: 1000000 },
          colWidthsEmu: [1000000, 2000000],
          rowHeightsEmu: [400000, 600000],
          cellProps: [[{ gridSpan: 2 }, { hMerge: true }], []],
        },
      ],
    })
    expect(r.applied).toBe(true)
    const el = els().find((e) => e.id === r.records![0]!.created![0]) as SlideElement
    expect(el.anchor.originalXml).toContain('<a:gridCol w="1000000"/><a:gridCol w="2000000"/>')
    expect(el.anchor.originalXml).toContain('<a:tc gridSpan="2">')
  })

  it('addChart validates colorScheme entries and writes per-series colors', () => {
    const bad = runTxn(opened, {
      ops: [
        {
          op: 'addChart',
          target: { slide: 0 },
          kind: 'bar',
          categories: ['A'],
          series: [{ name: 'S', values: [1] }],
          offset: { x: 0, y: 0, cx: 3000000, cy: 2000000 },
          colorScheme: ['red'],
        },
      ],
    })
    expect(bad.applied).toBe(false)
    expect(bad.failures![0]!.error).toContain('colorScheme[]')

    const ok = runTxn(opened, {
      ops: [
        {
          op: 'addChart',
          target: { slide: 0 },
          kind: 'doughnut',
          categories: ['A', 'B'],
          series: [{ name: 'S', values: [1, 2] }],
          offset: { x: 0, y: 0, cx: 3000000, cy: 2000000 },
          colorScheme: ['#C00000'],
          holeSizePct: 72,
        },
      ],
    })
    expect(ok.applied).toBe(true)
    const chartXml = [...opened.archive.entries.keys()]
      .filter((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
      .map((p) => Buffer.from(opened.archive.entries.get(p) as Uint8Array).toString('utf8'))
      .join('')
    expect(chartXml).toContain('<a:srgbClr val="C00000"/>')
    expect(chartXml).toContain('<c:holeSize val="72"/>')
  })

  it('addElement carries adjustments and bodyPr.autoFit into the shape XML', () => {
    const bad = runTxn(opened, {
      ops: [
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'roundRect',
          offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
          bodyPr: { autoFit: 'grow' },
        },
      ],
    })
    expect(bad.applied).toBe(false)
    expect(bad.failures![0]!.error).toContain('"shrink"')

    const r = runTxn(opened, {
      ops: [
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'roundRect',
          offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
          adjustments: { adj: 25000 },
          bodyPr: { autoFit: 'shrink' },
          paragraphs: [{ runs: [{ text: 'chip' }] }],
        },
      ],
    })
    expect(r.applied).toBe(true)
    const el = els().find((e) => e.id === r.records![0]!.created![0]) as SlideElement
    expect(el.anchor.originalXml).toContain('<a:gd name="adj" fmla="val 25000"/>')
    expect(el.anchor.originalXml).toContain('<a:normAutofit/>')
  })
})
