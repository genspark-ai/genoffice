/**
 * Durable identity (design step 0): elements born through the op layer mint an
 * a16:creationId in their bytes; durable ids ("e_<guid8>", falling back to
 * "e_<cNvPr id>") and slide ids ("s_<n>") stay resolvable across save→reopen,
 * reparse, group/ungroup, and paste (which must mint NEW identities).
 */
import { describe, it, expect, beforeEach } from 'vitest'
import {
  addElement,
  copyElementData,
  createBlankPptx,
  openPptx,
  savePptx,
  type OpenedPptx,
  type TextElement,
} from '@genoffice/pptx-engine'
import { runTxn, elementDurableId, slideDurableId } from '../src/main/ops'
import { elementCNvPrId, groupChildDurableId, patchSlideXml } from '@genoffice/pptx-engine'

let opened: OpenedPptx

beforeEach(async () => {
  opened = await openPptx(await createBlankPptx())
})

const els = () => opened.deck.slides[0]!.elements

describe('durable ids', () => {
  it('newborn elements mint a creationId; durable ids derive from it', () => {
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    expect(el.anchor.originalXml).toContain('a16:creationId')
    const durable = elementDurableId(el)!
    expect(durable).toMatch(/^e_[0-9a-f]{8}$/)
    expect(slideDurableId(opened.deck.slides[0]!)).toBe('s_1')
  })

  it('elements without a creationId fall back to the persisted cNvPr id', () => {
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    el.anchor.originalXml = el.anchor.originalXml.replace(/<a:extLst>.*?<\/a:extLst>/, '')
    expect(elementDurableId(el)).toMatch(/^e_\d+$/)
  })

  it('durable targets survive save→reopen (the step-0 acceptance test)', async () => {
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    const durable = elementDurableId(el)!
    const parseIdBefore = el.id

    const reopened = await openPptx(await savePptx(opened))
    const slide = reopened.deck.slides[0]!
    // The parse-time id changed; the durable id did not
    expect(slide.elements.some((x) => x.id === parseIdBefore)).toBe(false)
    expect(slide.elements.some((x) => elementDurableId(x) === durable)).toBe(true)

    // Old op targets still resolve: durable slide + element addressing
    const r = runTxn(reopened, {
      ops: [{ op: 'setFill', target: { slide: 's_1', el: durable }, fill: '#ABCDEF' }],
    })
    expect(r.applied).toBe(true)
    const target = slide.elements.find((x) => elementDurableId(x) === durable) as TextElement
    expect(target.fill).toEqual({ type: 'solid', color: '#ABCDEF' })
  })

  it('group → ungroup keeps the children durable', () => {
    const a = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#111111',
    })
    const b = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 1828800, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#222222',
    })
    const da = elementDurableId(a)!
    const db = elementDurableId(b)!
    const g = runTxn(opened, {
      ops: [{ op: 'groupElements', target: { slide: 0 }, els: [a.id, b.id] }],
    })
    expect(g.applied).toBe(true)
    const groupId = g.records![0]!.created![0]!
    const u = runTxn(opened, {
      ops: [{ op: 'ungroupElement', target: { slide: 0, el: groupId } }],
    })
    expect(u.applied).toBe(true)
    const durables = els().map((x) => elementDurableId(x))
    expect(durables).toContain(da)
    expect(durables).toContain(db)
  })

  it('a group wrapper without its own creationId never inherits a child GUID', () => {
    const a = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#111111',
    })
    const b = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 1828800, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#222222',
    })
    const childDurables = [elementDurableId(a)!, elementDurableId(b)!]
    const g = runTxn(opened, {
      ops: [{ op: 'groupElements', target: { slide: 0 }, els: [a.id, b.id] }],
    })
    const grp = els().find((x) => x.id === g.records![0]!.created![0]!)!
    // Simulate a foreign group: strip the creationId from the GROUP's own cNvPr
    // (the first extLst in the fragment) while the children keep theirs
    grp.anchor.originalXml = grp.anchor.originalXml.replace(/<a:extLst>.*?<\/a:extLst>/, '')
    const durable = elementDurableId(grp)!
    expect(durable).toMatch(/^e_\d+$/) // cNvPr fallback, not a child GUID
    expect(childDurables).not.toContain(durable)
  })

  it('pasted copies mint NEW identities (no same-slide durable collisions)', () => {
    const slide = opened.deck.slides[0]!
    const src = addElement(slide, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    const srcDurable = elementDurableId(src)!
    const items = [copyElementData(opened, slide, src)]
    const r = runTxn(opened, {
      ops: [{ op: 'pasteElements', target: { slide: 0 }, items, dx: 9525, dy: 9525 }],
    })
    expect(r.applied).toBe(true)
    const pastedId = r.records![0]!.created![0]!
    const pasted = els().find((x) => x.id === pastedId)!
    const pastedDurable = elementDurableId(pasted)!
    expect(pasted.anchor.originalXml).toContain('a16:creationId')
    expect(pastedDurable).toMatch(/^e_[0-9a-f]{8}$/)
    expect(pastedDurable).not.toBe(srcDurable)
  })

  it('unknown durable targets stay guided: available list carries both id forms', () => {
    addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    const r = runTxn(opened, {
      ops: [{ op: 'deleteElement', target: { slide: 0, el: 'e_deadbeef' } }],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toMatch(/\(e_[0-9a-f]{8}\)/)
    const bad = runTxn(opened, {
      ops: [{ op: 'setHidden', target: { slide: 's_9' }, hidden: true }],
    })
    expect(bad.applied).toBe(false)
    expect(bad.failures![0]!.error).toContain('Available: [s_1]')
  })
})

describe('progressive creationId injection (foreign decks)', () => {
  const stripCreationId = (el: { anchor: { originalXml: string } }) => {
    el.anchor.originalXml = el.anchor.originalXml.replace(/<a:extLst>.*?<\/a:extLst>/, '')
  }

  it('an edited foreign element gets a creationId minted on the patch-save path', async () => {
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    stripCreationId(el)
    const cnvprForm = elementDurableId(el)!
    expect(cnvprForm).toMatch(/^e_\d+$/)

    // Edit it (dirtyFill) → its bytes are rewritten → identity hardened
    const r = runTxn(opened, {
      ops: [{ op: 'setFill', target: { slide: 0, el: cnvprForm }, fill: '#ABCDEF' }],
    })
    expect(r.applied).toBe(true)

    const reopened = await openPptx(await savePptx(opened))
    const surviving = reopened.deck.slides[0]!.elements.find(
      (x) => elementCNvPrId(x) === cnvprForm,
    )!
    expect(elementDurableId(surviving)).toMatch(/^e_[0-9a-f]{8}$/) // upgraded to the GUID form
    // The pre-upgrade cNvPr-form ref still resolves as an alias
    const r2 = runTxn(reopened, {
      ops: [{ op: 'setFill', target: { slide: 's_1', el: cnvprForm }, fill: '#0000FF' }],
    })
    expect(r2.applied).toBe(true)
  })

  it('minting is idempotent: repeated patch passes keep the same GUID', () => {
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    stripCreationId(el)
    runTxn(opened, { ops: [{ op: 'setFill', target: { slide: 0, el: el.id }, fill: '#ABCDEF' }] })
    const first = patchSlideXml(opened.deck.slides[0]!)
    const second = patchSlideXml(opened.deck.slides[0]!)
    expect(second).toBe(first)
    expect(first.match(/a16:creationId/g)?.length).toBeGreaterThan(0)
  })

  it('untouched foreign elements stay byte-identical on save (no gratuitous hardening)', async () => {
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    stripCreationId(el)
    const cnvprForm = elementDurableId(el)!
    const reopened = await openPptx(await savePptx(opened))
    const surviving = reopened.deck.slides[0]!.elements.find(
      (x) => elementCNvPrId(x) === cnvprForm,
    )!
    expect(surviving.anchor.originalXml).not.toContain('a16:creationId')
    expect(elementDurableId(surviving)).toBe(cnvprForm) // still the cNvPr form
  })

  it('minting merges into an existing extLst instead of adding a second one', () => {
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    // Foreign shape whose cNvPr already carries another extension
    el.anchor.originalXml = el.anchor.originalXml.replace(
      /<a:extLst>.*?<\/a:extLst>/,
      '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">' +
        '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>' +
        '</a:ext></a:extLst>',
    )
    runTxn(opened, { ops: [{ op: 'setFill', target: { slide: 0, el: el.id }, fill: '#ABCDEF' }] })
    const cnvpr = /<p:cNvPr\b[\s\S]*?<\/p:cNvPr>/.exec(el.anchor.originalXml)![0]
    expect(cnvpr.match(/<a:extLst>/g)!.length).toBe(1) // maxOccurs=1 holds
    expect(cnvpr).toContain('adec:decorative')
    expect(cnvpr).toContain('a16:creationId')
  })

  it('minting skips a hyperlink-nested extLst and appends cNvPr-level extLst', () => {
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    // Foreign shape: hlinkClick carries its OWN extLst; cNvPr has none of its own
    el.anchor.originalXml = el.anchor.originalXml.replace(
      /<a:extLst>.*?<\/a:extLst>/,
      '<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId9">' +
        '<a:extLst><a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}">' +
        '<ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/>' +
        '</a:ext></a:extLst></a:hlinkClick>',
    )
    runTxn(opened, { ops: [{ op: 'setFill', target: { slide: 0, el: el.id }, fill: '#ABCDEF' }] })
    const cnvpr = /<p:cNvPr\b[\s\S]*?<\/p:cNvPr>/.exec(el.anchor.originalXml)![0]
    const hlink = /<a:hlinkClick[\s\S]*?<\/a:hlinkClick>/.exec(cnvpr)![0]
    expect(hlink).not.toContain('a16:creationId') // GUID must not land in the hyperlink's extLst
    // cNvPr gains its own extLst as the LAST child, after the hyperlink
    expect(cnvpr).toMatch(/<\/a:hlinkClick><a:extLst><a:ext uri="\{FF2B5EF4-/)
  })

  it('the cNvPr alias resolves in group refs and flipElements after the upgrade', () => {
    const a = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#111111',
    })
    const b = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 1828800, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#222222',
    })
    stripCreationId(a)
    const aAlias = elementDurableId(a)!
    const g = runTxn(opened, {
      ops: [{ op: 'groupElements', target: { slide: 0 }, els: [a.id, b.id] }],
    })
    const grp = els().find((x) => x.id === g.records![0]!.created![0]!)!
    stripCreationId(grp as { anchor: { originalXml: string } })
    const grpAlias = elementDurableId(grp)!
    expect(grpAlias).toMatch(/^e_\d+$/)

    // Any op on the group upgrades it to a GUID mid-session…
    const t = runTxn(opened, {
      ops: [
        {
          op: 'setTransform',
          target: { slide: 0, el: grp.id },
          box: { x: 0, y: 0, cx: 2743200, cy: 914400 },
        },
      ],
    })
    expect(t.applied).toBe(true)
    const live = els().find((x) => x.id === grp.id)!
    expect(elementDurableId(live)).toMatch(/^e_[0-9a-f]{8}$/)
    // …and refs held from before still resolve: group field + flip els
    const r = runTxn(opened, {
      ops: [
        { op: 'setFill', target: { slide: 0, el: aAlias }, group: grpAlias, fill: '#ABCDEF' },
        { op: 'flipElements', target: { slide: 0 }, els: [grpAlias], axis: 'h' },
      ],
    })
    expect(r.applied).toBe(true)
  })
})

describe('identity continuity across the group boundary', () => {
  it("a child's durable id is the SAME before grouping, while grouped, and after ungroup", () => {
    const a = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#111111',
    })
    const b = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 1828800, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#222222',
    })
    const before = [elementDurableId(a)!, elementDurableId(b)!]
    expect(before[0]).toMatch(/^e_[0-9a-f]{8}$/)

    const g = runTxn(opened, {
      ops: [{ op: 'groupElements', target: { slide: 0 }, els: [a.id, b.id] }],
    })
    const grp = els().find((x) => x.id === g.records![0]!.created![0]!)!
    const children = (grp as unknown as { children: (typeof a)[] }).children
    const whileGrouped = children.map((c) => groupChildDurableId(grp, c))
    expect(new Set(whileGrouped)).toEqual(new Set(before))

    // Ops address the child by its pre-group durable id while grouped
    const groupDurable = elementDurableId(grp)!
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setFill',
          target: { slide: 0, el: before[0]! },
          group: groupDurable,
          fill: '#ABCDEF',
        },
      ],
    })
    expect(r.applied).toBe(true)

    const u = runTxn(opened, {
      ops: [{ op: 'ungroupElement', target: { slide: 0, el: grp.id } }],
    })
    expect(u.applied).toBe(true)
    const after = els().map((x) => elementDurableId(x))
    for (const id of before) expect(after).toContain(id)
  })
})

describe('durable refs on ops that call the engine by id', () => {
  it('setImageFill applies when addressed by a top-level durable id', () => {
    const PNG_1PX = Uint8Array.from(
      atob(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
      ),
      (c) => c.charCodeAt(0),
    )
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#123456',
    })
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setImageFill',
          target: { slide: 0, el: elementDurableId(el)! },
          source: { bytes: PNG_1PX, ext: 'png' },
        },
      ],
    })
    expect(r.applied).toBe(true)
    expect((r.records![0]!.after as { mediaPath: string }).mediaPath).toContain('ppt/media/')
  })
})

describe('durable addressing for groups and their children', () => {
  it('group + child ops resolve durable refs end-to-end', () => {
    const a = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 0, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#111111',
    })
    const b = addElement(opened.deck.slides[0]!, {
      kind: 'rect',
      offset: { x: 1828800, y: 0, cx: 914400, cy: 914400 },
      fillColor: '#222222',
    })
    const g = runTxn(opened, {
      ops: [{ op: 'groupElements', target: { slide: 0 }, els: [a.id, b.id] }],
    })
    const groupParseId = g.records![0]!.created![0]!
    const grp = els().find((x) => x.id === groupParseId)! as unknown as {
      children: Array<{ id: string; anchor: { originalXml: string } }>
    }
    const groupDurable = elementDurableId(els().find((x) => x.id === groupParseId)!)!
    const childDurable = elementDurableId(grp.children[0]! as never)!
    // Both the group ref and the child ref are durable ids
    const r = runTxn(opened, {
      ops: [
        {
          op: 'setFill',
          target: { slide: 's_1', el: childDurable },
          group: groupDurable,
          fill: '#ABCDEF',
        },
      ],
    })
    expect(r.applied).toBe(true)
    // A child's bytes live inside the parent group's XML — the fill lands there
    const grpEl = els().find((x) => x.id === groupParseId)!
    expect(grpEl.anchor.originalXml).toContain('ABCDEF')
  })
})
