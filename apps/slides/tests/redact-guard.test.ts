import { describe, expect, it } from 'vitest'
import { parseSlide, setRedactExt } from '@genoffice/pptx-engine'
import { buildRenderSlide, type RenderSlide, type RenderNode } from '@genoffice/pptx-render'
import { redactGuardFor } from '../src/renderer/ai/redact-guard'

/**
 * The guard is a **damage** guard, not a secrecy guard: the prompt path already
 * keeps the words from the model, and what is left is the model quietly
 * rewriting or deleting them.
 *
 * The pairs below matter more than the individual cases. A guard that refuses
 * everything is safe and useless; each "allowed" case is the other half of a
 * "refused" one and says which distinction is being drawn.
 */

const LABEL = '客户电话'
const SECRET = '13800138000'
const MARKED = setRedactExt('<a:rPr lang="en-US"/>', LABEL)
const RPR = '<a:rPr lang="en-US"/>'

const SP = (id: number, cx: string, runs: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="${cx}" cy="900000"/></a:xfrm>` +
  '<a:prstGeom prst="rect"/></p:spPr>' +
  `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p>${runs}</a:p></p:txBody></p:sp>`

const run = (rPr: string, text: string) => `<a:r>${rPr}<a:t>${text}</a:t></a:r>`

const PIC = (id: number, spPr: string) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p${id}"/>` +
  '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  spPr +
  '</p:pic>'

const PIC_SP_PR =
  '<p:spPr><a:xfrm><a:off x="100000" y="1000000"/><a:ext cx="2000000" cy="1500000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"/></p:spPr>'

const GRP = (children: string) =>
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="3000000"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="4000000" cy="3000000"/></a:xfrm></p:grpSpPr>' +
  children +
  '</p:grpSp>'

const deck = (...pages: string[]): RenderSlide[] =>
  pages.map((body, i) =>
    buildRenderSlide(
      parseSlide({
        path: `ppt/slides/slide${i + 1}.xml`,
        slideXml:
          '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
          `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`,
        ctx: {},
      }),
      { cx: 9144000, cy: 6858000 },
      { fitWidthPx: 1280 },
    ),
  )

const withSecret = deck(SP(2, '3000000', run(MARKED, `Call ${SECRET}`)))
const plain = deck(SP(2, '3000000', run(RPR, 'Call now')))
const idOf = (slide: RenderSlide, pred: (n: RenderNode) => boolean) => {
  const walk = (nodes: RenderNode[]): string | null => {
    for (const n of nodes) {
      if (pred(n)) return n.id
      if (n.type === 'group') {
        const hit = walk(n.children)
        if (hit) return hit
      }
    }
    return null
  }
  return walk(slide.nodes)
}
const isText = (n: RenderNode) => n.type === 'shape' || n.type === 'text'

describe('an op that would rewrite withheld words', () => {
  it('refuses a findReplace whose needle is inside a withheld run', () => {
    const hit = redactGuardFor({ op: 'findReplace', find: SECRET }, withSecret)
    expect(hit).not.toBeNull()
    expect(hit!.reason).toContain(LABEL)
  })

  it('refuses a findReplace that eats part of a withheld run', () => {
    // the model never saw the digits, so a partial hit is not an intent it can have had
    expect(redactGuardFor({ op: 'findReplace', find: '1380013' }, withSecret)).not.toBeNull()
  })

  it('refuses a findReplace scoped to a different page, because it scans the deck', () => {
    const two = deck(
      SP(2, '3000000', run(RPR, 'nothing')),
      SP(3, '3000000', run(MARKED, `Call ${SECRET}`)),
    )
    expect(redactGuardFor({ op: 'findReplace', find: SECRET, slideIndex: 0 }, two)).not.toBeNull()
  })

  it('refuses a wholesale text replacement on a withheld shape', () => {
    const id = idOf(withSecret[0], isText)!
    const hit = redactGuardFor({ op: 'setText', elementId: id, slideIndex: 0 }, withSecret)
    expect(hit).not.toBeNull()
    expect(hit!.reason).toContain(LABEL)
  })

  it('refuses deleting a withheld element', () => {
    const id = idOf(withSecret[0], isText)!
    expect(
      redactGuardFor({ op: 'deleteElement', elementId: id, slideIndex: 0 }, withSecret),
    ).not.toBeNull()
  })

  it('refuses deleting a page that carries a withheld span', () => {
    const hit = redactGuardFor({ op: 'deleteSlide', slideIndex: 0 }, withSecret)
    expect(hit).not.toBeNull()
    expect(hit!.reason).toContain('page')
  })

  it('refuses replacing a withheld picture', () => {
    const slide = deck(
      SP(2, '3000000', run(RPR, 't')) + PIC(3, setRedactExt(PIC_SP_PR, '公司 logo')),
    )
    const id = idOf(slide[0], (n) => n.type === 'picture')!
    expect(
      redactGuardFor({ op: 'replacePicture', elementId: id, slideIndex: 0 }, slide),
    ).not.toBeNull()
  })

  it('refuses ungrouping a group that holds a withheld child', () => {
    const grouped = deck(GRP(SP(2, '3000000', run(MARKED, `Call ${SECRET}`))))
    const groupId = idOf(grouped[0], (n) => n.type === 'group')!
    expect(
      redactGuardFor({ op: 'ungroupElement', elementId: groupId, slideIndex: 0 }, grouped),
    ).not.toBeNull()
  })
})

describe('the same ops on something ordinary', () => {
  it('allows a findReplace that touches nothing withheld', () => {
    // note the whole run "Call <secret>" is withheld here, so "Call" IS inside
    // it and refusing is right — the needle has to miss the span entirely
    expect(redactGuardFor({ op: 'findReplace', find: 'Call' }, withSecret)).not.toBeNull()
    expect(redactGuardFor({ op: 'findReplace', find: 'quarterly' }, withSecret)).toBeNull()
  })

  it('allows replacing a plain shape', () => {
    const id = idOf(plain[0], isText)!
    expect(redactGuardFor({ op: 'setText', elementId: id, slideIndex: 0 }, plain)).toBeNull()
  })

  it('allows deleting a plain element and a page with nothing withheld', () => {
    const id = idOf(plain[0], isText)!
    expect(redactGuardFor({ op: 'deleteElement', elementId: id, slideIndex: 0 }, plain)).toBeNull()
    expect(redactGuardFor({ op: 'deleteSlide', slideIndex: 0 }, plain)).toBeNull()
  })

  it('leaves a deck with nothing withheld entirely alone', () => {
    const id = idOf(plain[0], isText)!
    for (const op of [
      { op: 'findReplace', find: 'anything' },
      { op: 'setText', elementId: id, slideIndex: 0 },
      { op: 'deleteSlide', slideIndex: 0 },
      { op: 'replacePicture', elementId: id, slideIndex: 0 },
    ]) {
      expect(redactGuardFor(op, plain), JSON.stringify(op)).toBeNull()
    }
  })
})

describe('the distinction that keeps the feature usable', () => {
  it('allows moving and resizing a withheld shape — geometry is not its content', () => {
    const id = idOf(withSecret[0], isText)!
    for (const op of [
      { op: 'setTransform', elementId: id, slideIndex: 0 },
      { op: 'setFill', elementId: id, slideIndex: 0 },
      { op: 'setStroke', elementId: id, slideIndex: 0 },
      { op: 'setFont', elementId: id, slideIndex: 0 },
      { op: 'setParagraphFormat', elementId: id, slideIndex: 0 },
      { op: 'reorderElement', elementId: id, slideIndex: 0 },
      { op: 'setNotes', elementId: id, slideIndex: 0 },
      { op: 'setPictureOpacity', elementId: id, slideIndex: 0 },
    ]) {
      expect(redactGuardFor(op, withSecret), `${op.op} should be allowed`).toBeNull()
    }
  })

  it('allows a findReplace of different words in the same shape', () => {
    // the withheld span and the sentence around it are separate runs
    const mixed = deck(
      SP(2, '4000000', run(RPR, 'Public ') + run(MARKED, SECRET) + run(RPR, ' today')),
    )
    expect(redactGuardFor({ op: 'findReplace', find: 'Public' }, mixed)).toBeNull()
    expect(redactGuardFor({ op: 'findReplace', find: SECRET }, mixed)).not.toBeNull()
  })

  it('honours matchCase when looking for the needle', () => {
    // a case-insensitive search is the default, so a differently-cased needle
    // must still be caught. A phone number cannot show this — it has no case —
    // so this fixture is a key with letters in it.
    // The needle has to sit inside one laid-out run: layout breaks a run on a
    // case transition, so `sk-ABC123secret` becomes `sk-` + `ABC123secret`.
    const keyed = deck(SP(2, '4000000', run(setRedactExt(RPR, 'api key'), 'sk-ABC123secret')))
    expect(redactGuardFor({ op: 'findReplace', find: 'abc123secret' }, keyed)).not.toBeNull()
    expect(
      redactGuardFor({ op: 'findReplace', find: 'ABC123secret', matchCase: true }, keyed),
    ).not.toBeNull()
    // a case-SENSITIVE search for the wrong case genuinely does not match, so
    // refusing it would be a false positive
    expect(
      redactGuardFor({ op: 'findReplace', find: 'abc123secret', matchCase: true }, keyed),
    ).toBeNull()
  })

  it('a needle straddling a run boundary is left to the op, which cannot match it either', () => {
    // Layout breaks a run on a case transition, and the executor documents that
    // a match spanning two differently formatted runs is not replaced. The guard
    // and the op therefore share a granularity, and neither damages a span for
    // a needle that crosses one — worth pinning, since a guard that DID refuse
    // here would be a false positive on every multi-token word.
    const keyed = deck(SP(2, '4000000', run(setRedactExt(RPR, 'api key'), 'sk-ABC123secret')))
    expect(redactGuardFor({ op: 'findReplace', find: 'sk-abc123secret' }, keyed)).toBeNull()
  })

  it('an op that is not object-shaped is not the guard’s business', () => {
    expect(redactGuardFor({}, withSecret)).toBeNull()
    expect(redactGuardFor({ op: 'setBackground', slideIndex: 0 }, withSecret)).toBeNull()
  })
})
