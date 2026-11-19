import { describe, it, expect, beforeEach } from 'vitest'
import {
  addElement,
  createBlankPptx,
  getSlideAnimations,
  openPptx,
  savePptx,
  type OpenedPptx,
  type TextElement,
} from '@genoffice/pptx-engine'
import { listSlideAnimations, runTxn } from '@genoffice/pptx-ops'

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

const slide0 = () => opened.deck.slides[0]!
const timingXml = () => /<p:timing>[\s\S]*<\/p:timing>/.exec(slide0().bodySuffix)?.[0] ?? ''

describe('addAnimation / removeAnimation / reorderAnimation', () => {
  it('appends effects with defaults, directions and aliases, then reads them back by element', () => {
    const r = runTxn(opened, {
      ops: [
        { op: 'addAnimation', target: { slide: 0, el: titleId }, effect: 'fade' },
        {
          op: 'addAnimation',
          target: { slide: 0, el: cardId },
          effect: 'fly',
          direction: 'left',
          trigger: 'afterPrevious',
          duration: 700,
          delay: 100,
        },
        { op: 'addAnimation', target: { slide: 0, el: titleId }, effect: 'flyOut', kind: 'exit' },
      ],
    })
    expect(r.applied).toBe(true)
    const list = listSlideAnimations(slide0())
    expect(list.map((a) => [a.seq, a.effect, a.trigger, a.durationMs, a.delayMs])).toEqual([
      [0, 'fade', 'onClick', 500, 0],
      [1, 'flyIn', 'afterPrev', 700, 100],
      [2, 'flyOut', 'onClick', 500, 0],
    ])
    expect(list[1]!.direction).toBe('left')
    expect(list[1]!.kind).toBe('entrance')
    expect(list[2]!.kind).toBe('exit')
    expect(list[1]!.el).toBeTruthy()
    // left fly-in: presetSubtype 8 and the x keyframe starts off the left edge
    expect(timingXml()).toContain('presetID="2" presetClass="entr" presetSubtype="8"')
    expect(timingXml()).toContain('0-#ppt_w/2')
  })

  it('rejects a kind that does not match the effect, bad directions and unknown effects with guidance', () => {
    const kind = runTxn(opened, {
      ops: [
        { op: 'addAnimation', target: { slide: 0, el: titleId }, effect: 'fade', kind: 'exit' },
      ],
    })
    expect(kind.applied).toBe(false)
    expect(kind.failures![0]!.error).toContain('fadeOut')
    const dir = runTxn(opened, {
      ops: [
        {
          op: 'addAnimation',
          target: { slide: 0, el: titleId },
          effect: 'fade',
          direction: 'left',
        },
      ],
    })
    expect(dir.failures![0]!.error).toContain('only applies to')
    const unknown = runTxn(opened, {
      ops: [{ op: 'addAnimation', target: { slide: 0, el: titleId }, effect: 'sparkle' }],
    })
    expect(unknown.failures![0]!.error).toContain('flyIn')
  })

  it('inserts after a position, removes by seq or element, reorders, all in one batch', () => {
    const r = runTxn(opened, {
      ops: [
        { op: 'addAnimation', target: { slide: 0, el: titleId }, effect: 'appear' },
        { op: 'addAnimation', target: { slide: 0, el: cardId }, effect: 'zoom' },
        { op: 'addAnimation', target: { slide: 0, el: cardId }, effect: 'spin', after: 0 },
        { op: 'reorderAnimation', target: { slide: 0 }, seq: 2, to: 0 },
        { op: 'removeAnimation', target: { slide: 0 }, seq: 1 },
      ],
    })
    expect(r.applied).toBe(true)
    expect(listSlideAnimations(slide0()).map((a) => a.effect)).toEqual(['zoom', 'spin'])
    const byEl = runTxn(opened, {
      ops: [{ op: 'removeAnimation', target: { slide: 0, el: cardId } }],
    })
    expect(byEl.applied).toBe(true)
    expect(byEl.records![0]!.after).toEqual({ removed: 2, count: 0 })
    expect(getSlideAnimations(slide0())).toEqual([])
  })

  it('seq out of range fails with the timeline range and applies nothing (atomic)', () => {
    runTxn(opened, {
      ops: [{ op: 'addAnimation', target: { slide: 0, el: titleId }, effect: 'fade' }],
    })
    const r = runTxn(opened, {
      ops: [
        { op: 'addAnimation', target: { slide: 0, el: cardId }, effect: 'fade' },
        { op: 'removeAnimation', target: { slide: 0 }, seq: 5 },
      ],
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('0-1')
    expect(getSlideAnimations(slide0())).toHaveLength(1)
  })

  it('presetXml passes a PowerPoint-exported effect through save/reopen byte-for-byte', async () => {
    const spid = /<p:cNvPr\s[^>]*\bid="(\d+)"/.exec(
      slide0().elements.find((e) => e.id === cardId)!.anchor.originalXml,
    )![1]
    // Random Bars (presetID 5) has no modeled writer
    const presetXml =
      '<p:par><p:cTn id="99" presetID="5" presetClass="entr" presetSubtype="10" fill="hold" grpId="0" nodeType="clickEffect">' +
      '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
      `<p:set><p:cBhvr><p:cTn id="98" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
      `<p:animEffect transition="in" filter="randombar(horizontal)"><p:cBhvr><p:cTn id="97" dur="500"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect>` +
      '</p:childTnLst></p:cTn></p:par>'
    const r = runTxn(opened, {
      ops: [
        { op: 'addAnimation', target: { slide: 0, el: titleId }, effect: 'fade' },
        { op: 'addAnimation', target: { slide: 0, el: cardId }, presetXml, trigger: 'withPrev' },
      ],
    })
    expect(r.applied).toBe(true)
    const wrongTarget = runTxn(opened, {
      ops: [{ op: 'addAnimation', target: { slide: 0, el: titleId }, presetXml }],
    })
    expect(wrongTarget.failures![0]!.error).toContain('spid')

    const reopened = await openPptx(await savePptx(opened))
    const list = listSlideAnimations(reopened.deck.slides[0]!)
    expect(list).toHaveLength(2)
    expect(list[1]!.custom).toBe(true)
    expect(list[1]!.preset).toEqual({ id: 5, cls: 'entr', sub: 10 })
    expect(list[1]!.trigger).toBe('withPrev')
    const xml = /<p:timing>[\s\S]*<\/p:timing>/.exec(reopened.deck.slides[0]!.bodySuffix)![0]
    expect(xml).toContain('filter="randombar(horizontal)"')
    expect(xml).toContain('nodeType="withEffect"')
    // ids renumbered into one unique space
    const ids = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)

    // editing another animation keeps the custom one intact
    const r2 = runTxn(reopened, {
      ops: [{ op: 'removeAnimation', target: { slide: 0 }, seq: 0 }],
    })
    expect(r2.applied).toBe(true)
    const after = /<p:timing>[\s\S]*<\/p:timing>/.exec(reopened.deck.slides[0]!.bodySuffix)![0]
    expect(after).toContain('filter="randombar(horizontal)"')
    expect(after).not.toContain('filter="fade"')
  })
})

describe('insertEquation', () => {
  const equationBlocks = (xml: string) =>
    [...xml.matchAll(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/g)].filter((m) =>
      m[0].includes('<a14:m>'),
    )

  it('appends an a14:m paragraph with a text fallback and survives save/reopen', async () => {
    const r = runTxn(opened, {
      ops: [{ op: 'insertEquation', target: { slide: 0, el: titleId }, latex: 'E = mc^2' }],
    })
    expect(r.applied).toBe(true)
    const el = slide0().elements.find((e) => e.id === titleId) as TextElement
    expect(el.text!.paragraphs).toHaveLength(2)
    expect(el.text!.paragraphs[1]!.runs[0]!.text).toBe('E = mc^2')

    const reopened = await openPptx(await savePptx(opened))
    const slideXml = reopened.archive.readText(reopened.deck.slides[0]!.path)!
    const blocks = equationBlocks(slideXml)
    expect(blocks).toHaveLength(1)
    expect(blocks[0]![0]).toContain('<m:oMathPara')
    expect(blocks[0]![0]).toContain('<m:sSup>')
    expect(blocks[0]![0]).toMatch(/<mc:Fallback><a:r>.*<a:t>E = mc\^2<\/a:t><\/a:r><\/mc:Fallback>/)
    // the block sits inside the paragraph, after the pPr slot and before endParaRPr
    expect(slideXml).toMatch(/<a:p><a:pPr algn="ctr"\/><mc:AlternateContent/)

    const again = reopened.deck.slides[0]!.elements.find((e) => e.type === 'text') as TextElement
    const runs = again.text!.paragraphs.flatMap((p) => p.runs)
    expect(runs.some((run) => run.rawXml?.includes('<a14:m>') && run.text === 'E = mc^2')).toBe(
      true,
    )

    // a later text edit of another paragraph keeps the equation bytes
    const r2 = runTxn(reopened, {
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el: again.id },
          paragraphs: [
            { runs: [{ text: 'Renamed' }] },
            { runs: [{ text: 'E = mc^2', srcRun: 0 }], srcPara: 1 },
          ],
        },
      ],
    })
    expect(r2.applied).toBe(true)
    const saved = (await openPptx(await savePptx(reopened))).archive.readText(
      reopened.deck.slides[0]!.path,
    )!
    expect(equationBlocks(saved)).toHaveLength(1)
    expect(saved).toContain('Renamed')
  })

  it('creates a text box when given a box, refuses unsupported LaTeX with a hint', () => {
    const r = runTxn(opened, {
      ops: [
        {
          op: 'insertEquation',
          target: { slide: 0 },
          box: { x: 914400, y: 914400, cx: 3657600, cy: 914400 },
          latex: '\\frac{a}{b}',
        },
      ],
    })
    expect(r.applied).toBe(true)
    const created = r.records![0]!.created![0]!
    const el = slide0().elements.find((e) => e.id === created) as TextElement
    expect(el.type).toBe('text')
    expect(el.text!.paragraphs[0]!.runs[0]!.rawXml).toContain('<m:f>')
    expect(el.text!.paragraphs[0]!.runs[0]!.text).toBe('a/b')

    const bad = runTxn(opened, {
      ops: [{ op: 'insertEquation', target: { slide: 0, el: titleId }, latex: '\\unknowncmd{x' }],
    })
    expect(bad.applied).toBe(false)
    expect(bad.failures![0]!.error).toContain('Supported')
    const noTarget = runTxn(opened, {
      ops: [{ op: 'insertEquation', target: { slide: 0 }, latex: 'x' }],
    })
    expect(noTarget.failures![0]!.error).toContain('"box"')
  })

  it('editing the equation run text turns it back into a plain run; a field-only paragraph is kept', async () => {
    runTxn(opened, {
      ops: [{ op: 'insertEquation', target: { slide: 0, el: titleId }, latex: 'a^2' }],
    })
    const reopened = await openPptx(await savePptx(opened))
    const el = reopened.deck.slides[0]!.elements.find((e) => e.type === 'text') as TextElement
    const r = runTxn(reopened, {
      ops: [
        {
          op: 'setText',
          target: { slide: 0, el: el.id },
          paragraphs: [
            { runs: [{ text: 'Title', srcRun: 0 }], srcPara: 0 },
            { runs: [{ text: 'plain now', srcRun: 0 }], srcPara: 1 },
          ],
        },
      ],
    })
    expect(r.applied).toBe(true)
    const xml = (await openPptx(await savePptx(reopened))).archive.readText(
      reopened.deck.slides[0]!.path,
    )!
    expect(xml).not.toContain('<a14:m>')
    expect(xml).toContain('<a:t>plain now</a:t>')

    const field = addElement(slide0(), {
      kind: 'textbox',
      offset: { x: 0, y: 1500000, cx: 914400, cy: 457200 },
      paragraphs: [{ runs: [{ text: '', field: 'slidenum' }] }],
    })
    const r2 = runTxn(opened, {
      ops: [{ op: 'insertEquation', target: { slide: 0, el: field.id }, latex: 'b' }],
    })
    expect(r2.applied).toBe(true)
    const paras = (slide0().elements.find((e) => e.id === field.id) as TextElement).text!.paragraphs
    expect(paras).toHaveLength(2)
    expect(paras[0]!.runs[0]!.field).toBe('slidenum')
    expect(paras[1]!.runs[0]!.rawXml).toContain('<a14:m>')
  })
})
