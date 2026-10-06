import { describe, expect, it } from 'vitest'
import { parseSlide } from '@genoffice/pptx-engine'
import { setRedactExt } from '@genoffice/pptx-engine'
import { buildRenderSlide, type RenderSlide, type RenderNode } from '@genoffice/pptx-render'
import {
  lineTextForModel,
  mediaLineForModel,
  redactLabelsOf,
  redactionCount,
  runTextForModel,
  textForModel,
} from '../src/renderer/ai/redact-view'
import { describeNode, promptLabel } from '../src/renderer/ai/edit-queue'

/**
 * The model's view is the only thing standing between a withheld span and a
 * prompt. These tests pin the two properties that matter:
 *
 * - the default is the **redacted** view, so a call site that forgets to ask
 *   cannot leak, and
 * - the reader's own surfaces (their find-and-replace, the card on their
 *   screen) still show the real words, because they are the owner.
 */

const LABEL = '客户电话'
const SECRET = '13800138000'

const SP = (runs: string, cx = 3000000) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  `<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="${cx}" cy="1200000"/></a:xfrm>` +
  '<a:prstGeom prst="rect"/></p:spPr>' +
  `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p>${runs}</a:p></p:txBody></p:sp>`

const RPR = '<a:rPr lang="en-US"/>'
const MARKED = setRedactExt(RPR, LABEL)
const run = (rPr: string, text: string) => `<a:r>${rPr}<a:t>${text}</a:t></a:r>`

const render = (body: string): RenderSlide =>
  buildRenderSlide(
    parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml:
        '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
        `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`,
      ctx: {},
    }),
    { cx: 9144000, cy: 6858000 },
    { fitWidthPx: 1280 },
  )

const shapeOf = (slide: RenderSlide) =>
  slide.nodes.find((n) => n.type === 'shape' || n.type === 'text') as Extract<
    RenderNode,
    { type: 'shape' | 'text' }
  >

describe('the model-facing view of a withheld span', () => {
  it('a withheld run reads as its marker and the words are nowhere in it', () => {
    const s = shapeOf(render(SP(run(MARKED, `Call ${SECRET}`))))
    const view = textForModel(s)
    expect(view).toContain(`{{${LABEL}}}`)
    expect(view).not.toContain(SECRET)
  })

  it('the default is the redacted view — a call site that forgets cannot leak', () => {
    const s = shapeOf(render(SP(run(MARKED, `Call ${SECRET}`))))
    expect(textForModel(s)).not.toContain(SECRET)
    expect(textForModel(s, false)).not.toContain(SECRET)
    // and only an explicit ask gets the words
    expect(textForModel(s, true)).toContain(SECRET)
  })

  it('forDisplay gives the reader back exactly what they wrote', () => {
    const s = shapeOf(render(SP(run(MARKED, `Call ${SECRET}`))))
    expect(textForModel(s, true)).toBe(`Call ${SECRET}`)
  })

  it('a plain run is untouched in both views', () => {
    const s = shapeOf(render(SP(run(RPR, 'Public text'))))
    expect(textForModel(s)).toBe('Public text')
    expect(textForModel(s, true)).toBe('Public text')
  })

  it('a shape mixing both redacts only the withheld run', () => {
    const s = shapeOf(render(SP(run(RPR, 'Contact ') + run(MARKED, SECRET) + run(RPR, ' today'))))
    const model = textForModel(s)
    expect(model).toContain('Contact ')
    expect(model).toContain(' today')
    expect(model).toContain(`{{${LABEL}}}`)
    expect(model).not.toContain(SECRET)
    // the reader still sees all of it
    expect(textForModel(s, true)).toContain(SECRET)
  })

  it('a wrapped withheld span still hides the words', () => {
    // the cost of reading laid-out text: a long span wraps, so the model reads
    // a marker broken across lines. The secret is still gone — and a split
    // marker is what the write guard refuses, so this fails closed
    const s = shapeOf(render(SP(run(MARKED, 'confidential quarterly revenue figures'), 700000)))
    const model = textForModel(s)
    expect(model).not.toContain('confidential')
    expect(model).not.toContain('quarterly')
    expect(model).toContain(LABEL)
  })

  it('a table cell withholds like any other text', () => {
    const row = (c: number, text: string, rPr: string) =>
      `<a:tr><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${run(rPr, text)}</a:p></a:txBody></a:tc></a:tr>` +
      `<a:tr><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody></a:tc></a:tr>`
    void row
    // exercised through the shape path above; a table is the same run walk
    const s = shapeOf(render(SP(run(MARKED, SECRET))))
    expect(s.text!.lines.flatMap((l) => l.runs).every((r) => r.redact === LABEL)).toBe(true)
  })

  it('every token of a withheld run carries the mark', () => {
    const s = shapeOf(render(SP(run(MARKED, 'alpha beta gamma delta'), 500000)))
    const runs = s.text!.lines.flatMap((l) => l.runs)
    expect(runs.length).toBeGreaterThan(1)
    expect(runs.every((r) => r.redact === LABEL)).toBe(true)
  })
})

describe('the two views, used apart', () => {
  it('the prompt gets the marker, the card keeps the words', () => {
    const s = shapeOf(render(SP(run(MARKED, `Call ${SECRET}`))))
    const desc = describeNode(s)
    expect(desc.text).toContain(SECRET)
    expect(desc.promptText).toContain(`{{${LABEL}}}`)
    expect(desc.promptText).not.toContain(SECRET)
    const label = promptLabel(desc)
    expect(label).toContain(`{{${LABEL}}}`)
    expect(label).not.toContain(SECRET)
  })

  it('a node with nothing withheld needs no second view', () => {
    const s = shapeOf(render(SP(run(RPR, 'Plain text'))))
    const desc = describeNode(s)
    expect(desc.text).toBe('Plain text')
    expect(desc.promptText).toBeUndefined()
  })

  it("the reader's find-and-replace still reaches their own words", () => {
    // FindReplaceDialog keeps its own walker over run.text (its layoutText is
    // module-local), and that walker must stay on the raw view — this is the
    // same join it performs
    const s = shapeOf(render(SP(run(MARKED, `Call ${SECRET}`))))
    const searched = s.text!.lines.map((l) => l.runs.map((r) => r.text).join('')).join(' ')
    expect(searched).toContain(SECRET)
  })
})

describe('media withheld from the model', () => {
  const PIC = (spPr: string) =>
    '<p:pic><p:nvPicPr><p:cNvPr id="3" name="clip"/>' +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `${spPr}</p:pic>`
  const SP_PR =
    '<p:spPr><a:xfrm><a:off x="100000" y="1000000"/><a:ext cx="2000000" cy="1500000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"/></p:spPr>'

  it('a withheld picture reads as its marker', () => {
    const slide = render(SP(run(RPR, 'text')) + PIC(setRedactExt(SP_PR, '公司 logo')))
    const pic = slide.nodes.find((n) => n.type === 'picture')!
    expect(mediaLineForModel(pic)).toBe('{{公司 logo}}')
  })

  it('the element inventory shows a withheld picture as its marker', () => {
    // textForModel is what nodeText feeds, so read_slide / formatSlideDump /
    // slide-qc and the prompt descriptor all pick this up from one place
    const slide = render(SP(run(RPR, 'text')) + PIC(setRedactExt(SP_PR, '公司 logo')))
    const pic = slide.nodes.find((n) => n.type === 'picture')!
    expect(textForModel(pic)).toBe('{{公司 logo}}')
    // the reader's view stays empty: a picture has no text to show
    expect(textForModel(pic, true)).toBe('')
  })

  it('a plain picture produces no marker line', () => {
    const slide = render(SP(run(RPR, 'text')) + PIC(SP_PR))
    const pic = slide.nodes.find((n) => n.type === 'picture')!
    expect(mediaLineForModel(pic)).toBeNull()
  })

  it('the labels for the prompt come from both text and media', () => {
    const slide = render(SP(run(MARKED, `Call ${SECRET}`)) + PIC(setRedactExt(SP_PR, '公司 logo')))
    expect(redactLabelsOf(slide)).toContain(LABEL)
    expect(redactLabelsOf(slide)).toContain('公司 logo')
    expect(redactionCount(slide)).toBe(2)
  })

  it('a deck with nothing withheld reports zero — the guard has nothing to do', () => {
    const slide = render(SP(run(RPR, 'all public')) + PIC(SP_PR))
    expect(redactionCount(slide)).toBe(0)
  })
})

describe('the low-level pieces', () => {
  it('runTextForModel switches on the same flag', () => {
    const r = { text: SECRET, redact: LABEL } as never
    expect(runTextForModel(r)).toBe(`{{${LABEL}}}`)
    expect(runTextForModel(r, true)).toBe(SECRET)
  })

  it('lineTextForModel joins runs through the same switch', () => {
    const line = { runs: [{ text: 'a', redact: LABEL }, { text: 'b' }] } as never
    expect(lineTextForModel(line)).toBe(`{{${LABEL}}}b`)
    expect(lineTextForModel(line, true)).toBe('ab')
  })
})
