import { describe, expect, it } from 'vitest'
import { parseSlide } from '@genoffice/pptx-engine'
import { buildRenderSlide } from '../src/index'
import { setRedactExt } from '@genoffice/pptx-engine'

/**
 * The mark has to reach the render tree, because that is what the model's read
 * path walks. The canvas still draws the real words — a withheld span is a flag
 * on the node, not a substitution.
 */

const LABEL = '客户电话'
const SECRET = '13800138000'

const SP = (rPr: string, text: string) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="3000000" cy="800000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"/></p:spPr>' +
  `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r>${rPr}<a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`

const PIC = (nvPr: string) =>
  '<p:pic><p:nvPicPr>' +
  `<p:cNvPr id="3" name="clip"/>${nvPr}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
  '</p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  '<p:spPr><a:xfrm><a:off x="100000" y="1000000"/><a:ext cx="2000000" cy="1500000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"/></p:spPr></p:pic>'

const slideWith = (body: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`

const render = (body: string) =>
  buildRenderSlide(
    parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(body), ctx: {} }),
    { cx: 9144000, cy: 6858000 },
    { fitWidthPx: 1280 },
  )

const MARKED = setRedactExt('<a:rPr lang="en-US"/>', LABEL)
const PLAIN = '<a:rPr lang="en-US"/>'

describe('the mark reaches the render tree', () => {
  it('a withheld run arrives on its glyph runs, words intact', () => {
    const rs = render(SP(MARKED, `Call ${SECRET}`))
    const shape = rs.nodes.find((n) => n.type === 'shape' || n.type === 'text') as {
      text?: { lines: Array<{ runs: Array<{ text: string; redact?: string }> }> }
    }
    const runs = shape.text!.lines.flatMap((l) => l.runs)
    expect(runs.length).toBeGreaterThan(0)
    expect(runs.every((r) => r.redact === LABEL)).toBe(true)
    // the canvas still holds the real words — that is the whole point
    expect(runs.map((r) => r.text).join('')).toBe(`Call ${SECRET}`)
  })

  it('a plain run carries no mark', () => {
    const rs = render(SP(PLAIN, 'Call now'))
    const shape = rs.nodes.find((n) => n.type === 'shape' || n.type === 'text') as {
      text: { lines: Array<{ runs: Array<{ redact?: string }> }> }
    }
    for (const run of shape.text.lines.flatMap((l) => l.runs)) {
      expect(run.redact).toBeUndefined()
    }
  })

  it('only the withheld run is marked when a shape mixes both', () => {
    const body = SP(PLAIN, 'Public part ').replace(
      '</a:p>',
      `<a:r>${MARKED}<a:t>${SECRET}</a:t></a:r></a:p>`,
    )
    const rs = render(body)
    const shape = rs.nodes.find((n) => n.type === 'shape' || n.type === 'text') as {
      text: { lines: Array<{ runs: Array<{ text: string; redact?: string }> }> }
    }
    const runs = shape.text.lines.flatMap((l) => l.runs)
    const marked = runs.filter((r) => r.redact)
    const plain = runs.filter((r) => !r.redact)
    expect(marked.length).toBeGreaterThan(0)
    expect(plain.length).toBeGreaterThan(0)
    expect(marked.map((r) => r.text).join('')).toBe(SECRET)
    expect(plain.map((r) => r.text).join('')).toContain('Public part')
  })

  it('a withheld picture carries the mark onto the node', () => {
    // p:spPr is the one legal extension slot on a picture: p:nvPicPr admits
    // only cNvPr + cNvPicPr, and cNvPr takes attributes alone (which the user's
    // own alt text already uses)
    const markedSpPr = setRedactExt(
      '<p:spPr><a:xfrm><a:off x="100000" y="1000000"/><a:ext cx="2000000" cy="1500000"/></a:xfrm>' +
        '<a:prstGeom prst="rect"/></p:spPr>',
      LABEL,
    )
    const rs = render(SP(PLAIN, 'text') + PIC('').replace(/<p:spPr>.*?<\/p:spPr>/, markedSpPr))
    const pic = rs.nodes.find((n) => n.type === 'picture') as { redact?: string }
    expect(pic).toBeTruthy()
    expect(pic.redact).toBe(LABEL)
  })

  it('a plain picture carries no mark', () => {
    const rs = render(SP(PLAIN, 'text') + PIC(''))
    const pic = rs.nodes.find((n) => n.type === 'picture') as { redact?: string }
    expect(pic.redact).toBeUndefined()
  })
})
