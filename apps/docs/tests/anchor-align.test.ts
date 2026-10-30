import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  computeSectionedSlicesF2,
  type BlockBox,
  type SectionGeom,
  type SliceOutputs,
} from '../src/renderer/pagination'

const PIC =
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'

const pos = (v: number | string) =>
  typeof v === 'number' ? `<wp:posOffset>${v}</wp:posOffset>` : `<wp:align>${v}</wp:align>`

function anchor(
  relH: string,
  x: number | string,
  relV: string,
  y: number | string,
  wrap = '<wp:wrapSquare wrapText="bothSides"/>',
): string {
  return (
    '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${relH}">${pos(x)}</wp:positionH>` +
    `<wp:positionV relativeFrom="${relV}">${pos(y)}</wp:positionV>` +
    `<wp:extent cx="914400" cy="457200"/>${wrap}<wp:docPr id="1" name="pic"/>${PIC}</wp:anchor></w:drawing></w:r>`
  )
}

/** style attribute without whitespace: jsdom re-serialises declarations (and folds margins) */
const css = (el: Element | null) => (el?.getAttribute('style') ?? '').replace(/\s+/g, '')

async function open(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
}

// probe 2026-10-29 (fixture anchor-align-origin): a column-relative X ignores
// w:ind, keyword positions resolve in the margin / page band
describe('column-relative run pictures ignore the paragraph indent', () => {
  it('the host hands its start margin to a side float and the overlay', async () => {
    const editor = await open(
      `<w:p><w:pPr><w:ind w:left="720" w:firstLine="360"/></w:pPr><w:r><w:t>a</w:t></w:r>${anchor('column', 914400, 'paragraph', 0)}<w:r><w:t>b</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>a</w:t></w:r>${anchor('column', 914400, 'paragraph', 0, '<wp:wrapNone/>')}</w:p>`,
    )
    const [host, overlayHost] = Array.from(editor.view.dom.querySelectorAll<HTMLElement>('p'))
    expect(host.classList.contains('doc-anchor-host')).toBe(true)
    expect(host.style.getPropertyValue('--anchor-ind')).toBe('36pt')
    const float = host.querySelector<HTMLElement>('img.doc-inline-img--wrap-square-left')!
    expect(css(float)).toContain('calc(96.0px-var(--anchor-ind,0px)-var(--anchor-pad,0px))')
    expect(float.style.getPropertyValue('--obs-left')).toBe('calc(96.0px - var(--anchor-ind,0px))')
    expect(overlayHost.style.getPropertyValue('--anchor-ind')).toBe('36pt')
    const overlay = overlayHost.querySelector<HTMLElement>('.doc-inline-img-anchor > img')!
    expect(overlay.style.transform).toContain('translate(calc(96.0px - var(--anchor-ind,0px))')
    editor.destroy()
  })

  it('a page-relative X keeps its own undo channel and no indent term', async () => {
    const editor = await open(
      `<w:p><w:pPr><w:ind w:left="720"/></w:pPr><w:r><w:t>a</w:t></w:r>${anchor('page', 914400, 'paragraph', 0)}</w:p>`,
    )
    const host = editor.view.dom.querySelector<HTMLElement>('p')!
    expect(host.classList.contains('doc-anchor-host')).toBe(false)
    const float = host.querySelector<HTMLElement>('img.doc-inline-img--wrap-square-left')!
    expect(css(float)).toContain('calc(96.0px-var(--doc-margin-left,0px))')
    expect(css(float)).not.toContain('anchor-ind')
    editor.destroy()
  })
})

describe('wp:align keywords on in-flow picture blocks', () => {
  it('margin center / bottom wrapSquare centres the float and hands the band Y to the paginator', async () => {
    const editor = await open(`<w:p>${anchor('margin', 'center', 'margin', 'bottom')}</w:p>`)
    const block = editor.view.dom.querySelector<HTMLElement>('.doc-protected-image')!
    expect(block.dataset.tblpVy).toBe('0.0')
    expect(block.dataset.tblpVanchor).toBe('margin')
    expect(block.dataset.tblpVspec).toBe('bottom')
    expect(css(block)).toContain('calc((100%-96.0px)/2)')
    editor.destroy()
  })

  it('page right / top wrapTopAndBottom right-aligns the inner wrap in the page band', async () => {
    const editor = await open(
      `<w:p>${anchor('page', 'right', 'page', 'top', '<wp:wrapTopAndBottom/>')}</w:p>`,
    )
    const block = editor.view.dom.querySelector<HTMLElement>('.doc-protected-image')!
    expect(block.dataset.tblpVspec).toBe('top')
    expect(block.dataset.tblpVanchor).toBe('page')
    expect(css(block)).toContain('text-align:left')
    const wrap = block.querySelector<HTMLElement>('.doc-img-wrap')!
    expect(css(wrap)).toContain(
      'margin-left:calc(calc(calc(100%+var(--doc-margin-left,0px)+var(--doc-margin-right,0px))-96.0px)-var(--doc-margin-left,0px))',
    )
    editor.destroy()
  })

  it('a run picture with a centre keyword centres in the host box', async () => {
    const editor = await open(
      `<w:p><w:r><w:t>a</w:t></w:r>${anchor('margin', 'center', 'margin', 'bottom')}</w:p>`,
    )
    const img = editor.view.dom.querySelector<HTMLElement>('img.doc-inline-img--wrap-square-left')!
    expect(css(img)).toContain('calc(50%-48.0px-var(--anchor-ind,0px)/2-var(--anchor-pad,0px)/2)')
    // no page-height channel for a run picture: the bottom keyword stays at the anchor
    expect(img.dataset.pageRelV).toBeUndefined()
    editor.destroy()
  })
})

describe('in-flow wrapTopAndBottom band with a margin-relative Y', () => {
  const geoms: SectionGeom[] = [
    { contentHeight: 600, forceBreak: false, topPx: 60, pageHeightPx: 720, contentWidth: 400 },
  ]
  const block = (top: number, height: number, extra: Partial<BlockBox> = {}): BlockBox =>
    ({ top, height, ...extra }) as BlockBox

  it('lifts the band to the margin top over empty paragraphs and charges the shift once', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    const blocks = [
      block(0, 20, { emptyPara: true }),
      block(20, 20, { emptyPara: true }),
      block(40, 20, { emptyPara: true }),
      block(60, 36, { pageRelVyPx: 0, pageRelVAnchor: 'margin' }),
      block(96, 20),
    ]
    computeSectionedSlicesF2(blocks, geoms, 116, out)
    expect(out.floatVShifts).toEqual([{ blockTop: 60, dyPx: -60 }])
    // the next pass measures the applied lift in the flow: no double charge
    const again: SliceOutputs = { floatVShifts: [] }
    computeSectionedSlicesF2(
      [
        ...blocks.slice(0, 3),
        block(60, 36, { pageRelVyPx: 0, pageRelVAnchor: 'margin', pageRelVApplied: -60 }),
        block(36, 20),
      ],
      geoms,
      116,
      again,
    )
    expect(again.floatVShifts).toEqual([{ blockTop: 60, dyPx: -60 }])
  })

  it('does not lift over text above the anchor and pushes a lower target down', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    computeSectionedSlicesF2(
      [block(0, 20), block(20, 36, { pageRelVyPx: 0, pageRelVAnchor: 'margin' })],
      geoms,
      56,
      out,
    )
    expect(out.floatVShifts).toEqual([{ blockTop: 20, dyPx: 0 }])
    const down: SliceOutputs = { floatVShifts: [] }
    computeSectionedSlicesF2(
      [
        block(0, 20),
        block(20, 36, { pageRelVyPx: 0, pageRelVAnchor: 'margin', pageRelVSpec: 'bottom' }),
      ],
      geoms,
      56,
      down,
    )
    expect(down.floatVShifts).toEqual([{ blockTop: 20, dyPx: 544 }])
  })
})
