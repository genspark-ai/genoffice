import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  computeSectionedSlicesF2,
  type BlockBox,
  type SectionGeom,
  type SliceOutputs,
} from '../src/renderer/pagination'

const cell = { type: 'docTableCell', content: [{ type: 'docParagraph' }] }
const row = { type: 'docTableRow', content: [cell] }
const editor = new Editor({ element: document.createElement('div'), extensions: editorExtensions })

function tableAttrs(attrs: Record<string, unknown>): Record<string, string> {
  const table = editor.schema.nodeFromJSON({ type: 'docTable', attrs, content: [row] })
  return (editor.schema.nodes.docTable.spec.toDOM!(table) as [string, Record<string, string>])[1]
}

const block = (top: number, height: number, extra?: Partial<BlockBox>): BlockBox => ({
  top,
  height,
  ...extra,
})

describe('w:tblpXSpec / w:tblpYSpec keyword anchors', () => {
  it('a keyword Y on a margin anchor emits the spec instead of a fixed target', () => {
    const attrs = tableAttrs({
      tblFloat: 'left',
      tblFloatSource: 'left',
      tblFloatHorzAnchor: 'margin',
      tblFloatVertAnchor: 'margin',
      tblFloatXSpec: 'center',
      tblFloatYSpec: 'bottom',
    })
    expect(attrs['data-tblp-vy']).toBe('0.0')
    expect(attrs['data-tblp-vanchor']).toBe('margin')
    expect(attrs['data-tblp-vspec']).toBe('bottom')
    expect(attrs.style).toContain('margin-top:var(--tblp-dy,0px)')
  })

  it('a page-anchored X of zero reaches into the left margin', () => {
    const attrs = tableAttrs({
      tblFloat: 'left',
      tblFloatSource: 'left',
      tblFloatHorzAnchor: 'page',
      tblFloatVertAnchor: 'page',
      tblFloatXTwips: 0,
      tblFloatYSpec: 'top',
      tblFloatWidthPx: 264,
    })
    expect(attrs.style).toContain(
      'margin-left:max(calc(0px - var(--doc-margin-left,0px)),min(calc(0.0px - var(--doc-margin-left,0px)),max(0px,calc(var(--doc-content-w,100%) - 264.0px))))',
    )
    // a page-anchored X inside the content box with a table wider than the
    // content stays at the content edge (the cap must not pull it into the margin)
    const wide = tableAttrs({
      tblFloat: 'left',
      tblFloatSource: 'left',
      tblFloatHorzAnchor: 'page',
      tblFloatVertAnchor: 'text',
      tblFloatXTwips: 4000,
      tblFloatWidthPx: 700,
    })
    expect(wide.style).toContain(
      'margin-left:max(calc(0px - var(--doc-margin-left,0px)),min(calc(266.7px - var(--doc-margin-left,0px)),max(0px,calc(var(--doc-content-w,100%) - 700.0px))))',
    )
    const marginAnchored = tableAttrs({
      tblFloat: 'left',
      tblFloatSource: 'left',
      tblFloatHorzAnchor: 'margin',
      tblFloatVertAnchor: 'text',
      tblFloatXTwips: 0,
    })
    expect(marginAnchored.style ?? '').not.toContain('margin-left')
  })

  it('centers a fixed-width float in the content box', () => {
    const attrs = tableAttrs({
      tblFloat: 'left',
      tblFloatSource: 'left',
      tblFloatHorzAnchor: 'margin',
      tblFloatVertAnchor: 'text',
      tblFloatXSpec: 'center',
      tblFloatWidthPx: 200,
    })
    expect(attrs.style).toContain(
      'margin-left:max(0px,calc((var(--doc-content-w,100%) - 200.0px) / 2))',
    )
  })
})

describe('keyword-anchored float shifts in the slice engine', () => {
  const geoms: SectionGeom[] = [
    { contentHeight: 600, forceBreak: false, topPx: 60, pageHeightPx: 720, contentWidth: 400 },
  ]

  it('bottom aligns the table bottom with the content bottom', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    const blocks = [
      block(0, 40),
      block(40, 60, {
        floated: true,
        pageRelVyPx: 0,
        pageRelVAnchor: 'margin',
        pageRelVSpec: 'bottom',
      }),
    ]
    computeSectionedSlicesF2(blocks, geoms, 600, out)
    expect(out.floatVShifts).toEqual([{ blockTop: 40, dyPx: 500 }])
  })

  it('a page-anchored top keyword on a blank page hangs the table into the top margin', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    const blocks = [
      block(0, 100, { floated: true, pageRelVyPx: 0, pageRelVAnchor: 'page', pageRelVSpec: 'top' }),
      block(0, 40),
    ]
    computeSectionedSlicesF2(blocks, geoms, 600, out)
    expect(out.floatVShifts).toEqual([{ blockTop: 0, dyPx: -60 }])
  })

  it('a float stacked under an earlier float measures its shift from where CSS put it', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    const blocks = [
      block(0, 100, { floated: true, pageRelVyPx: 0, pageRelVAnchor: 'page', pageRelVSpec: 'top' }),
      block(0, 40),
      block(100, 60, {
        floated: true,
        pageRelVyPx: 0,
        pageRelVAnchor: 'margin',
        pageRelVSpec: 'bottom',
      }),
    ]
    computeSectionedSlicesF2(blocks, geoms, 600, out)
    expect(out.floatVShifts).toEqual([
      { blockTop: 0, dyPx: -60 },
      { blockTop: 100, dyPx: 440 },
    ])
  })

  it('never lifts a float above content already on the page', () => {
    const out: SliceOutputs = { floatVShifts: [] }
    const blocks = [
      block(0, 200),
      block(200, 100, {
        floated: true,
        pageRelVyPx: 0,
        pageRelVAnchor: 'page',
        pageRelVSpec: 'top',
      }),
    ]
    computeSectionedSlicesF2(blocks, geoms, 600, out)
    expect(out.floatVShifts).toEqual([{ blockTop: 200, dyPx: 0 }])
  })
})

describe('legacy cell-margin hang of a float', () => {
  it('only a margin-anchored float without an offset hangs by its indent', () => {
    const base = { tblFloat: 'left', tblFloatSource: 'left', tblFloatVertAnchor: 'text' }
    const hung = tableAttrs({ ...base, tblFloatHorzAnchor: 'margin', indentTwips: -108 })
    expect(hung.style).toContain('margin-left:-7.2px')
    const textAnchored = tableAttrs({ ...base, tblFloatHorzAnchor: 'text', indentTwips: -108 })
    expect(textAnchored.style ?? '').not.toContain('margin-left')
    const unanchored = tableAttrs({ ...base, indentTwips: -108 })
    expect(unanchored.style ?? '').not.toContain('margin-left')
    const offset = tableAttrs({
      ...base,
      tblFloatHorzAnchor: 'margin',
      tblFloatXTwips: 300,
      indentTwips: -108,
    })
    expect(offset.style).not.toContain('-7.2px')
  })
})
