import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import type { PmNode } from '../src/renderer/editor/convert'
import {
  buildBlockPlan,
  planBreakFlags,
  type BlockPlan,
  type PlanProps,
} from '../src/renderer/pagination-plan'

const PROPS: PlanProps = {
  paraEffective: () => undefined,
  additiveSpacing: false,
  autoSpacingPt: 14,
  showDeleted: false,
}

const text = (t: string) => ({ type: 'text', text: t })
const pageBr = { type: 'hardBreak', attrs: { pageBreak: true } }
const colBr = { type: 'hardBreak', attrs: { colBreak: true } }
const para = (content: unknown[], attrs: Record<string, unknown> = {}) => ({
  type: 'docParagraph',
  attrs: { docxIndex: null, styleId: null, ...attrs },
  ...(content.length ? { content } : {}),
})

function plansOf(blocks: unknown[], props: PlanProps = PROPS): BlockPlan[] {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: blocks } as unknown as PmNode as never,
  })
  const out = [...buildBlockPlan(editor.state.doc, props).values()]
  editor.destroy()
  return out
}

describe('buildBlockPlan: P5 break classification', () => {
  it('a trailing page break pushes the next block, a leading one leads its block', () => {
    const [a, b, c] = plansOf([
      para([text('one'), pageBr]),
      para([pageBr, text('two')]),
      para([text('three')]),
    ])
    expect(planBreakFlags(a)).toMatchObject({ breakAfter: true, innerBreakCount: 0 })
    // a single leading break line cuts like a mid-paragraph break
    expect(planBreakFlags(b)).toMatchObject({ innerBreakCount: 1, leadingBreakCut: true })
    expect(planBreakFlags(b).breakBefore).toBeUndefined()
    expect(planBreakFlags(c)).toEqual({ innerBreakCount: 0 })
  })

  it('two leading breaks: the last one cuts, the first is an extra sheet', () => {
    const [p] = plansOf([para([pageBr, pageBr, text('x')])])
    expect(planBreakFlags(p)).toMatchObject({
      extraBreaksBefore: 1,
      innerBreakCount: 1,
      leadingBreakCut: true,
    })
  })

  it('a break with text on both sides is an inner break', () => {
    const [p] = plansOf([para([text('a'), pageBr, text('b'), pageBr, text('c'), pageBr])])
    expect(planBreakFlags(p)).toMatchObject({ innerBreakCount: 2, breakAfter: true })
    expect(p.innerBreakIdx).toEqual([0, 1])
  })

  it('a break-only paragraph is marked and also breaks after', () => {
    const [p] = plansOf([para([pageBr])])
    expect(p.breakOnly).toBe(true)
    expect(planBreakFlags(p)).toMatchObject({ breakAfter: true, innerBreakCount: 0 })
    const [withImg] = plansOf([para([pageBr, { type: 'docInlineImage', attrs: { src: 'x.png' } }])])
    expect(withImg.breakOnly).toBe(false)
  })

  it('a footnote reference counts as text (its number is rendered)', () => {
    const [p] = plansOf([
      para([{ type: 'docNoteRef', attrs: { id: '1', kind: 'footnote', num: 1 } }, pageBr]),
    ])
    expect(p.hasText).toBe(true)
    expect(p.breakOnly).toBe(false)
  })

  it('a deleted break only breaks in the Original view', () => {
    const doc = [
      para([text('a'), { ...pageBr, marks: [{ type: 'del', attrs: { author: 'x' } }] }, text('b')]),
    ]
    expect(planBreakFlags(plansOf(doc)[0]).innerBreakCount).toBe(0)
    expect(planBreakFlags(plansOf(doc, { ...PROPS, showDeleted: true })[0]).innerBreakCount).toBe(1)
  })

  it('column breaks: leading moves the block, otherwise it breaks after', () => {
    const [lead, tail] = plansOf([para([colBr, text('a')]), para([text('b'), colBr])])
    expect(planBreakFlags(lead)).toMatchObject({ colBreakBefore: true })
    expect(planBreakFlags(tail)).toMatchObject({ colBreakAfter: true })
  })

  it('pageBreakBefore: direct from the node, inherited from the resolved style', () => {
    const eff = {
      pageBreakBefore: true,
      spacing: { before: 0, after: 0, beforeAuto: false, afterAuto: false },
      source: { pageBreakBefore: 'paraStyle' },
    } as unknown as NonNullable<ReturnType<PlanProps['paraEffective']>>
    const [direct, styled] = plansOf(
      [para([text('a')], { pageBreakBefore: true }), para([text('b')], { docxIndex: 7 })],
      { ...PROPS, paraEffective: (i) => (i === 7 ? eff : undefined) },
    )
    expect(direct.pageBreakBeforeDirect).toBe(true)
    expect(planBreakFlags(direct).breakBefore).toBe(true)
    expect(styled.pageBreakBeforeDirect).toBe(false)
    expect(styled.pageBreakBefore).toBe(true)
    expect(styled.dropAtPageTop).toBe(false)
  })

  it('a protected page-break field breaks after its block', () => {
    const [p] = plansOf([
      {
        type: 'docProtected',
        attrs: { blockType: 'passthrough', fieldDisplay: { kind: 'pageBreak' } },
      },
    ])
    expect(planBreakFlags(p)).toMatchObject({ breakAfter: true, innerBreakCount: 0 })
    expect(p.breakOnly).toBe(false)
  })

  it('only an in-line picture host owns its space before as the outer margin', () => {
    const host = (attrs: Record<string, unknown>) => ({
      type: 'docProtected',
      attrs: {
        blockType: 'image',
        imageDataUrl: 'data:,',
        imageParagraphSpaceBefore: 120,
        ...attrs,
      },
    })
    const [inline, anchored, boxes] = plansOf([
      host({}),
      host({ imageWrap: 'square' }),
      host({ blockType: 'passthrough', textboxes: [{ floating: true, paras: [] }] }),
    ])
    expect(inline).toMatchObject({ spaceBeforePt: 6 })
    expect(inline.spaceBeforeInner).toBeUndefined()
    expect(anchored).toMatchObject({ spaceBeforePt: 6, spaceBeforeInner: true })
    expect(boxes).toMatchObject({ spaceBeforePt: 6, spaceBeforeInner: true })
  })

  it("a host with auto space before drops it on the document's first block", () => {
    const host = {
      type: 'docProtected',
      attrs: {
        blockType: 'passthrough',
        textboxes: [{ floating: true, paras: [] }],
        imageParagraphSpaceBeforeAuto: true,
      },
    }
    const [first, , later] = plansOf([host, para([text('a')]), host])
    expect(first.spaceBeforePt).toBe(0)
    expect(later.spaceBeforePt).toBe(PROPS.autoSpacingPt)
  })

  it('section-break strips get no plan (measureBlocks skips them)', () => {
    const plans = plansOf([
      {
        type: 'docProtected',
        attrs: { blockType: 'passthrough', label: 'Section break paragraph', previewText: '' },
      },
      para([text('a')]),
    ])
    expect(plans).toHaveLength(1)
  })
})

describe('buildBlockPlan: P3 table row policy', () => {
  const cell = (content: unknown[]) => ({
    type: 'docTableCell',
    attrs: { colspan: 1, rowspan: 1 },
    content,
  })
  const row = (attrs: Record<string, unknown>, cells: unknown[]) => ({
    type: 'docTableRow',
    attrs,
    content: cells,
  })
  const table = (rows: unknown[]) => ({
    type: 'docTable',
    attrs: { docxIndex: 3 },
    content: rows,
  })

  it('reads header / cantSplit / atLeast height / keepNext / forced break per row', () => {
    const [t] = plansOf([
      table([
        row({ repeatHeader: true, heightTwips: 440 }, [cell([para([text('h')])])]),
        row({ cantSplit: true, heightTwips: 600, heightRule: 'exact' }, [
          cell([para([text('a')], { keepNext: true })]),
        ]),
        row({}, [cell([para([text('b')], { pageBreakBefore: true })])]),
      ]),
    ])
    expect(t.kind).toBe('table')
    expect(t.hasCellBreak).toBe(true)
    expect(t.rowPlan).toEqual([
      { isHeader: true, cantSplit: false, minHPt: 22 },
      { isHeader: false, cantSplit: true, keepNext: true },
      { isHeader: false, cantSplit: false, hasForcedBreak: true },
    ])
  })

  it('keepNext inherited by the first cell paragraph comes from the resolved cell props', () => {
    const eff = { keepNext: true, source: { keepNext: 'paraStyle' } } as unknown as NonNullable<
      ReturnType<PlanProps['paraEffective']>
    >
    const [t] = plansOf([table([row({}, [cell([para([text('a')])])])])], {
      ...PROPS,
      cellParaEffective: (idx, r, c, p) =>
        idx === 3 && r === 0 && c === 0 && p === 0 ? eff : undefined,
      tableModel: () => ({ rows: [[{ paras: ['a'] }]] }),
    })
    expect(t.rowPlan?.[0].keepNext).toBe(true)
  })

  it('the first cell paragraph counts even behind a nested table', () => {
    const nested = {
      type: 'docNestedTable',
      attrs: {},
      content: [row({}, [cell([para([text('n')])])])],
    }
    const [t] = plansOf([table([row({}, [cell([nested, para([text('a')], { keepNext: true })])])])])
    expect(t.rowPlan?.[0].keepNext).toBe(true)
  })

  it('a row of vMerge continuations is flagged from the model', () => {
    const [t] = plansOf([table([row({}, [cell([para([])])]), row({}, [cell([para([])])])])], {
      ...PROPS,
      tableModel: () => ({ rows: [[{ paras: [''] }], [{ paras: [''], vMerge: 'continue' }]] }),
    })
    expect(t.rowPlan?.[1].vMergeContinue).toBe(true)
    expect(t.rowPlan?.[0].vMergeContinue).toBeUndefined()
  })
})

describe('buildBlockPlan: P2 space before', () => {
  it('direct twips, auto spacing, style chain, contextual spacing between same-style neighbours', () => {
    const eff = {
      spacing: { before: 240, after: 0, beforeAuto: false, afterAuto: false },
      contextualSpacing: true,
      source: {},
    } as unknown as NonNullable<ReturnType<PlanProps['paraEffective']>>
    const [a, b, c, d] = plansOf(
      [
        para([text('a')], { spaceBefore: 120 }),
        para([text('b')], { spaceBeforeAuto: true }),
        para([text('c')], { docxIndex: 1, styleId: 'S' }),
        para([text('d')], { docxIndex: 2, styleId: 'S' }),
      ],
      { ...PROPS, paraEffective: () => eff },
    )
    expect(a.spaceBeforePt).toBe(6)
    expect(b.spaceBeforePt).toBe(14)
    expect(c.spaceBeforePt).toBe(12)
    expect(d.spaceBeforePt).toBe(0)
    expect(a.dropAtPageTop).toBe(false)
    expect(b.dropAtPageTop).toBe(true)
  })

  it('contextual spacing compares styles: outline-only headings pair with Normal, the built-in heading look does not', () => {
    const heading = (attrs: Record<string, unknown>) => ({
      type: 'docHeading',
      attrs: { docxIndex: null, styleId: null, level: 1, ...attrs },
      content: [text('h')],
    })
    const item = para([text('i')], { spaceBefore: 240, contextualSpacing: true })
    const [, a] = plansOf([heading({ outlineOnly: true }), item])
    const [, b] = plansOf([heading({}), item])
    const [, c] = plansOf([heading({ outlineOnly: true, styleId: 'H' }), item])
    expect(a.spaceBeforePt).toBe(0)
    expect(b.spaceBeforePt).toBe(12)
    expect(c.spaceBeforePt).toBe(12)
  })

  it('a direct contextualSpacing on a styled paragraph suppresses like a style-level one', () => {
    const eff = {
      spacing: { before: 240, after: 160, beforeAuto: false, afterAuto: false },
      contextualSpacing: true,
      source: { contextualSpacing: 'direct' },
    } as unknown as NonNullable<ReturnType<PlanProps['paraEffective']>>
    const [, a, b, c] = plansOf(
      [
        para([text('body')]),
        para([text('one')], { docxIndex: 1, styleId: 'LP', contextualSpacing: true }),
        para([text('two')], { docxIndex: 2, styleId: 'LP', contextualSpacing: true }),
        para([text('off')], { docxIndex: 3, styleId: 'LP', contextualSpacing: false }),
      ],
      { ...PROPS, paraEffective: (i) => (i >= 1 ? eff : undefined) },
    )
    expect(a.spaceBeforePt).toBe(12)
    expect(b.spaceBeforePt).toBe(0)
    expect(c.spaceBeforePt).toBe(12)
  })

  it('a forced section start keeps its space before', () => {
    const [, b] = plansOf([para([text('a')]), para([text('b')], { docxIndex: 5 })], {
      ...PROPS,
      forcedSectionStarts: new Set([5]),
    })
    expect(b.dropAtPageTop).toBe(false)
  })
})

describe('buildBlockPlan: float geometry (tableGeom / anchorGeom / floatShift)', () => {
  const table = (attrs: Record<string, unknown>) => ({
    type: 'docTable',
    attrs: { docxIndex: 1, ...attrs },
    content: [
      {
        type: 'docTableRow',
        attrs: {},
        content: [{ type: 'docTableCell', attrs: { colspan: 1, rowspan: 1 }, content: [para([])] }],
      },
    ],
  })
  const floatAttrs = {
    tblFloat: 'left',
    tblFloatSource: true,
    tblFloatXTwips: 0,
    tblFloatYTwips: 1440,
    tblFloatHorzAnchor: 'margin',
    tblFloatVertAnchor: 'page',
    tblFloatDistance: { top: 0, right: 180, bottom: 0, left: 180 },
    widthPx: 300,
  }

  it('a page-anchored floating table carries its target Y and band', () => {
    const [t] = plansOf([table(floatAttrs)])
    expect(t.tableGeom).toEqual({ floated: true, pageRelV: { yPx: 96, band: 'page' } })
    expect(t.floatShift).toBeUndefined()
  })

  it('a keyword Y zeroes the offset and keeps the spec; a text anchor has no target', () => {
    const [kw, text] = plansOf([
      table({ ...floatAttrs, tblFloatYSpec: 'bottom', tblFloatVertAnchor: 'margin' }),
      table({ ...floatAttrs, tblFloatVertAnchor: 'text', tblFloatYTwips: 200 }),
    ])
    expect(kw.tableGeom?.pageRelV).toEqual({ yPx: 0, band: 'margin', spec: 'bottom' })
    expect(text.tableGeom).toEqual({ floated: true })
  })

  it('a suppressed float keeps its lift; a text float with topFromText pushes whole', () => {
    const [lift, push] = plansOf([
      table({
        ...floatAttrs,
        tblFloatSuppressed: true,
        tblFloatVertAnchor: 'text',
        tblFloatYTwips: -300,
      }),
      table({
        ...floatAttrs,
        tblFloatVertAnchor: 'text',
        tblFloatYTwips: 0,
        tblFloatDistance: { top: 180, right: 0, bottom: 0, left: 0 },
      }),
    ])
    expect(lift.tableGeom).toEqual({ floated: false, liftPx: 20 })
    expect(push.tableGeom).toEqual({ floated: true, pushWhole: true })
  })

  it('the engine shift state rides the plan; a flowed table loses its target with the float', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [table(floatAttrs), table(floatAttrs)],
      } as unknown as PmNode as never,
    })
    const shifts = new Map([
      [0, { dyPx: 12.5, flow: false }],
      [editor.state.doc.child(0).nodeSize, { dyPx: 0, flow: true }],
    ])
    const [a, b] = [...buildBlockPlan(editor.state.doc, PROPS, (pos) => shifts.get(pos)).values()]
    editor.destroy()
    expect(a.floatShift).toEqual({ dyPx: 12.5, flow: false })
    expect(a.tableGeom?.pageRelV).toBeDefined()
    expect(b.floatShift).toEqual({ dyPx: 0, flow: true })
    expect(b.tableGeom).toBeUndefined()
  })

  it('a picture block with a margin Y carries the target for in-flow wraps only', () => {
    const pic = (attrs: Record<string, unknown>) => ({
      type: 'docProtected',
      attrs: {
        blockType: 'image',
        imageSrc: 'x.png',
        imageWidthPx: 100,
        imageHeightPx: 50,
        imageRelV: 'margin',
        imageOffsetYEmu: 0,
        ...attrs,
      },
    })
    const [band, side, front, banded] = plansOf([
      pic({ imageWrap: 'topBottom' }),
      pic({ imageWrap: 'square-left' }),
      pic({ imageWrap: 'front' }),
      pic({ imageWrap: 'square-right', imageBand: true, imageOffsetXEmu: 0 }),
    ])
    expect(band.anchorGeom).toEqual({ floated: false, pageRelV: { yPx: 0, band: 'margin' } })
    expect(side.anchorGeom).toEqual({ floated: true, pageRelV: { yPx: 0, band: 'margin' } })
    expect(front.anchorGeom).toBeUndefined()
    expect(banded.anchorGeom).toBeUndefined()
  })

  it('a floating text-box block with a column-spanning band keeps the band', () => {
    const [keep, plain, inline] = plansOf([
      {
        type: 'docProtected',
        attrs: {
          blockType: 'passthrough',
          textboxes: [{ floating: true, bandOverflow: true, paras: [] }],
        },
      },
      {
        type: 'docProtected',
        attrs: { blockType: 'passthrough', textboxes: [{ floating: true, paras: [] }] },
      },
      {
        type: 'docProtected',
        attrs: { blockType: 'passthrough', textboxes: [{ floating: false, paras: [] }] },
      },
    ])
    expect(keep.anchorGeom).toEqual({ floated: false, bandKeep: true })
    expect(plain.anchorGeom).toEqual({ floated: false })
    expect(inline.anchorGeom).toBeUndefined()
  })
})
