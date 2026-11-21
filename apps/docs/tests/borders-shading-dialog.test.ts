import { describe, expect, it } from 'vitest'
import type { SectionSettings } from '@genoffice/docx-engine'
import {
  borderDraftOf,
  borderSpecOfDraft,
  sectionWithPageBorder,
  withPen,
  withSetting,
  withSideToggled,
} from '../src/renderer/components/BordersShadingDialog'
import {
  borderAttrsOf,
  borderSpecOf,
  borderSpecOfPayload,
  shadingAttrsOf,
  shadingSpecOf,
  shadingSpecOfPayload,
} from '../src/renderer/commands/borders-shading'
import {
  paraBorderCss,
  paraBorderDropColor,
  paraBorderShadowDecls,
} from '../src/renderer/editor/para-border-css'
import { darkPageColor } from '../src/renderer/editor/dark-page'
import { SHADING_HATCHES, shadingHatchCss } from '../src/renderer/editor/shading-hatch'
import { textDirectionOf } from '../src/renderer/commands/text-direction'
import { CHART_TYPES, chartKindOf } from '../src/renderer/components/chart-types'

describe('Borders and Shading', () => {
  const LINES = JSON.stringify({
    t: { color: 'C00000', szPt: 1.5, spacePt: 1, val: 'double' },
    b: { szPt: 0.75 },
  })

  it('reads the drawn sides with the look each declares, and writes the same attributes back', () => {
    const spec = borderSpecOf('bt', LINES)
    expect(spec).toEqual({
      t: { color: 'C00000', szPt: 1.5, spacePt: 1, val: 'double' },
      b: { szPt: 0.75 },
    })
    expect(borderAttrsOf(spec)).toEqual({ borders: 'tb', borderLines: LINES })
    expect(borderSpecOf('tblr', null)).toEqual({ t: {}, b: {}, l: {}, r: {} })
    expect(borderAttrsOf({ t: {}, b: {} })).toEqual({ borders: 'tb', borderLines: null })
    expect(borderAttrsOf({})).toEqual({ borders: null, borderLines: null })
    expect(borderSpecOf(null, '{broken')).toEqual({})
  })

  it('opens on what the paragraph has and gives it back unchanged when nothing is touched', () => {
    const spec = borderSpecOf('tb', LINES)
    const draft = borderDraftOf(spec)
    expect(draft.setting).toBe('custom')
    expect(draft.pen).toEqual({ val: 'double', color: 'C00000', szPt: 1.5 })
    expect(borderSpecOfDraft(draft)).toEqual(spec)
    expect(borderDraftOf({}).setting).toBe('none')
    expect(borderDraftOf({ t: {}, b: {}, l: {}, r: {} }).setting).toBe('box')
  })

  it('Shadow and 3-D put the pen on all four sides with the setting on each line', () => {
    const shadow = withSetting(withPen(borderDraftOf({}), { szPt: 1.5 }), 'shadow')
    const line = { szPt: 1.5, shadow: true }
    expect(borderSpecOfDraft(shadow)).toEqual({ t: line, b: line, l: line, r: line })
    expect(borderDraftOf(borderSpecOfDraft(shadow)).setting).toBe('shadow')
    expect(borderAttrsOf(borderSpecOfDraft(shadow)).borderLines).toContain('"shadow":true')

    const framed = withSetting(shadow, '3d')
    expect(borderSpecOfDraft(framed).t).toEqual({ szPt: 1.5, frame: true })
    expect(borderDraftOf(borderSpecOfDraft(framed)).setting).toBe('3d')
    // a new width keeps the setting; Box takes it off again
    expect(borderSpecOfDraft(withPen(framed, { szPt: 3 })).b).toEqual({ szPt: 3, frame: true })
    expect(borderSpecOfDraft(withSetting(framed, 'box')).t).toEqual({ szPt: 1.5 })
    expect(borderSpecOfPayload(JSON.stringify(borderSpecOfDraft(shadow)))).toEqual(
      borderSpecOfDraft(shadow),
    )
  })

  it('a shadowed border drops at the bottom-right in its own colour, a 3-D one bevels', () => {
    const lines = { t: { szPt: 1.5, color: 'C00000', shadow: true as const } }
    expect(paraBorderShadowDecls(lines)).toEqual([
      'box-shadow:4px var(--pbdr-drop-y, 4px) 0 var(--pbdr-drop-c, #C00000)',
      `--dk-drop:${darkPageColor('#C00000')}`,
    ])
    expect(paraBorderDropColor(lines)).toBe('#C00000')
    expect(paraBorderDropColor({ t: { szPt: 1.5 } })).toBeNull()
    expect(paraBorderShadowDecls({ t: { szPt: 1.5 } })).toEqual([])
    expect(paraBorderCss({ szPt: 0.5, frame: true })).toBe('3px ridge #444')
    expect(paraBorderCss({ szPt: 4.5, color: '0070C0', frame: true })).toBe('6px ridge #0070C0')
    // a dashed or double line keeps its own style
    expect(paraBorderCss({ szPt: 1, val: 'dashed', frame: true })).toContain('dashed')
  })

  it('a hatch pattern draws as stripes in its colour; every other pattern has no hatch', () => {
    expect(SHADING_HATCHES).toHaveLength(12)
    for (const val of SHADING_HATCHES) expect(shadingHatchCss(val)).toContain('#000000')
    expect(shadingHatchCss('diagStripe', 'C00000')).toBe(
      'repeating-linear-gradient(135deg,#C00000 0 3px,transparent 3px 6px)',
    )
    expect(shadingHatchCss('thinHorzCross')).toBe(
      'repeating-linear-gradient(0deg,#000000 0 1px,transparent 1px 4px),' +
        'repeating-linear-gradient(90deg,#000000 0 1px,transparent 1px 4px)',
    )
    for (const val of ['pct10', 'solid', 'clear', 'nil', 'thin'])
      expect(shadingHatchCss(val)).toBeNull()
  })

  it('a new colour on a box keeps the style and the width of its lines', () => {
    const line = { szPt: 2.25, val: 'double' }
    const draft = withPen(borderDraftOf({ t: line, b: line, l: line, r: line }), {
      color: '0070C0',
    })
    const blue = { color: '0070C0', szPt: 2.25, val: 'double' }
    expect(borderSpecOfDraft(draft)).toEqual({ t: blue, b: blue, l: blue, r: blue })
  })

  it("Box draws the pen on four sides, None takes them off, and a side click is Word's", () => {
    const pen = { val: 'dotted', color: null, szPt: 3 }
    let draft = withPen(borderDraftOf({ b: { szPt: 0.75 } }), pen)
    // Custom: the pen waits for a click
    expect(borderSpecOfDraft(draft)).toEqual({ b: { szPt: 0.75 } })
    // a side drawn with another pen takes this one; clicked again it comes off
    draft = withSideToggled(draft, 'b')
    expect(borderSpecOfDraft(draft)).toEqual({ b: { szPt: 3, val: 'dotted' } })
    draft = withSideToggled(draft, 'b')
    expect(draft.setting).toBe('none')
    draft = withSetting(draft, 'box')
    expect(Object.keys(borderSpecOfDraft(draft))).toEqual(['t', 'b', 'l', 'r'])
    expect(borderSpecOfDraft(draft).l).toEqual({ szPt: 3, val: 'dotted' })
    draft = withSideToggled(draft, 'l')
    expect(draft.setting).toBe('custom')
    expect(borderSpecOfDraft(withSetting(draft, 'none'))).toEqual({})
  })

  it('takes a border payload as JSON only, with nothing in it but sides and their look', () => {
    expect(borderSpecOfPayload('tb')).toBeNull()
    expect(
      borderSpecOfPayload(
        JSON.stringify({ t: { val: 'dashed', color: 'not-hex', szPt: -1, extra: 1 }, x: {} }),
      ),
    ).toEqual({ t: { val: 'dashed' } })
  })

  it('draws a line in its style, a double wide enough for both of its lines', () => {
    expect(paraBorderCss()).toBe('1px solid #444')
    expect(paraBorderCss({ color: 'FF0000', szPt: 1.5 })).toBe('2px solid #FF0000')
    expect(paraBorderCss({ szPt: 1.5, val: 'dotted' })).toBe('2px dotted #444')
    expect(paraBorderCss({ szPt: 1.5, val: 'dashed' })).toBe('2px dashed #444')
    expect(paraBorderCss({ szPt: 1.5, val: 'double' })).toBe('6px double #444')
    expect(paraBorderCss({ szPt: 0.5, val: 'double' })).toBe('3px double #444')
  })

  it('a pattern shows as its colour blended over the fill and saves as itself', () => {
    const spec = shadingSpecOf('FFF2CC', JSON.stringify({ val: 'pct25', color: 'C00000' }))
    expect(spec).toEqual({ fill: 'FFF2CC', pattern: 'pct25', color: 'C00000' })
    expect(shadingAttrsOf(spec)).toEqual({
      shadingFill: 'FFF2CC',
      shadingPattern: JSON.stringify({ val: 'pct25', color: 'C00000' }),
      shadingDisplay: 'EFB699',
    })
    expect(shadingAttrsOf({ fill: 'FFF2CC', pattern: null, color: null })).toEqual({
      shadingFill: 'FFF2CC',
      shadingPattern: null,
      shadingDisplay: null,
    })
    expect(shadingAttrsOf({ fill: null, pattern: 'solid', color: null }).shadingDisplay).toBe(
      '000000',
    )
    expect(shadingSpecOfPayload('FFF2CC')).toBeNull()
    expect(shadingSpecOfPayload(JSON.stringify(spec))).toEqual(spec)
    expect(shadingSpecOfPayload(JSON.stringify({ fill: 'FFF2CC', pattern: null }))).toEqual({
      fill: 'FFF2CC',
      pattern: null,
      color: null,
    })
  })
})

describe('Borders and Shading page border', () => {
  it('puts what the Page Border tab held on the section, or takes the border off', () => {
    const section = { marginTop: 1440, pageBorder: true, pageBorderProps: { display: 'allPages' } }
    const props = { display: 'firstPage' }
    expect(
      sectionWithPageBorder(
        section as SectionSettings,
        props as SectionSettings['pageBorderProps'],
      ),
    ).toEqual({ marginTop: 1440, pageBorder: true, pageBorderProps: props })
    expect(sectionWithPageBorder(section as SectionSettings, null)).toEqual({
      marginTop: 1440,
      pageBorder: false,
    })
  })
})

describe('Text Direction', () => {
  const values = (target: Parameters<typeof textDirectionOf>[0]) =>
    textDirectionOf(target, {}).choices.map((c) => c.value)

  it('opens on the direction the cell has', () => {
    expect(textDirectionOf('cells', { cell: null }).current).toBe('lrTb')
    expect(textDirectionOf('cells', { cell: 'tbRl' }).current).toBe('tbRl')
    expect(textDirectionOf('cells', { cell: 'btLr' }).current).toBe('btLr')
    // a direction the dialog does not offer is not replaced by a default
    expect(textDirectionOf('cells', { cell: 'tbRlV' }).current).toBeNull()
  })

  it('offers each target the directions the page lays out for it', () => {
    expect(values('section')).toEqual(['lrTb', 'tbRl'])
    expect(values('cells')).toEqual(['lrTb', 'tbRl', 'btLr'])
    expect(values('textBox')).toEqual(['horz', 'eaVert', 'vert', 'vert270', 'wordArtVert'])
  })

  it('a section reads vertical in any flow the page turns, horizontal in the others', () => {
    const section = (textDirection?: string) =>
      textDirectionOf('section', { section: { textDirection } as SectionSettings }).current
    expect(section()).toBe('lrTb')
    expect(section('tbRl')).toBe('tbRl')
    expect(section('btLr')).toBe('tbRl')
    expect(section('lrTbV')).toBe('lrTb')
    expect(textDirectionOf('textBox', { box: null }).current).toBe('horz')
    expect(textDirectionOf('textBox', { box: 'vert270' }).current).toBe('vert270')
  })
})

describe('Insert > Chart types', () => {
  it("Word's Bar is a bar chart lying down; a type of the chart extension is a layout over a classic kind; the others are their own kind", () => {
    expect(chartKindOf('bar')).toEqual({ kind: 'bar' })
    expect(chartKindOf('barH')).toEqual({ kind: 'bar', horizontal: true })
    expect(chartKindOf('waterfall')).toEqual({ kind: 'bar', layout: 'waterfall' })
    expect(chartKindOf('boxWhisker')).toEqual({ kind: 'bar', layout: 'boxWhisker' })
    expect(chartKindOf('treemap')).toEqual({ kind: 'pie', layout: 'treemap' })
    const extension = CHART_TYPES.filter((t) => chartKindOf(t).layout !== undefined)
    expect(extension).toHaveLength(7)
    for (const type of CHART_TYPES.filter((t) => t !== 'barH' && !extension.includes(t)))
      expect(chartKindOf(type)).toEqual({ kind: type })
  })
})
