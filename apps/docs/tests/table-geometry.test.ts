import { describe, expect, it } from 'vitest'
import type { TableModel } from '@genoffice/docx-engine'
import {
  modelWidthInput,
  resolveTableWidth,
  tablePositionCss,
  tableWidthCss,
} from '../src/renderer/editor/table-geometry'

const cell = (text: string) => ({ paras: [text] })
const CW = 'var(--doc-content-w,100%)'
const CTX = { kind: 'body', surface: 'static' } as const

describe('table-geometry: over-wide explicit dxa tables (Word probe 2026-10-29)', () => {
  const wide: TableModel = {
    rows: [[cell('c0'), cell('c1'), cell('c2')]],
    colWidthsTwips: [3000, 6000, 3000],
    dxaWidth: true,
  }
  const letter = { avail: 10800, fit: 9360, paper: 12240 }

  it('centred: held, centred on the margin box, bleeding both margins', () => {
    const geom = resolveTableWidth({ ...wide, align: 'center' }, letter, { legacyIndent: false })
    expect(geom).toMatchObject({
      hold: true,
      origin: 'marginBox',
      bleed: 'both',
      widthTwips: 12000,
    })
    const css = tableWidthCss(modelWidthInput(geom.model, 800), CTX)
    expect(css.styles).toEqual(['width:800px', 'max-width:none'])
    expect(tablePositionCss(modelWidthInput(geom.model, 800), css)).toEqual([
      `margin-left:calc((${CW} - 800px)/2)`,
    ])
  })

  it('right-aligned: held, right edge pinned to the right margin, bleeding left', () => {
    const geom = resolveTableWidth({ ...wide, align: 'right' }, letter, { legacyIndent: false })
    expect(geom).toMatchObject({ hold: true, origin: 'rightMargin', bleed: 'left' })
    const css = tableWidthCss(modelWidthInput(geom.model, 800), CTX)
    expect(tablePositionCss(modelWidthInput(geom.model, 800), css)).toEqual([
      `margin-left:calc(${CW} - 800px)`,
    ])
  })

  it('a tblW-auto centred grid still compresses to the paper', () => {
    const auto: TableModel = { ...wide, dxaWidth: undefined, autoLayout: true, align: 'center' }
    const geom = resolveTableWidth(auto, letter, { legacyIndent: false })
    expect(geom.hold).toBe(false)
    expect(geom.widthTwips).toBeLessThanOrEqual(9360)
  })

  it('bidiVisual: every surface agrees on the paper clamp (no Word probe yet)', () => {
    const bidi: TableModel = { ...wide, align: 'center', bidiVisual: true }
    const geom = resolveTableWidth(bidi, letter, { legacyIndent: false })
    expect(geom.hold).toBe(false)
    expect(geom.widthTwips).toBeLessThanOrEqual(12240)
    const editorInput = {
      ...modelWidthInput(bidi, 800),
      autoFit: 'fixed' as const,
      autoFitDeclared: 'fixed' as const,
    }
    expect(tableWidthCss(editorInput, { kind: 'body', surface: 'editor' }).hold).toBe(false)
    expect(tableWidthCss(modelWidthInput(bidi, 800), CTX).hold).toBe(false)
    const right: TableModel = { ...bidi, align: 'right' }
    expect(resolveTableWidth(right, letter, { legacyIndent: false }).hold).toBe(false)
    expect(tableWidthCss(modelWidthInput(right, 800), CTX).hold).toBe(false)
  })

  it('over-wide and narrowed grids flag overWide so w:noWrap cells wrap like Word', () => {
    expect(
      resolveTableWidth({ ...wide, align: 'center' }, letter, { legacyIndent: false }).overWide,
    ).toBe(true)
    const auto: TableModel = { ...wide, dxaWidth: undefined, autoLayout: true }
    const narrowed = resolveTableWidth(auto, letter, { legacyIndent: false })
    expect(narrowed.bleed).toBe('none')
    expect(narrowed.overWide).toBe(true)
    const fits: TableModel = { ...wide, colWidthsTwips: [3000, 3000, 3000] }
    expect(resolveTableWidth(fits, letter, { legacyIndent: false }).overWide).toBe(false)
    // a pct table resolving below its saved grid still has room to grow
    const pct: TableModel = {
      ...wide,
      dxaWidth: undefined,
      autoLayout: true,
      widthPct: 50,
      colWidthsPct: [25, 50, 25],
    }
    expect(resolveTableWidth(pct, letter, { legacyIndent: false }).overWide).toBe(false)
    expect(
      resolveTableWidth({ ...pct, widthPct: 120 }, letter, { legacyIndent: false }).overWide,
    ).toBe(true)
  })

  it('a nested table never holds past its cell', () => {
    const css = tableWidthCss(modelWidthInput(wide, 800), { kind: 'nested', surface: 'static' })
    expect(css.hold).toBe(false)
    expect(css.styles).toEqual(['width:min(800px,100%)'])
  })
})

describe('table-geometry: fixed-layout grids past the paper', () => {
  const letter = { avail: 10800, fit: 9360, paper: 12240 }
  const fixed: TableModel = {
    rows: [[cell('c0'), cell('c1'), cell('c2'), cell('c3')]],
    colWidthsTwips: [3024, 3600, 2592, 3024],
    fixedLayout: true,
    dxaWidth: true,
  }

  it('holds the declared columns (Word clips at the paper, never rewraps)', () => {
    const wider = { ...fixed, colWidthsTwips: [3456, 5616, 3888, 7344] }
    const geom = resolveTableWidth(wider, letter, { legacyIndent: false })
    expect(geom.hold).toBe(true)
    expect(geom.colTwips).toEqual([3456, 5616, 3888, 7344])
    expect(geom.overWide).toBe(true)
  })

  it('still clamps a garbage grid wider than two papers', () => {
    const garbage = { ...fixed, colWidthsTwips: [6236, 130620472, 100, 100] }
    const geom = resolveTableWidth(garbage, letter, { legacyIndent: false })
    expect(geom.hold).toBe(false)
    expect(geom.widthTwips).toBeLessThanOrEqual(12240)
  })
})
