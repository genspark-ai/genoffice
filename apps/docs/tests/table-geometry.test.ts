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

  it('a nested table never holds past its cell', () => {
    const css = tableWidthCss(modelWidthInput(wide, 800), { kind: 'nested', surface: 'static' })
    expect(css.hold).toBe(false)
    expect(css.styles).toEqual(['width:min(800px,100%)'])
  })
})
