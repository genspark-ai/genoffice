import { describe, expect, it } from 'vitest'

import { applyStructuralOps } from '../src/gateway/xlsx-structure'

const SHEET = 'Data'

/// D2:D5 filled down from D2 — the shape Excel writes for a column of
/// per-row formulas.
const filledColumn =
  '<worksheet><sheetData>' +
  '<row r="2"><c r="D2"><f t="shared" ref="D2:D5" si="0">B2*C2</f><v>2</v></c></row>' +
  '<row r="3"><c r="D3"><f t="shared" si="0"/><v>6</v></c></row>' +
  '<row r="4"><c r="D4"><f t="shared" si="0"/><v>12</v></c></row>' +
  '<row r="5"><c r="D5"><f t="shared" si="0"/><v>20</v></c></row>' +
  '</sheetData></worksheet>'

describe('removing the cell that anchors a shared formula', () => {
  it('turns the surviving followers into plain formulas when the anchor row goes', () => {
    const xml = applyStructuralOps(
      filledColumn,
      [{ kind: 'remove-rows', index: 1, count: 1 }],
      SHEET,
    )
    expect(xml).not.toContain('t="shared"')
    expect(xml).toContain('<c r="D2"><f>B2*C2</f><v>6</v></c>')
    expect(xml).toContain('<c r="D3"><f>B3*C3</f><v>12</v></c>')
    expect(xml).toContain('<c r="D4"><f>B4*C4</f><v>20</v></c>')
  })

  it('turns the surviving followers into plain formulas when the anchor column goes', () => {
    const filledRow =
      '<worksheet><sheetData><row r="1">' +
      '<c r="A1"><f t="shared" ref="A1:C1" si="0">A2+A3</f></c>' +
      '<c r="B1"><f t="shared" si="0"/></c>' +
      '<c r="C1"><f t="shared" si="0"/></c>' +
      '</row></sheetData></worksheet>'
    const xml = applyStructuralOps(filledRow, [{ kind: 'remove-cols', index: 0, count: 1 }], SHEET)
    expect(xml).not.toContain('t="shared"')
    expect(xml).toContain('<c r="A1"><f>A2+A3</f></c>')
    expect(xml).toContain('<c r="B1"><f>B2+B3</f></c>')
  })

  it('keeps the group shared when only followers are removed', () => {
    const xml = applyStructuralOps(
      filledColumn,
      [{ kind: 'remove-rows', index: 2, count: 2 }],
      SHEET,
    )
    expect(xml).toContain('<c r="D2"><f t="shared" ref="D2:D3" si="0">B2*C2</f><v>2</v></c>')
    expect(xml).toContain('<c r="D3"><f t="shared" si="0"/><v>20</v></c>')
  })
})
