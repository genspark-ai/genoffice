import { describe, expect, it } from 'vitest'

import { draftFromSelection, formatCellsCommands } from '../src/renderer/format-cells'

import type { SelectionFormat } from '../src/renderer/selection-format'

const SELECTION: SelectionFormat = {
  fontFamily: 'Calibri',
  fontSize: 14,
  bold: true,
  italic: false,
  underline: false,
  strike: false,
  vertAlign: null,
  wrap: false,
  horizontalAlignment: 'left',
  verticalAlignment: 'center',
  textRotation: 45,
  indent: 0,
  fontColor: '#C00000',
  fillColor: '#FFF2CC',
  numberFormat: '0.00%',
  link: null,
}

describe('draftFromSelection', () => {
  it('prefills every echoed setting', () => {
    expect(draftFromSelection(SELECTION)).toEqual({
      pattern: '0.00%',
      family: 'Calibri',
      size: '14',
      bold: 'on',
      italic: 'off',
      underline: 'off',
      strike: 'off',
      fontColor: '#c00000',
      border: '',
      borderColor: '#000000',
      borderStyle: 'thin',
      fill: '#fff2cc',
      noFill: false,
      hAlign: 'left',
      vAlign: 'middle',
      wrapText: 'off',
      rotation: '45',
      indent: '',
      locked: '',
      hidden: '',
    })
  })

  it('falls back to unchanged tri-states without a selection', () => {
    const draft = draftFromSelection(null)
    expect(draft.pattern).toBe('')
    expect(draft.bold).toBe('')
    expect(draft.fontColor).toBe('')
    expect(draft.fill).toBe('')
  })
})

describe('formatCellsCommands', () => {
  const initial = draftFromSelection(SELECTION)

  it('emits nothing when the draft is untouched', () => {
    expect(formatCellsCommands(initial, initial)).toEqual([])
  })

  it('emits only the changed settings', () => {
    const edited = { ...initial, pattern: '$#,##0.00', bold: 'off' as const, fill: '#00aa55' }
    expect(formatCellsCommands(initial, edited)).toEqual([
      'format:$#,##0.00',
      'bold:off',
      'fill:#00aa55',
    ])
  })

  it('re-picking the current values emits nothing', () => {
    const edited = { ...initial, family: 'Calibri', fontColor: '#c00000' }
    expect(formatCellsCommands(initial, edited)).toEqual([])
  })

  it('any border preset counts as a change and carries its color and line style', () => {
    const edited = { ...initial, border: 'all', borderColor: '#112233' }
    expect(formatCellsCommands(initial, edited)).toEqual(['border:all:#112233:thin'])
    const doubled = { ...initial, border: 'bottom', borderStyle: 'double' }
    expect(formatCellsCommands(initial, doubled)).toEqual(['border:bottom:#000000:double'])
  })

  it('no-fill clears only when the selection had a fill', () => {
    expect(formatCellsCommands(initial, { ...initial, noFill: true })).toEqual(['fill:none'])
    const unfilled = draftFromSelection({ ...SELECTION, fillColor: null })
    expect(formatCellsCommands(unfilled, { ...unfilled, noFill: true })).toEqual([])
  })

  it('explicit white on an unfilled selection emits a fill command', () => {
    const unfilled = draftFromSelection({ ...SELECTION, fillColor: null })
    expect(unfilled.fill).toBe('')
    expect(formatCellsCommands(unfilled, { ...unfilled, fill: '#ffffff' })).toEqual([
      'fill:#ffffff',
    ])
  })

  it('emits alignment, wrap, rotation, and indent changes', () => {
    const edited = {
      ...initial,
      hAlign: 'distributed',
      vAlign: 'top',
      wrapText: 'on' as const,
      rotation: 'vertical',
      indent: '2',
    }
    expect(formatCellsCommands(initial, edited)).toEqual([
      'align:distributed',
      'valign:top',
      'wrap:on',
      'rotate:vertical',
      'indent:2',
    ])
  })

  it('re-picking the echoed alignment emits nothing', () => {
    const edited = { ...initial, hAlign: 'left', vAlign: 'middle', wrapText: 'off' as const }
    expect(formatCellsCommands(initial, edited)).toEqual([])
  })

  it('echoes an existing indent and clears it with an explicit 0', () => {
    const indented = draftFromSelection({ ...SELECTION, indent: 3 })
    expect(indented.indent).toBe('3')
    expect(formatCellsCommands(indented, indented)).toEqual([])
    expect(formatCellsCommands(indented, { ...indented, indent: '0' })).toEqual(['indent:0'])
  })

  it('combines protection flags into one command', () => {
    expect(formatCellsCommands(initial, { ...initial, locked: 'off' as const })).toEqual([
      'cellprot:locked-off:',
    ])
    expect(
      formatCellsCommands(initial, { ...initial, locked: 'off' as const, hidden: 'on' as const }),
    ).toEqual(['cellprot:locked-off:hidden-on'])
  })
})
