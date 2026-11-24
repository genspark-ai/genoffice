// @vitest-environment jsdom
import { Editor, type JSONContent } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  convertTableToText,
  convertTextToTable,
  convertedRows,
  suggestedColumns,
  textToTableScope,
} from '../src/renderer/editor/table-convert'
import {
  cellNumber,
  bookmarkContext,
  evaluateFormula,
  formatFormulaResult,
  formatNumberPicture,
  formulaFieldAt,
  formulaInstr,
  formulaPlace,
  insertTableFormula,
  NO_TABLE,
  parseFormulaInstr,
  refreshTableFormulas,
  suggestedFormula,
  type FormulaGrid,
} from '../src/renderer/editor/table-formula'

const para = (text: string, marks?: JSONContent['marks']) => ({
  type: 'docParagraph',
  content: text ? [{ type: 'text', text, ...(marks ? { marks } : {}) }] : [],
})

const tableOf = (rows: string[][]) => ({
  type: 'docTable',
  content: rows.map((row) => ({
    type: 'docTableRow',
    content: row.map((text) => ({ type: 'docTableCell', content: [para(text)] })),
  })),
})

function makeEditor(content: JSONContent[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

const gridOf = (rows: string[][]): FormulaGrid => ({
  rows: rows.length,
  cols: Math.max(...rows.map((row) => row.length)),
  text: (row, col) => rows[row]?.[col] ?? null,
})

/** the body as lines: a paragraph's text, or a table's rows as `a|b` */
const linesOf = (editor: Editor) => {
  const lines: string[] = []
  editor.state.doc.forEach((node) => {
    if (node.type.name !== 'docTable') {
      lines.push(node.textContent)
      return
    }
    node.forEach((row) => {
      const cells: string[] = []
      row.forEach((cell) => cells.push(cell.textContent))
      lines.push(cells.join('|'))
    })
  })
  return lines
}

/** puts the caret at the start of the first text block holding `text` */
function caretIn(editor: Editor, text: string, offset = 0) {
  let at = -1
  editor.state.doc.descendants((node, pos) => {
    if (at < 0 && node.isTextblock && node.textContent === text) at = pos + 1 + offset
  })
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)))
}

const fieldsOf = (editor: Editor) => {
  const fields: Array<{ instr: string; text: string; dirty: boolean }> = []
  editor.state.doc.descendants((node) => {
    const mark = node.marks.find((m) => m.type.name === 'instrField')
    if (mark)
      fields.push({ instr: mark.attrs.instr, text: node.text ?? '', dirty: mark.attrs.dirty })
  })
  return fields
}

describe('formula instructions', () => {
  it('parts the expression from the number picture', () => {
    expect(parseFormulaInstr('=SUM(ABOVE)')).toEqual({ expr: 'SUM(ABOVE)', format: null })
    expect(parseFormulaInstr(' =SUM(ABOVE) \\# "#,##0.00" ')).toEqual({
      expr: 'SUM(ABOVE)',
      format: '#,##0.00',
    })
    expect(parseFormulaInstr('=A1*2 \\# 0.0')).toEqual({ expr: 'A1*2', format: '0.0' })
    expect(parseFormulaInstr('PAGE')).toBeNull()
    expect(parseFormulaInstr('=SUM(ABOVE) \\* MERGEFORMAT')).toBeNull()
  })

  it('writes the instruction Word reads', () => {
    expect(formulaInstr('=SUM(ABOVE)', '')).toBe('=SUM(ABOVE)')
    expect(formulaInstr('SUM(LEFT)', '0.00')).toBe('=SUM(LEFT) \\# "0.00"')
  })
})

describe('cell numbers', () => {
  it('reads the first number a cell shows', () => {
    expect(cellNumber('12')).toBe(12)
    expect(cellNumber('$1,234.50')).toBe(1234.5)
    expect(cellNumber('-3.5 kg')).toBe(-3.5)
    expect(cellNumber('(12)')).toBe(-12)
    expect(cellNumber('Total')).toBeNull()
    expect(cellNumber('')).toBeNull()
  })
})

describe('evaluating a formula', () => {
  const grid = gridOf([
    ['Item', 'First', 'Second', ''],
    ['A', '10', '2.5', ''],
    ['B', '20', '4', ''],
    ['Sum', '', '', ''],
  ])

  it('sums the cells of a direction, skipping text', () => {
    expect(evaluateFormula('SUM(ABOVE)', grid, { row: 3, col: 1 })).toBe(30)
    expect(evaluateFormula('SUM(LEFT)', grid, { row: 1, col: 3 })).toBe(12.5)
    expect(evaluateFormula('SUM(BELOW)', grid, { row: 0, col: 2 })).toBe(6.5)
    expect(evaluateFormula('SUM(RIGHT)', grid, { row: 2, col: 0 })).toBe(24)
    expect(evaluateFormula('sum(above)', grid, { row: 3, col: 2 })).toBe(6.5)
  })

  it('counts a number inside a heading, as Word does', () => {
    const headed = gridOf([['Region 1'], ['5'], ['']])
    expect(evaluateFormula('SUM(ABOVE)', headed, { row: 2, col: 0 })).toBe(6)
  })

  it('ends a direction at the first empty cell', () => {
    const gap = gridOf([['5'], [''], ['7'], ['8'], ['']])
    expect(evaluateFormula('SUM(ABOVE)', gap, { row: 4, col: 0 })).toBe(15)
  })

  it('runs the six statistics over directions, cells and ranges', () => {
    const self = { row: 3, col: 1 }
    expect(evaluateFormula('AVERAGE(ABOVE)', grid, self)).toBe(15)
    expect(evaluateFormula('COUNT(ABOVE)', grid, self)).toBe(2)
    expect(evaluateFormula('MIN(ABOVE)', grid, self)).toBe(10)
    expect(evaluateFormula('MAX(ABOVE)', grid, self)).toBe(20)
    expect(evaluateFormula('PRODUCT(ABOVE)', grid, self)).toBe(200)
    expect(evaluateFormula('SUM(B2:C3)', grid, self)).toBe(36.5)
    expect(evaluateFormula('SUM(B2,C3)', grid, self)).toBe(14)
    expect(evaluateFormula('SUM(B2;C3;1)', grid, self)).toBe(15)
    expect(evaluateFormula('MAX(B2:B3,100)', grid, self)).toBe(100)
  })

  it('does arithmetic on cells and numbers', () => {
    const self = { row: 3, col: 3 }
    expect(evaluateFormula('B2+B3', grid, self)).toBe(30)
    expect(evaluateFormula('(B2+B3)*C2/5', grid, self)).toBe(15)
    expect(evaluateFormula('-B2+2^3', grid, self)).toBe(-2)
    expect(evaluateFormula('ROUND(C2/3,2)', grid, self)).toBe(0.83)
    expect(evaluateFormula('ABS(-4)+INT(2.9)+MOD(7,4)+SIGN(-9)', grid, self)).toBe(8)
    // a cell without a number counts as zero in arithmetic
    expect(evaluateFormula('A1+B2', grid, self)).toBe(10)
  })

  it('leaves what it does not evaluate alone', () => {
    const self = { row: 3, col: 1 }
    expect(evaluateFormula('IF(B2>5,1)', grid, self)).toBeNull()
    expect(evaluateFormula('AND(1)', grid, self)).toBeNull()
    expect(evaluateFormula('SUM(Sales A1:B2)', grid, self)).toBeNull()
    expect(evaluateFormula('SUM(9:9)', grid, self)).toBeNull()
    expect(evaluateFormula('SUM(Z:Z)', grid, self)).toBeNull()
    expect(evaluateFormula('SUM(Sales)', grid, self)).toBeNull()
    expect(evaluateFormula('SUM(Z9)', grid, self)).toBeNull()
    expect(evaluateFormula('B2/0', grid, self)).toBeNull()
    expect(evaluateFormula('SUM(ABOVE', grid, self)).toBeNull()
    expect(evaluateFormula('', grid, self)).toBeNull()
  })
})

describe('logic, whole rows and columns, bookmarks', () => {
  const grid = gridOf([
    ['Item', 'First', 'Second', ''],
    ['A', '10', '2.5', ''],
    ['B', '20', '4', ''],
    ['Sum', '', '', ''],
  ])
  const self = { row: 3, col: 3 }
  const other = gridOf([
    ['1', '2'],
    ['3', '4'],
  ])
  const context = {
    bookmark: (name: string) => ({ RATE: 'Rate: 8%', LABEL: 'none', OTHER: '3' })[name] ?? null,
    table: (name: string) => (name === 'OTHER' ? other : null),
  }
  const value = (expr: string, at = self) => evaluateFormula(expr, grid, at, context)

  it('compares with the six operators', () => {
    expect(value('B2=10')).toBe(1)
    expect(value('B2<>10')).toBe(0)
    expect(value('B2<B3')).toBe(1)
    expect(value('B2<=10')).toBe(1)
    expect(value('B2>B3')).toBe(0)
    expect(value('C3>=4.5')).toBe(0)
    expect(value('B2+B3=30')).toBe(1)
    expect(value('(B2>5)+(B3>5)')).toBe(2)
  })

  it('evaluates IF, AND, OR, NOT, TRUE and FALSE', () => {
    expect(value('IF(B2>5,100,200)')).toBe(100)
    expect(value('if(B2 > 50; B2; B3*2)')).toBe(40)
    expect(value('IF(SUM(ABOVE)>=30,1,0)', { row: 3, col: 1 })).toBe(1)
    expect(value('AND(B2>5,B3>50)')).toBe(0)
    expect(value('OR(B2>5,B3>50)')).toBe(1)
    expect(value('NOT(B2>5)')).toBe(0)
    expect(value('NOT(0)')).toBe(1)
    expect(value('IF(AND(B2=10,NOT(C2=1)),TRUE,FALSE)')).toBe(1)
    expect(value('TRUE()+FALSE()+TRUE')).toBe(2)
  })

  it('does not fail on the branch IF leaves', () => {
    expect(value('IF(D2=0,0,B2/D2)')).toBe(0)
    expect(value('IF(D2=0,B2/D2,0)')).toBeNull()
  })

  it('answers DEFINED by whether its argument evaluates', () => {
    expect(value('DEFINED(Rate)')).toBe(1)
    expect(value('DEFINED(Missing)')).toBe(0)
    expect(value('DEFINED(B2/0)')).toBe(0)
    expect(value('DEFINED((B2+1)*2)')).toBe(1)
    expect(value('IF(DEFINED(Missing),Missing,7)')).toBeNull()
    expect(value('IF(DEFINED(SUM(Missing)),1,7)')).toBe(7)
    expect(value('DEFINED(B2')).toBeNull()
  })

  it('reads whole rows and whole columns, leaving its own cell out', () => {
    expect(value('SUM(B:B)')).toBe(30)
    expect(value('SUM(b:c)')).toBe(36.5)
    expect(value('SUM(2:2)')).toBe(12.5)
    expect(value('SUM(2:3)')).toBe(36.5)
    expect(value('COUNT(3:2)')).toBe(4)
    expect(value('SUM(B:B)', { row: 2, col: 1 })).toBe(10)
    expect(value('MAX(1:1,B:B)')).toBe(20)
  })

  it('reads a bookmark as the number it marks', () => {
    expect(value('Rate*100')).toBe(800)
    expect(value('rate+B2')).toBe(18)
    expect(value('SUM(Rate,B2)')).toBe(18)
    expect(value('Label+1')).toBe(1)
    expect(value('Missing+1')).toBeNull()
    expect(evaluateFormula('Rate*100', grid, self)).toBeNull()
  })

  it('reads a name shaped like a cell as a bookmark where the table has no such cell', () => {
    const named = {
      bookmark: (name: string) => ({ Q4: 'total 40', B2: '99' })[name] ?? null,
      table: () => null,
    }
    // inside the grid the cell wins
    expect(evaluateFormula('B2+1', grid, self, named)).toBe(11)
    expect(evaluateFormula('Q4+1', grid, self, named)).toBe(41)
    expect(evaluateFormula('SUM(Q4,B2)', grid, self, named)).toBe(50)
    expect(evaluateFormula('Q4*2', NO_TABLE, { row: -1, col: -1 }, named)).toBe(80)
    expect(evaluateFormula('B2*2', NO_TABLE, { row: -1, col: -1 }, named)).toBe(198)
    expect(evaluateFormula('Z9+1', grid, self, named)).toBeNull()
    expect(evaluateFormula('SUM(Z9)', grid, self, named)).toBeNull()
  })

  it('reads the cells of a bookmarked table', () => {
    expect(value('SUM(Other A1:B2)')).toBe(10)
    expect(value('SUM(Other B:B)')).toBe(6)
    expect(value('SUM(Other 2:2,1)')).toBe(8)
    expect(value('MAX(Other B2)')).toBe(4)
    expect(value('SUM(Other C3)')).toBeNull()
    expect(value('SUM(Other)')).toBe(3)
    expect(value('SUM(Rate A1:B2)')).toBeNull()
  })
})

describe('number pictures', () => {
  it('formats with digit places, grouping and literals', () => {
    expect(formatNumberPicture(1234.5, '#,##0')).toBe('1,235')
    expect(formatNumberPicture(1234.5, '#,##0.00')).toBe('1,234.50')
    expect(formatNumberPicture(0.5, '0.00')).toBe('0.50')
    expect(formatNumberPicture(7, '0')).toBe('7')
    expect(formatNumberPicture(45, '0%')).toBe('45%')
    expect(formatNumberPicture(45.678, '0.00%')).toBe('45.68%')
    expect(formatNumberPicture(3.5, '0.##')).toBe('3.5')
    expect(formatNumberPicture(-2, '0.0')).toBe('-2.0')
  })

  it('picks the section for the sign', () => {
    const money = '$#,##0.00;($#,##0.00)'
    expect(formatNumberPicture(1234.5, money)).toBe('$1,234.50')
    expect(formatNumberPicture(-1234.5, money)).toBe('($1,234.50)')
  })

  it('shows up to two decimals without a picture', () => {
    expect(formatFormulaResult(30, null)).toBe('30')
    expect(formatFormulaResult(1 / 3, null)).toBe('0.33')
    expect(formatFormulaResult(-12.5, null)).toBe('-12.5')
  })
})

describe('a formula field in a table', () => {
  const sheet = () =>
    makeEditor([
      tableOf([
        ['Item', 'Amount'],
        ['A', '10'],
        ['B', '20'],
        ['Total', ''],
      ]),
      para('after'),
    ])

  const caretInLastCell = (editor: Editor) => {
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docTableCell') at = pos + 2
    })
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)))
  }

  it('proposes a sum of the numbers above', () => {
    const editor = sheet()
    caretInLastCell(editor)
    expect(suggestedFormula(editor.state)).toBe('=SUM(ABOVE)')
  })

  it('inserts the field with its result', () => {
    const editor = sheet()
    caretInLastCell(editor)
    expect(insertTableFormula(editor, '=SUM(ABOVE) \\# "0.00"')).toBeNull()
    expect(fieldsOf(editor)).toEqual([
      { instr: '=SUM(ABOVE) \\# "0.00"', text: '30.00', dirty: false },
    ])
    expect(linesOf(editor)[3]).toBe('Total|30.00')
  })

  it('refuses a formula it cannot evaluate, and one that reads cells from outside a table', () => {
    const editor = sheet()
    caretInLastCell(editor)
    expect(insertTableFormula(editor, '=SUM(NoSuchBookmark)')).toBe('unsupported')
    caretIn(editor, 'after')
    expect(insertTableFormula(editor, '=SUM(ABOVE)')).toBe('outsideTable')
    expect(fieldsOf(editor)).toEqual([])
  })

  it('follows the cells it reads as they are edited', () => {
    const editor = sheet()
    caretInLastCell(editor)
    insertTableFormula(editor, '=SUM(ABOVE)')
    caretIn(editor, '10', 2)
    editor.commands.insertContent('5')
    expect(linesOf(editor).slice(0, 4)).toEqual(['Item|Amount', 'A|105', 'B|20', 'Total|125'])
    editor.commands.undo()
    expect(linesOf(editor).slice(0, 4)).toEqual(['Item|Amount', 'A|10', 'B|20', 'Total|30'])
  })

  it('settles a formula that reads another formula', () => {
    const editor = makeEditor([tableOf([['1', '2', '', '']]), para('')])
    const cells: number[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docTableCell') cells.push(pos + 2)
    })
    const put = (at: number, instr: string) => {
      editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)))
      expect(insertTableFormula(editor, instr)).toBeNull()
    }
    put(cells[3], '=C1*2')
    put(cells[2], '=SUM(LEFT)')
    expect(linesOf(editor)[0]).toBe('1|2|3|6')
  })

  it('recomputes a stale result on Update Field, and keeps one it cannot evaluate', () => {
    const field = (instr: string, text: string) => ({
      type: 'docParagraph',
      content: [{ type: 'text', text, marks: [{ type: 'instrField', attrs: { instr } }] }],
    })
    const editor = makeEditor([
      {
        type: 'docTable',
        content: [
          { type: 'docTableRow', content: [{ type: 'docTableCell', content: [para('4')] }] },
          { type: 'docTableRow', content: [{ type: 'docTableCell', content: [para('6')] }] },
          {
            type: 'docTableRow',
            content: [{ type: 'docTableCell', content: [field('=SUM(ABOVE)', '99')] }],
          },
          {
            type: 'docTableRow',
            content: [{ type: 'docTableCell', content: [field('=SUM(NoSuchBookmark)', '1')] }],
          },
        ],
      },
      para(''),
    ])
    // opening a document recomputes nothing
    expect(fieldsOf(editor).map((f) => f.text)).toEqual(['99', '1'])
    const tr = editor.state.tr
    expect(refreshTableFormulas(tr)).toBe(1)
    editor.view.dispatch(tr)
    expect(fieldsOf(editor).map((f) => f.text)).toEqual(['10', '1'])
  })
})

describe('a formula field in body text', () => {
  const marked = (text: string, name: string) => ({
    type: 'docParagraph',
    attrs: { bookmarks: [name] },
    content: [{ type: 'text', text }],
  })

  it('evaluates numbers, arithmetic, functions and bookmarks, and nothing that names a cell', () => {
    const value = (expr: string) => evaluateFormula(expr, NO_TABLE, { row: -1, col: -1 })
    expect(value('(2+3)*4')).toBe(20)
    expect(value('SUM(1,2,3)')).toBe(6)
    expect(value('IF(2>1,10,20)')).toBe(10)
    expect(value('ROUND(2.567,1)')).toBe(2.6)
    for (const expr of ['SUM(ABOVE)', 'SUM(LEFT)', 'A1', 'SUM(A1:B2)', 'SUM(1:2)', 'A1+1'])
      expect(value(expr), expr).toBeNull()
  })

  it('goes in at the caret with its result, and takes the place of selected text', () => {
    const editor = makeEditor([para('Total: '), para('Net 50 gross')])
    caretIn(editor, 'Total: ', 7)
    expect(formulaPlace(editor.state)).toBe(true)
    expect(suggestedFormula(editor.state)).toBe('=')
    expect(insertTableFormula(editor, '=(2+3)*4 \\# "0.00"')).toBeNull()
    expect(fieldsOf(editor)).toEqual([
      { instr: '=(2+3)*4 \\# "0.00"', text: '20.00', dirty: false },
    ])
    expect(linesOf(editor)).toEqual(['Total: 20.00', 'Net 50 gross'])

    // the caret in the field: the dialog's formula replaces it instead of adding a second one
    caretIn(editor, 'Total: 20.00', 9)
    expect(formulaFieldAt(editor.state)?.instr).toBe('=(2+3)*4 \\# "0.00"')
    expect(insertTableFormula(editor, '=MAX(3,9)')).toBeNull()
    expect(linesOf(editor)[0]).toBe('Total: 9')

    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.isTextblock && node.textContent === 'Net 50 gross') at = pos + 1
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at + 4, at + 6)),
    )
    expect(insertTableFormula(editor, '=7*6')).toBeNull()
    expect(linesOf(editor)[1]).toBe('Net 42 gross')
    editor.commands.undo()
    expect(linesOf(editor)[1]).toBe('Net 50 gross')
  })

  it('takes the place of a selection that runs across paragraphs', () => {
    const editor = makeEditor([para('one two'), para('three four'), para('tail')])
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 5, 15)),
    )
    expect(formulaPlace(editor.state)).toBe(true)
    expect(insertTableFormula(editor, '=6*7')).toBeNull()
    expect(linesOf(editor)).toEqual(['one 42 four', 'tail'])
    expect(fieldsOf(editor)).toEqual([{ instr: '=6*7', text: '42', dirty: false }])
    editor.commands.undo()
    expect(linesOf(editor)).toEqual(['one two', 'three four', 'tail'])
  })

  it('reads a bookmark, and follows it on Update Field only', () => {
    const editor = makeEditor([marked('Rate 20', 'Rate'), para('Result: ')])
    caretIn(editor, 'Result: ', 8)
    expect(insertTableFormula(editor, '=Rate*2')).toBeNull()
    expect(linesOf(editor)[1]).toBe('Result: 40')
    caretIn(editor, 'Rate 20', 7)
    editor.commands.insertContent('5')
    expect(linesOf(editor)[1]).toBe('Result: 40')
    const tr = editor.state.tr
    expect(refreshTableFormulas(tr)).toBe(1)
    editor.view.dispatch(tr)
    expect(linesOf(editor)[1]).toBe('Result: 410')
  })

  it('keeps the result of a body formula it cannot evaluate, and has no place on a selected block', () => {
    const field = (instr: string, text: string) => ({
      type: 'docParagraph',
      content: [{ type: 'text', text, marks: [{ type: 'instrField', attrs: { instr } }] }],
    })
    const editor = makeEditor([field('=SUM(ABOVE)', '12'), field('=2*8', '0'), tableOf([['1']])])
    const tr = editor.state.tr
    expect(refreshTableFormulas(tr)).toBe(1)
    editor.view.dispatch(tr)
    expect(fieldsOf(editor).map((f) => f.text)).toEqual(['12', '16'])

    let tablePos = -1
    editor.state.doc.forEach((node, pos) => {
      if (node.type.name === 'docTable') tablePos = pos
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, tablePos)),
    )
    expect(formulaPlace(editor.state)).toBe(false)
    expect(insertTableFormula(editor, '=1+1')).toBe('notInCell')
  })
})

describe('formulas with bookmarks in a document', () => {
  const marked = (text: string, name: string) => ({
    type: 'docParagraph',
    attrs: { bookmarks: [name] },
    content: [{ type: 'text', text }],
  })
  const lastCell = (editor: Editor) => {
    let at = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docTableCell') at = pos + 2
    })
    editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, at)))
  }

  it('reads a bookmarked paragraph and a bookmarked table', () => {
    const prices = {
      type: 'docTable',
      content: [
        {
          type: 'docTableRow',
          content: [
            { type: 'docTableCell', content: [marked('5', 'Prices')] },
            { type: 'docTableCell', content: [para('7')] },
          ],
        },
      ],
    }
    const editor = makeEditor([
      marked('Tax 20 percent', 'Tax'),
      prices,
      para('between'),
      tableOf([
        ['Net', '50'],
        ['Gross', ''],
      ]),
      para(''),
    ])
    const context = bookmarkContext(editor.state.doc)
    expect(context.bookmark('TAX')).toBe('Tax 20 percent')
    expect(context.bookmark('NONE')).toBeNull()
    expect(context.table('TAX')).toBeNull()
    expect(context.table('PRICES')?.cols).toBe(2)

    lastCell(editor)
    expect(
      insertTableFormula(editor, '=IF(Tax>0,B1*(100+Tax)/100,B1)+SUM(Prices A1:B1)'),
    ).toBeNull()
    expect(linesOf(editor)[4]).toBe('Gross|72')

    // Update Field picks up a change outside the table
    caretIn(editor, 'Tax 20 percent', 5)
    editor.commands.insertContent('1')
    expect(linesOf(editor)[4]).toBe('Gross|72')
    const tr = editor.state.tr
    expect(refreshTableFormulas(tr)).toBe(1)
    editor.view.dispatch(tr)
    expect(linesOf(editor)[4]).toBe('Gross|167')
  })
})

describe('a bookmark read from the file', () => {
  const R = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`
  const body =
    `<w:p>${R('Price 100, discount ')}<w:bookmarkStart w:id="1" w:name="Discount"/>${R('25')}` +
    `<w:bookmarkEnd w:id="1"/>${R(' each')}</w:p>` +
    `<w:p><w:bookmarkStart w:id="2" w:name="Point"/><w:bookmarkEnd w:id="2"/>${R('Rate 8 of 10')}</w:p>` +
    `<w:p><w:bookmarkStart w:id="3" w:name="Wide"/>${R('7 first')}</w:p>` +
    `<w:p>${R('9 second')}<w:bookmarkEnd w:id="3"/></w:p><w:p/>`

  async function opened() {
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    return { editor, parsed }
  }

  it('stands for the text between its start and end, not for its paragraph', async () => {
    const { editor, parsed } = await opened()
    const context = bookmarkContext(editor.state.doc, parsed.blocks)
    expect(context.bookmark('DISCOUNT')).toBe('25')
    expect(evaluateFormula('Discount*2', NO_TABLE, { row: -1, col: -1 }, context)).toBe(50)
    // without the file the paragraph's first number is all there is
    expect(bookmarkContext(editor.state.doc, []).bookmark('DISCOUNT')).toBe(
      'Price 100, discount 25 each',
    )
    // set at a point, or running on into the next paragraph: its paragraph
    expect(context.bookmark('POINT')).toBe('Rate 8 of 10')
    expect(context.bookmark('WIDE')).toBe('7 first')
    editor.destroy()
  })

  it('falls back to its paragraph once the paragraph no longer reads as the file has it', async () => {
    const { editor, parsed } = await opened()
    caretIn(editor, 'Price 100, discount 25 each', 6)
    editor.commands.insertContent('9')
    const context = bookmarkContext(editor.state.doc, parsed.blocks)
    expect(context.bookmark('DISCOUNT')).toBe('Price 9100, discount 25 each')
    editor.destroy()
  })
})

describe('Convert Text to Table', () => {
  const selectAll = (editor: Editor) => editor.commands.selectAll()

  it('counts the columns and rows the separator makes', () => {
    const texts = ['a\tb\tc', 'd\te']
    expect(suggestedColumns(texts, { kind: 'text', text: '\t' })).toBe(3)
    expect(suggestedColumns(texts, { kind: 'paragraph' })).toBe(1)
    expect(convertedRows(texts, { kind: 'text', text: '\t' }, 3)).toBe(2)
    expect(convertedRows(texts, { kind: 'text', text: '\t' }, 2)).toBe(3)
    expect(convertedRows(texts, { kind: 'paragraph' }, 2)).toBe(1)
  })

  it('makes a row of each paragraph, cut at the separator', () => {
    const editor = makeEditor([para('a\tb\tc'), para('d\te'), para('tail')])
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 12)),
    )
    expect(textToTableScope(editor.state)).toEqual(['a\tb\tc', 'd\te'])
    expect(
      convertTextToTable(
        editor,
        { cols: 3, separator: { kind: 'text', text: '\t' }, autoFit: 'window' },
        600,
      ),
    ).toBe(true)
    expect(linesOf(editor)).toEqual(['a|b|c', 'd|e|', 'tail'])
    editor.commands.undo()
    expect(linesOf(editor)).toEqual(['a\tb\tc', 'd\te', 'tail'])
  })

  it('keeps character formatting in the cells', () => {
    const editor = makeEditor([
      {
        type: 'docParagraph',
        content: [
          { type: 'text', text: 'plain,' },
          { type: 'text', text: 'bold', marks: [{ type: 'bold' }] },
        ],
      },
      para(''),
    ])
    caretIn(editor, 'plain,bold')
    convertTextToTable(
      editor,
      { cols: 2, separator: { kind: 'text', text: ',' }, autoFit: 'window' },
      600,
    )
    expect(linesOf(editor)[0]).toBe('plain|bold')
    let bold = ''
    editor.state.doc.descendants((node) => {
      if (node.isText && node.marks.some((m) => m.type.name === 'bold')) bold += node.text
    })
    expect(bold).toBe('bold')
  })

  it('fills the cells row by row from paragraphs', () => {
    const editor = makeEditor([para('a'), para('b'), para('c')])
    selectAll(editor)
    convertTextToTable(
      editor,
      { cols: 2, separator: { kind: 'paragraph' }, autoFit: 'window' },
      600,
    )
    expect(linesOf(editor).slice(0, 2)).toEqual(['a|b', 'c|'])
  })

  it('treats an empty Other character as paragraph marks, as the row count does', () => {
    const editor = makeEditor([para('a'), para('b'), para('c')])
    selectAll(editor)
    const separator = { kind: 'text', text: '' } as const
    expect(convertedRows(['a', 'b', 'c'], separator, 2)).toBe(2)
    convertTextToTable(editor, { cols: 2, separator, autoFit: 'window' }, 600)
    expect(linesOf(editor).slice(0, 2)).toEqual(['a|b', 'c|'])
  })

  it('applies the AutoFit choice in the same undo step', () => {
    const editor = makeEditor([para('a,b'), para('tail')])
    caretIn(editor, 'a,b')
    convertTextToTable(
      editor,
      { cols: 2, separator: { kind: 'text', text: ',' }, autoFit: 'contents' },
      600,
    )
    expect(editor.state.doc.firstChild?.attrs.tblAutoFit).toBe('contents')
    editor.commands.undo()
    expect(linesOf(editor)).toEqual(['a,b', 'tail'])
  })

  it('has nothing to convert inside a table', () => {
    const editor = makeEditor([tableOf([['x']]), para('')])
    caretIn(editor, 'x')
    expect(textToTableScope(editor.state)).toBeNull()
  })
})

describe('Convert Table to Text', () => {
  const twoByTwo = () => {
    const editor = makeEditor([
      para('before'),
      tableOf([
        ['a', 'b'],
        ['c', 'd'],
      ]),
      para('after'),
    ])
    caretIn(editor, 'a')
    return editor
  }

  it('joins the cells of a row with the separator', () => {
    const editor = twoByTwo()
    expect(
      convertTableToText(editor, { separator: { kind: 'text', text: '\t' }, nested: true }),
    ).toBe(true)
    expect(linesOf(editor)).toEqual(['before', 'a\tb', 'c\td', 'after'])
    editor.commands.undo()
    expect(linesOf(editor)).toEqual(['before', 'a|b', 'c|d', 'after'])
  })

  it('makes a paragraph of each cell with paragraph marks', () => {
    const editor = twoByTwo()
    convertTableToText(editor, { separator: { kind: 'paragraph' }, nested: true })
    expect(linesOf(editor)).toEqual(['before', 'a', 'b', 'c', 'd', 'after'])
  })

  it('round-trips through Convert Text to Table', () => {
    const editor = twoByTwo()
    convertTableToText(editor, { separator: { kind: 'text', text: ';' }, nested: true })
    expect(linesOf(editor)).toEqual(['before', 'a;b', 'c;d', 'after'])
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 9, 15)),
    )
    convertTextToTable(
      editor,
      { cols: 2, separator: { kind: 'text', text: ';' }, autoFit: 'window' },
      600,
    )
    expect(linesOf(editor)).toEqual(['before', 'a|b', 'c|d', 'after'])
  })

  it('turns a nested table into text, or leaves it a table', () => {
    const nested = {
      rows: [
        [
          { paras: ['n1'], richParas: [{ runs: [{ text: 'n1' }] }] },
          { paras: ['n2'], richParas: [{ runs: [{ text: 'n2' }] }] },
        ],
      ],
      colWidthsPct: [50, 50],
    }
    const make = () => {
      const editor = makeEditor([
        {
          type: 'docTable',
          content: [
            {
              type: 'docTableRow',
              content: [
                {
                  type: 'docTableCell',
                  content: [para('outer'), { type: 'docNestedTable', attrs: { model: nested } }],
                },
                { type: 'docTableCell', content: [para('z')] },
              ],
            },
          ],
        },
        para('after'),
      ])
      caretIn(editor, 'outer')
      return editor
    }
    const flat = make()
    convertTableToText(flat, { separator: { kind: 'text', text: ',' }, nested: true })
    expect(linesOf(flat)).toEqual(['outer', 'n1,n2,z', 'after'])
    const kept = make()
    convertTableToText(kept, { separator: { kind: 'text', text: ',' }, nested: false })
    expect(linesOf(kept)).toEqual(['outer', 'n1|n2', 'z', 'after'])
  })
})
